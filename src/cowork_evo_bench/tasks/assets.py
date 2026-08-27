"""公开端到端任务资产的严格、无模型校验。"""

from __future__ import annotations

import hashlib
import json
import os
import tomllib
import zipfile
from pathlib import Path

from cowork_evo_bench.tasks.validation import validate_task_package
from cowork_evo_bench.validation import ValidationIssue

_OFFICE_SUFFIXES = {".docx", ".pptx", ".xlsx"}
_DELIVERABLE_SUFFIXES = _OFFICE_SUFFIXES | {".pdf", ".html"}


def _issue(code: str, path: Path, message: str) -> ValidationIssue:
    return ValidationIssue(code=code, location=str(path), message=message)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for block in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _read_json(path: Path, code: str) -> tuple[object | None, list[ValidationIssue]]:
    try:
        return json.loads(path.read_text(encoding="utf-8")), []
    except (OSError, json.JSONDecodeError) as error:
        return None, [_issue(code, path, f"无法读取 JSON：{error}")]


def _check_openable(path: Path) -> list[ValidationIssue]:
    suffix = path.suffix.lower()
    if suffix in _OFFICE_SUFFIXES:
        if not zipfile.is_zipfile(path):
            return [_issue("task.asset_corrupt", path, "Office 文件不是有效的 ZIP/OOXML 包。")]
        try:
            with zipfile.ZipFile(path) as archive:
                if "[Content_Types].xml" not in archive.namelist():
                    return [
                        _issue(
                            "task.asset_corrupt",
                            path,
                            "OOXML 文件缺少 [Content_Types].xml。",
                        )
                    ]
        except (OSError, zipfile.BadZipFile) as error:
            return [_issue("task.asset_corrupt", path, f"无法打开 Office 文件：{error}")]
    elif suffix == ".pdf":
        try:
            if path.read_bytes()[:5] != b"%PDF-":
                return [_issue("task.asset_corrupt", path, "PDF 文件头无效。")]
        except OSError as error:
            return [_issue("task.asset_corrupt", path, f"无法读取 PDF：{error}")]
    elif suffix == ".html":
        try:
            content = path.read_text(encoding="utf-8").lower()
            if "<html" not in content or "</html>" not in content:
                return [_issue("task.asset_corrupt", path, "HTML 根元素不完整。")]
        except (OSError, UnicodeDecodeError) as error:
            return [_issue("task.asset_corrupt", path, f"无法读取 HTML：{error}")]
    return []


def validate_task_assets(root: Path) -> list[ValidationIssue]:
    """在基础结构之上检查真实输入、Golden、哈希和隐藏侧入口。"""

    root = root.resolve()
    issues = list(validate_task_package(root))
    required_hidden = (
        "solution/solve.py",
        "tests/judge.py",
    )
    for relative in required_hidden:
        path = root / relative
        if not path.is_file():
            issues.append(_issue("task.asset_file_missing", path, "公开端到端样例缺少资产文件。"))

    input_root = root / "data/input_files"
    input_files = sorted(
        path
        for path in input_root.rglob("*")
        if path.is_file() and path.name not in {"README.md", ".gitkeep"}
    ) if input_root.is_dir() else []
    if not input_files:
        issues.append(_issue("task.input_empty", input_root, "公开端到端样例至少需要一个真实输入文件。"))
    for path in input_files:
        issues.extend(_check_openable(path))

    environment_input_root = root / "environment/input_files"
    environment_input_files = sorted(
        path
        for path in environment_input_root.rglob("*")
        if path.is_file() and path.name not in {"README.md", ".gitkeep"}
    ) if environment_input_root.is_dir() else []
    author_inputs = {
        path.relative_to(input_root): _sha256(path) for path in input_files
    }
    packaged_inputs = {
        path.relative_to(environment_input_root): _sha256(path)
        for path in environment_input_files
    }
    if author_inputs != packaged_inputs:
        issues.append(
            _issue(
                "task.environment_inputs_mismatch",
                environment_input_root,
                "environment/input_files 必须与 data/input_files 的路径和 SHA256 完全一致。",
            )
        )

    dockerfile = root / "environment/Dockerfile"
    if dockerfile.is_file():
        dockerfile_text = dockerfile.read_text(encoding="utf-8")
        if "COPY input_files" not in dockerfile_text:
            issues.append(
                _issue(
                    "task.environment_inputs_not_copied",
                    dockerfile,
                    "Dockerfile 必须把 environment/input_files 复制到 Agent 容器。",
                )
            )
        if "/app/output" not in dockerfile_text:
            issues.append(
                _issue(
                    "task.output_directory_missing",
                    dockerfile,
                    "Dockerfile 必须预先创建 /app/output，供 Agent 写入并由 Harbor 采集。",
                )
            )
        if "/data/input_files" not in dockerfile_text:
            issues.append(
                _issue(
                    "task.verifier_inputs_missing",
                    dockerfile,
                    "Dockerfile 必须让 Harbor Verifier 能从 /data/input_files 读取同一份输入。",
                )
            )

    solution_root = root / "solution"
    golden_root = solution_root / "Golden"
    if not golden_root.is_dir():
        golden_root = solution_root
    deliverables = sorted(
        path
        for path in golden_root.iterdir()
        if path.is_file() and path.suffix.lower() in _DELIVERABLE_SUFFIXES
    ) if golden_root.is_dir() else []
    if not deliverables:
        issues.append(_issue("task.golden_missing", golden_root, "公开端到端样例缺少 Golden 办公文件。"))
    for path in deliverables:
        issues.extend(_check_openable(path))

    for relative in ("solution/solve.sh", "tests/test.sh"):
        path = root / relative
        if path.is_file() and not os.access(path, os.X_OK):
            issues.append(_issue("task.script_not_executable", path, "任务入口脚本缺少可执行权限。"))

    provenance_path = root / "metadata/provenance.json"
    provenance, provenance_issues = _read_json(provenance_path, "task.provenance_invalid")
    issues.extend(provenance_issues)
    if isinstance(provenance, dict):
        declared = provenance.get("source_sha256")
        actual = sorted(_sha256(path) for path in input_files)
        if not isinstance(declared, list) or sorted(declared) != actual:
            issues.append(
                _issue(
                    "task.input_hash_mismatch",
                    provenance_path,
                    "source_sha256 必须与全部实际输入文件一一对应。",
                )
            )

    metadata_path = root / "metadata/task.json"
    output_names: list[str] = [path.name for path in deliverables]
    if metadata_path.is_file():
        metadata, metadata_issues = _read_json(metadata_path, "task.metadata_invalid")
        issues.extend(metadata_issues)
    else:
        metadata = None
    if isinstance(metadata, dict):
        raw_output_names = metadata.get("output_filenames")
        if not isinstance(raw_output_names, list) or not raw_output_names:
            issues.append(_issue("task.outputs_missing", metadata_path, "缺少 output_filenames。"))
        else:
            output_names = [name for name in raw_output_names if isinstance(name, str)]
            for name in output_names:
                golden_output = golden_root / name
                if not golden_output.is_file():
                    issues.append(
                        _issue(
                            "task.golden_output_missing",
                            golden_output,
                            "声明的输出没有对应 Golden。",
                        )
                    )
                input_copy = input_root / name
                if (
                    input_copy.is_file()
                    and golden_output.is_file()
                    and _sha256(input_copy) != _sha256(golden_output)
                ):
                    issues.append(
                        _issue(
                            "task.golden_leakage",
                            input_copy,
                            "输出同名文件出现在 Agent 可见输入目录，且与 Golden 内容不同。",
                        )
                    )

    expectations_path = root / "tests/expectations.json"
    if expectations_path.is_file():
        expectations, expectation_issues = _read_json(
            expectations_path, "task.expectations_invalid"
        )
        issues.extend(expectation_issues)
    else:
        expectations = None
    if isinstance(expectations, dict) and output_names:
        required_files = expectations.get("required_files")
        if not isinstance(required_files, list) or set(required_files) != set(output_names):
            issues.append(
                _issue(
                    "task.expectations_outputs_mismatch",
                    expectations_path,
                    "Judge 的 required_files 必须与任务声明输出一致。",
                )
            )

    task_toml = root / "task.toml"
    if task_toml.is_file():
        try:
            with task_toml.open("rb") as file:
                payload = tomllib.load(file)
            if payload.get("environment", {}).get("allow_internet") is not False:
                issues.append(
                    _issue(
                        "task.internet_not_locked",
                        task_toml,
                        "公开离线样例必须显式设置 allow_internet = false。",
                    )
                )
            if payload.get("environment", {}).get("network_mode") != "no-network":
                issues.append(
                    _issue(
                        "task.network_mode_not_locked",
                        task_toml,
                        '公开离线样例必须设置 network_mode = "no-network"。',
                    )
                )
            if payload.get("artifacts") != ["/app/output"]:
                issues.append(
                    _issue(
                        "task.artifacts_invalid",
                        task_toml,
                        "Harbor 只应采集 /app/output，避免重复采集 Judge 日志。",
                    )
                )
        except (OSError, tomllib.TOMLDecodeError):
            pass

    test_script = root / "tests/test.sh"
    judge_script = root / "tests/judge.py"
    if test_script.is_file() and "/logs/verifier/reward.txt" not in test_script.read_text(
        encoding="utf-8"
    ):
        issues.append(
            _issue(
                "task.reward_contract_missing",
                test_script,
                "Harbor Judge 入口必须写入 /logs/verifier/reward.txt。",
            )
        )
    if judge_script.is_file() and "--reward-file" not in judge_script.read_text(
        encoding="utf-8"
    ):
        issues.append(
            _issue(
                "task.reward_argument_missing",
                judge_script,
                "Judge 必须支持显式 --reward-file 输出参数。",
            )
        )
    return issues


def discover_task_roots(tasks_root: Path) -> tuple[Path, ...]:
    """发现领域目录下的正式题，跳过模板和说明目录。"""

    roots: list[Path] = []
    if not tasks_root.is_dir():
        return ()
    for instruction in tasks_root.glob("*/*/instruction.md"):
        root = instruction.parent
        if "templates" not in root.parts and "examples" not in root.parts:
            roots.append(root)
    return tuple(sorted(roots))
