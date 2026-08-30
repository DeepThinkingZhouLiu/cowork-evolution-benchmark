"""Harbor Office 单题的本地结构校验。"""

from __future__ import annotations

import json
import tomllib
from pathlib import Path

from cowork_evo_bench.validation import ValidationIssue

_REQUIRED_FILES = (
    "instruction.md",
    "task.toml",
    "rubric.json",
    "environment/Dockerfile",
    "solution/solve.sh",
    "tests/test.sh",
    "metadata/provenance.json",
)

_REQUIRED_DIRECTORIES = (
    "data/input_files",
    "environment",
    "solution",
    "tests",
    "metadata",
)

_DIFFICULTY_TIER_BY_LEVEL = {
    "unassigned": "unassigned",
    "L1": "easy",
    "L2": "medium",
    "L3": "hard",
    "L4": "expert",
}


def _issue(code: str, path: Path, message: str) -> ValidationIssue:
    return ValidationIssue(code=code, location=str(path), message=message)


def validate_difficulty_metadata(
    metadata: object,
    path: Path,
    code_prefix: str = "task.difficulty",
) -> list[ValidationIssue]:
    """校验 L1-L4 与英文难度标签的一一映射。"""

    if not isinstance(metadata, dict):
        return [_issue(f"{code_prefix}_missing", path, "缺少难度元数据。")]

    level = metadata.get("difficulty_level")
    tier = metadata.get("difficulty_tier")
    expected_tier = _DIFFICULTY_TIER_BY_LEVEL.get(level)
    if expected_tier is None:
        return [
            _issue(
                f"{code_prefix}_level_invalid",
                path,
                "difficulty_level 必须是 unassigned、L1、L2、L3 或 L4。",
            )
        ]
    if tier != expected_tier:
        return [
            _issue(
                f"{code_prefix}_pair_invalid",
                path,
                f"{level} 必须与 difficulty_tier={expected_tier!r} 配对。",
            )
        ]
    return []


def _validate_task_toml(path: Path) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    try:
        with path.open("rb") as file:
            payload = tomllib.load(file)
    except (OSError, tomllib.TOMLDecodeError) as error:
        return [_issue("task.toml.invalid", path, f"无法解析 TOML：{error}")]

    task = payload.get("task")
    if not isinstance(task, dict):
        issues.append(_issue("task.toml.task_missing", path, "缺少 [task] 配置段。"))
        return issues

    name = task.get("name")
    if not isinstance(name, str) or name.count("/") != 1 or any(
        not part.strip() for part in name.split("/")
    ):
        issues.append(
            _issue(
                "task.toml.name_invalid",
                path,
                "task.name 必须采用 <命名空间>/<任务编号> 两段格式。",
            )
        )
    metadata = payload.get("metadata")
    issues.extend(validate_difficulty_metadata(metadata, path))
    return issues


def _validate_rubric(path: Path) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        return [_issue("rubric.invalid", path, f"无法解析 JSON：{error}")]

    threshold = payload.get("pass_threshold")
    if not isinstance(threshold, (int, float)) or isinstance(threshold, bool):
        issues.append(_issue("rubric.threshold_missing", path, "缺少数值型成功阈值。"))
    elif not 0 <= float(threshold) <= 1:
        issues.append(_issue("rubric.threshold_range", path, "成功阈值必须位于 [0, 1]。"))

    criteria = payload.get("criteria")
    if not isinstance(criteria, list) or len(criteria) < 6:
        issues.append(_issue("rubric.criteria_count", path, "Rubric 至少需要六条判据。"))
        return issues

    ids: list[str] = []
    for index, criterion in enumerate(criteria):
        location = f"{path}#criteria[{index}]"
        if not isinstance(criterion, dict):
            issues.append(
                ValidationIssue("rubric.criterion_invalid", location, "判据必须是对象。")
            )
            continue
        criterion_id = criterion.get("id")
        if not isinstance(criterion_id, str) or not criterion_id:
            issues.append(
                ValidationIssue("rubric.id_missing", location, "判据缺少稳定编号。")
            )
        else:
            ids.append(criterion_id)

        criterion_type = criterion.get("type")
        weight = criterion.get("weight")
        if not isinstance(weight, (int, float)) or isinstance(weight, bool):
            issues.append(
                ValidationIssue("rubric.weight_invalid", location, "判据权重必须是数值。")
            )
        elif criterion_type == "penalty" and weight >= 0:
            issues.append(
                ValidationIssue("rubric.penalty_weight", location, "惩罚项权重必须小于零。")
            )
        elif criterion_type in {"positive", "hurdle"} and weight <= 0:
            issues.append(
                ValidationIssue("rubric.positive_weight", location, "正向或门槛项权重必须大于零。")
            )

    if len(ids) != len(set(ids)):
        issues.append(_issue("rubric.duplicate_ids", path, "判据编号不能重复。"))
    return issues


def validate_task_package(root: Path) -> list[ValidationIssue]:
    """检查单题的必要结构与基础语义，不启动环境或 Judge。"""

    root = root.resolve()
    issues: list[ValidationIssue] = []
    if not root.is_dir():
        return [_issue("task.root_missing", root, "单题目录不存在。")]

    for relative in _REQUIRED_DIRECTORIES:
        path = root / relative
        if not path.is_dir():
            issues.append(_issue("task.directory_missing", path, "缺少必需目录。"))

    for relative in _REQUIRED_FILES:
        path = root / relative
        if not path.is_file():
            issues.append(_issue("task.file_missing", path, "缺少必需文件。"))

    instruction = root / "instruction.md"
    if instruction.is_file() and not instruction.read_text(encoding="utf-8").strip():
        issues.append(_issue("task.instruction_empty", instruction, "任务指令不能为空。"))

    task_toml = root / "task.toml"
    if task_toml.is_file():
        issues.extend(_validate_task_toml(task_toml))

    rubric = root / "rubric.json"
    if rubric.is_file():
        issues.extend(_validate_rubric(rubric))
    return issues
