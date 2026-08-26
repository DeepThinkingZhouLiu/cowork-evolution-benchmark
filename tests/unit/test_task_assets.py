"""公开端到端任务资产校验的回归测试。"""

import json
import tempfile
import unittest
import zipfile
from hashlib import sha256
from pathlib import Path

from cowork_evo_bench.tasks.assets import validate_task_assets


class TaskAssetValidationTest(unittest.TestCase):
    def _build_task(self, root: Path) -> None:
        for relative in ("environment", "data/input_files", "solution", "tests", "metadata"):
            (root / relative).mkdir(parents=True, exist_ok=True)
        (root / "instruction.md").write_text("生成真实办公文件。\n", encoding="utf-8")
        (root / "task.toml").write_text(
            'schema_version = "1.4"\n[task]\nname = "cowork-evo/sample"\n'
            '[environment]\nallow_internet = false\n',
            encoding="utf-8",
        )
        criteria = [
            {"id": f"R{index:03d}", "type": "positive", "weight": 1}
            for index in range(1, 7)
        ]
        (root / "rubric.json").write_text(
            json.dumps({"pass_threshold": 0.7, "criteria": criteria}), encoding="utf-8"
        )
        (root / "environment/Dockerfile").write_text("FROM scratch\n", encoding="utf-8")
        input_path = root / "data/input_files/source.csv"
        input_path.write_text("name,value\nsample,1\n", encoding="utf-8")
        golden = root / "solution/result.xlsx"
        with zipfile.ZipFile(golden, "w") as archive:
            archive.writestr("[Content_Types].xml", "<Types />")
        for relative in ("solution/solve.sh", "solution/solve.py", "tests/test.sh", "tests/judge.py"):
            path = root / relative
            path.write_text("#!/usr/bin/env python3\n", encoding="utf-8")
            path.chmod(0o755)
        (root / "tests/expectations.json").write_text(
            json.dumps({"required_files": ["result.xlsx"]}), encoding="utf-8"
        )
        (root / "metadata/task.json").write_text(
            json.dumps({"output_filenames": ["result.xlsx"]}), encoding="utf-8"
        )
        (root / "metadata/provenance.json").write_text(
            json.dumps({"source_sha256": [sha256(input_path.read_bytes()).hexdigest()]}),
            encoding="utf-8",
        )

    def test_accepts_complete_public_task(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._build_task(root)
            self.assertEqual(validate_task_assets(root), [])

    def test_rejects_stale_input_hash(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._build_task(root)
            (root / "data/input_files/source.csv").write_text("changed\n", encoding="utf-8")
            issues = validate_task_assets(root)
            self.assertTrue(any(issue.code == "task.input_hash_mismatch" for issue in issues))

    def test_accepts_nested_golden_and_unchanged_companion(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._build_task(root)
            nested = root / "solution/Golden"
            nested.mkdir()
            (root / "solution/result.xlsx").replace(nested / "result.xlsx")
            companion = root / "data/input_files/companion.xlsx"
            companion.write_bytes((nested / "result.xlsx").read_bytes())
            (nested / "companion.xlsx").write_bytes(companion.read_bytes())
            provenance = json.loads((root / "metadata/provenance.json").read_text())
            provenance["source_sha256"].append(sha256(companion.read_bytes()).hexdigest())
            (root / "metadata/provenance.json").write_text(json.dumps(provenance))
            metadata = json.loads((root / "metadata/task.json").read_text())
            metadata["output_filenames"].append("companion.xlsx")
            (root / "metadata/task.json").write_text(json.dumps(metadata))
            expectations = json.loads((root / "tests/expectations.json").read_text())
            expectations["required_files"].append("companion.xlsx")
            (root / "tests/expectations.json").write_text(json.dumps(expectations))

            self.assertEqual(validate_task_assets(root), [])

    def test_rejects_changed_same_name_companion(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._build_task(root)
            nested = root / "solution/Golden"
            nested.mkdir()
            (root / "solution/result.xlsx").replace(nested / "result.xlsx")
            companion = root / "data/input_files/companion.xlsx"
            companion.write_bytes((nested / "result.xlsx").read_bytes())
            with zipfile.ZipFile(nested / "companion.xlsx", "w") as archive:
                archive.writestr("[Content_Types].xml", "<Types />")
                archive.writestr("changed.txt", "changed")
            provenance = json.loads((root / "metadata/provenance.json").read_text())
            provenance["source_sha256"].append(sha256(companion.read_bytes()).hexdigest())
            (root / "metadata/provenance.json").write_text(json.dumps(provenance))
            metadata = json.loads((root / "metadata/task.json").read_text())
            metadata["output_filenames"].append("companion.xlsx")
            (root / "metadata/task.json").write_text(json.dumps(metadata))
            expectations = json.loads((root / "tests/expectations.json").read_text())
            expectations["required_files"].append("companion.xlsx")
            (root / "tests/expectations.json").write_text(json.dumps(expectations))

            issues = validate_task_assets(root)
            self.assertTrue(any(issue.code == "task.golden_leakage" for issue in issues))


if __name__ == "__main__":
    unittest.main()
