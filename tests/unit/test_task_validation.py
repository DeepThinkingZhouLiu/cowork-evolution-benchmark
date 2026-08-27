"""单题结构校验的最小回归测试。"""

import json
import tempfile
import unittest
from pathlib import Path

from cowork_evo_bench.tasks.validation import validate_task_package


class TaskValidationTest(unittest.TestCase):
    def test_accepts_minimal_complete_structure(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for relative in (
                "environment",
                "environment/input_files",
                "data/input_files",
                "solution",
                "tests",
                "metadata",
            ):
                (root / relative).mkdir(parents=True, exist_ok=True)

            (root / "instruction.md").write_text("生成一个办公交付物。", encoding="utf-8")
            (root / "task.toml").write_text(
                'artifacts = ["/app/output"]\n[task]\nname = "example/task-001"\n', encoding="utf-8"
            )
            criteria = [
                {
                    "id": f"R{index:03d}",
                    "type": "positive",
                    "weight": 1,
                }
                for index in range(1, 7)
            ]
            (root / "rubric.json").write_text(
                json.dumps({"pass_threshold": 0.7, "criteria": criteria}),
                encoding="utf-8",
            )
            (root / "environment/Dockerfile").write_text("FROM scratch\n", encoding="utf-8")
            (root / "solution/solve.sh").write_text("exit 0\n", encoding="utf-8")
            (root / "tests/test.sh").write_text("exit 0\n", encoding="utf-8")
            (root / "metadata/provenance.json").write_text("{}\n", encoding="utf-8")

            self.assertEqual(validate_task_package(root), [])

    def test_rejects_missing_files(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            issues = validate_task_package(Path(directory))
            self.assertTrue(any(issue.code == "task.file_missing" for issue in issues))


if __name__ == "__main__":
    unittest.main()
