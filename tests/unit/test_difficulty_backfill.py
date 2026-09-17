"""难度回填不得丢失来源说明，也不得把不完整评测当成有效Panel。"""

import csv
import json
import sys
import tempfile
import tomllib
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
from sync_difficulty_labels import SOURCES, metadata_updates, source_rows


class DifficultyBackfillTest(unittest.TestCase):
    def test_preserves_distinct_notes_and_unrelated_fields(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            task = root / "tasks/docx/example"
            (task / "metadata").mkdir(parents=True)
            (task / "task.toml").write_text(
                '[metadata]\ndifficulty_level = "unassigned"\n'
                'difficulty_tier = "unassigned"\n'
                'notes = "no_vlm=true; difficulty=unassigned_pending_panel"\n'
                '[environment]\nallow_internet = true\n'
            )
            (task / "metadata/task.json").write_text(json.dumps({
                "difficulty_level": "unassigned", "difficulty_tier": "unassigned",
                "notes": '来源保留："原始合同"', "output_filenames": ["report.docx"],
            }))
            record = {"task_path": "tasks/docx/example",
                      "difficulty_level": "L3", "difficulty_tier": "hard"}
            updates = metadata_updates(root, record)
            config = tomllib.loads(updates[task / "task.toml"])
            mirror = json.loads(updates[task / "metadata/task.json"])
            self.assertIn("no_vlm=true", config["metadata"]["notes"])
            self.assertNotIn("unassigned", config["metadata"]["notes"])
            self.assertIn('来源保留："原始合同"', mirror["notes"])
            self.assertEqual(mirror["output_filenames"], ["report.docx"])
            self.assertTrue(config["environment"]["allow_internet"])
            for item in (config["metadata"], mirror):
                self.assertEqual(item["difficulty_tier"], "hard")
                self.assertIn("calibration_status=pilot_single_panel", item["notes"])
            for path, content in updates.items():
                path.write_text(content)
            self.assertEqual(metadata_updates(root, record), updates)

    def _single_cross_row(self, root, status):
        path = root / SOURCES["cross_office"]
        path.parent.mkdir()
        row = {"task": "example", "harness": "claude-code", "status": status,
               "visual_enabled": "false", "deterministic_reward": "0",
               "visual_reward": "", "combined_reward": "0"}
        with path.open("w", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=list(row))
            writer.writeheader()
            writer.writerow(row)

    def test_rejects_infrastructure_failure_instead_of_zero(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._single_cross_row(root, "infrastructure_failure")
            with self.assertRaisesRegex(ValueError, "基础设施失败"):
                source_rows(root, "cross_office")

    def test_rejects_incomplete_harness_panel(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._single_cross_row(root, "completed")
            with self.assertRaisesRegex(ValueError, "缺少三套"):
                source_rows(root, "cross_office")
