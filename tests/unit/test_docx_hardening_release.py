"""Word 分区迁移和缺失分数的回归检查。"""

import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "tools"))
from freeze_docx_hardening_release import repair_assignments
from run_docx_hardening_panel import LABELS, summarize


class DocxHardeningTest(unittest.TestCase):
    def test_moves_minimum_and_keeps_external_partition(self):
        tasks = [SimpleNamespace(domain="docx", instance_id=f"docx/{i}") for i in range(60)]
        tasks.append(SimpleNamespace(domain="cross_office", instance_id="cross_office/a"))
        before = {t.instance_id: "feedback" if i < 30 else "selection" if i < 40 else "final"
                  for i, t in enumerate(tasks)}
        components = [[0, 60]] + [[i] for i in range(1, 60)]
        result, cost = repair_assignments(tasks, components, before)
        self.assertEqual(cost, 2)
        self.assertEqual(result["docx/0"], "final")
        self.assertEqual(result["cross_office/a"], "final")
        self.assertEqual(sum(result[f"docx/{i}"] == "feedback" for i in range(60)), 30)
        self.assertEqual(sum(result[f"docx/{i}"] == "selection" for i in range(60)), 10)

    def test_rejects_conflicting_external_partitions(self):
        tasks = [SimpleNamespace(domain="docx", instance_id="docx/a"),
                 SimpleNamespace(domain="xlsx", instance_id="xlsx/b"),
                 SimpleNamespace(domain="cross_office", instance_id="cross_office/c")]
        with self.assertRaises(ValueError):
            repair_assignments(tasks, [[0, 1, 2]], {"docx/a": "feedback", "xlsx/b": "feedback", "cross_office/c": "final"})

    def test_missing_trial_does_not_become_zero_score(self):
        frozen = SimpleNamespace(completed_job=lambda *args: None)
        baseline = {h: [{"task": "valid", "reward": 1.0}] for h in LABELS}
        changed = {h: [(1, Path("pending"))] for h in LABELS}
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)
            summarize(output, changed, baseline, frozen)
            panel = json.loads((output / "panel-summary.json").read_text())
            for row in panel["harnesses"]:
                self.assertEqual(row["valid_count"], 1)
                self.assertEqual(row["mean_reward_valid"], 1.0)
                self.assertIsNone(row["mean_reward_all_60"])
                self.assertEqual(row["pending_or_infrastructure_failure_tasks"], ["pending"])
