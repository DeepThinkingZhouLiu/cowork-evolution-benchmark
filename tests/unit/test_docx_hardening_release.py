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
from docx_panel_results import budget_timeout_record, difficulty_report, recover_source


class DocxHardeningTest(unittest.TestCase):
    def test_timeout_requires_real_sdk_usage_evidence(self):
        record = {"reward": 0, "infrastructure_failure": True,
                  "exception": {"exception_type": "AgentTimeoutError"}}
        with tempfile.TemporaryDirectory() as directory:
            job = Path(directory)
            self.assertTrue(budget_timeout_record(job, record)["infrastructure_failure"])
            log = job / "trial/agent/openhands_sdk.txt"
            log.parent.mkdir(parents=True)
            log.write_text("Agent Action\nTokens: ↑ input 334.99K (total 3.06M)\n")
            result = budget_timeout_record(job, record)
            self.assertFalse(result["infrastructure_failure"])
            self.assertEqual(result["reward"], 0)

    def test_recover_old_source_does_not_guess_between_retries(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            row = {"task": "sample", "source_result_path": None}
            for attempt in (1, 2):
                path = root / f"trials/01-sample-a0{attempt}/trial/result.json"
                path.parent.mkdir(parents=True)
                path.write_text(json.dumps({"finished_at": "done", "verifier_result": {"rewards": {"reward": 0}}}))
            with self.assertRaises(ValueError):
                recover_source(row, root)

    def test_difficulty_has_four_quotas_and_discloses_ties(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            output = root / "results"
            for i in range(60):
                task = root / "tasks/docx" / f"task-{i:02}"
                task.mkdir(parents=True)
                (task / "task.toml").write_text("[metadata]\nestimated_minutes = 10\n")
            for harness in LABELS:
                path = output / harness / "results.json"
                path.parent.mkdir(parents=True)
                path.write_text(json.dumps({"valid_count": 60, "tasks": [
                    {"task": f"task-{i:02}", "reward": 0.5} for i in range(60)]}))
            difficulty_report(root, output, LABELS)
            data = json.loads((output / "difficulty.json").read_text())
            self.assertEqual([r["count"] for r in data["tiers"]], [12, 18, 18, 12])
            self.assertEqual(len(data["boundary_ties"]), 3)
            self.assertTrue(all(r["empirical_difficulty"] == 0.5 for r in data["tasks"]))

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
