"""聚合与严格配对指标的回归测试。"""

import unittest

from cowork_evo_bench.evaluation.metrics import compare_paired
from cowork_evo_bench.evaluation.models import TaskResult, TaskStatus


def _result(instance_id: str, reward: float, passed: bool) -> TaskResult:
    return TaskResult(
        instance_id=instance_id,
        attempt=1,
        seed=7,
        status=TaskStatus.COMPLETED,
        reward=reward,
        passed=passed,
    )


class MetricsTest(unittest.TestCase):
    def test_reports_new_success_and_regression(self) -> None:
        metrics = compare_paired(
            [_result("a", 0.2, False), _result("b", 0.8, True)],
            [_result("a", 0.9, True), _result("b", 0.4, False)],
        )
        self.assertEqual(metrics.newly_passed, 1)
        self.assertEqual(metrics.regressed, 1)
        self.assertEqual(metrics.net_passed, 0)
        self.assertAlmostEqual(metrics.mean_reward_delta, 0.15)

    def test_rejects_unpaired_results(self) -> None:
        with self.assertRaises(ValueError):
            compare_paired([_result("a", 0.2, False)], [_result("b", 0.9, True)])


if __name__ == "__main__":
    unittest.main()
