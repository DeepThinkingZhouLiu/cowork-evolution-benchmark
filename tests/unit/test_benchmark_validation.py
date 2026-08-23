"""Benchmark 三段式划分的回归测试。"""

import unittest

from cowork_evo_bench.benchmarking.validation import validate_benchmark_payload


def _payload() -> dict:
    return {
        "spec": {
            "partitions": {
                "feedback": {
                    "visibility": "detailed",
                    "expected_count": 1,
                    "instance_ids": ["ppt/f-1"],
                },
                "selection": {
                    "visibility": "aggregate_only",
                    "expected_count": 1,
                    "instance_ids": ["ppt/s-1"],
                },
                "final": {
                    "visibility": "sealed",
                    "expected_count": 1,
                    "instance_ids": ["ppt/t-1"],
                },
            },
            "expected_total": 3,
        }
    }


class BenchmarkValidationTest(unittest.TestCase):
    def test_accepts_disjoint_partitions(self) -> None:
        self.assertEqual(validate_benchmark_payload(_payload()), [])

    def test_rejects_cross_partition_duplicate(self) -> None:
        payload = _payload()
        payload["spec"]["partitions"]["final"]["instance_ids"] = ["ppt/f-1"]
        issues = validate_benchmark_payload(payload)
        self.assertTrue(
            any(issue.code == "benchmark.cross_partition_leakage" for issue in issues)
        )


if __name__ == "__main__":
    unittest.main()
