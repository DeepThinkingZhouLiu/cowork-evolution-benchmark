"""Benchmark 级别的确定性聚合与配对指标。"""

from dataclasses import dataclass
from typing import Iterable

from cowork_evo_bench.evaluation.models import TaskResult


@dataclass(frozen=True, slots=True)
class AggregateMetrics:
    """一组标准单题结果的基础指标。"""

    count: int
    mean_reward: float
    success_rate: float
    total_latency_ms: int
    policy_errors: int


@dataclass(frozen=True, slots=True)
class PairedMetrics:
    """相同题目、次数和随机种子下的 H0/Candidate 差异。"""

    count: int
    baseline: AggregateMetrics
    candidate: AggregateMetrics
    mean_reward_delta: float
    success_rate_delta: float
    newly_passed: int
    regressed: int
    net_passed: int


def aggregate_results(results: Iterable[TaskResult]) -> AggregateMetrics:
    """把全部结果纳入分母，错误和超时不能被静默删除。"""

    rows = tuple(results)
    if not rows:
        raise ValueError("不能聚合空结果。")
    return AggregateMetrics(
        count=len(rows),
        mean_reward=sum(row.reward for row in rows) / len(rows),
        success_rate=sum(row.passed for row in rows) / len(rows),
        total_latency_ms=sum(row.latency_ms for row in rows),
        policy_errors=sum(row.status != "completed" for row in rows),
    )


def _index_results(results: Iterable[TaskResult], label: str) -> dict[tuple[str, int, int], TaskResult]:
    index: dict[tuple[str, int, int], TaskResult] = {}
    for result in results:
        if result.pair_key in index:
            raise ValueError(f"{label} 中存在重复配对键：{result.pair_key}")
        index[result.pair_key] = result
    return index


def compare_paired(
    baseline_results: Iterable[TaskResult],
    candidate_results: Iterable[TaskResult],
) -> PairedMetrics:
    """严格比较两组完全配对的结果，拒绝缺题或额外题目。"""

    baseline = _index_results(baseline_results, "H0")
    candidate = _index_results(candidate_results, "Candidate")
    if not baseline:
        raise ValueError("不能比较空结果。")
    if baseline.keys() != candidate.keys():
        missing = sorted(baseline.keys() - candidate.keys())
        extra = sorted(candidate.keys() - baseline.keys())
        raise ValueError(f"结果未完全配对，Candidate 缺少 {missing}，额外包含 {extra}。")

    keys = sorted(baseline)
    baseline_metrics = aggregate_results(baseline[key] for key in keys)
    candidate_metrics = aggregate_results(candidate[key] for key in keys)
    newly_passed = sum(not baseline[key].passed and candidate[key].passed for key in keys)
    regressed = sum(baseline[key].passed and not candidate[key].passed for key in keys)

    return PairedMetrics(
        count=len(keys),
        baseline=baseline_metrics,
        candidate=candidate_metrics,
        mean_reward_delta=candidate_metrics.mean_reward - baseline_metrics.mean_reward,
        success_rate_delta=candidate_metrics.success_rate - baseline_metrics.success_rate,
        newly_passed=newly_passed,
        regressed=regressed,
        net_passed=newly_passed - regressed,
    )
