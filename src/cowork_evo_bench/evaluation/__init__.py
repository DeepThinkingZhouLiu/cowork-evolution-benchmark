"""标准单题结果、聚合指标与 H0/Candidate 配对比较。"""

from cowork_evo_bench.evaluation.metrics import aggregate_results, compare_paired
from cowork_evo_bench.evaluation.models import TaskResult, TaskStatus

__all__ = ["TaskResult", "TaskStatus", "aggregate_results", "compare_paired"]
