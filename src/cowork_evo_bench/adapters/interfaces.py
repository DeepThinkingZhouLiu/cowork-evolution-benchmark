"""Runner Adapter 接口，隔离具体 Agent 产品和 Harbor 版本。"""

from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from cowork_evo_bench.evaluation.models import TaskResult


@dataclass(frozen=True, slots=True)
class RunContext:
    """一次单题运行所需的固定配对条件。"""

    run_id: str
    benchmark_id: str
    partition: str
    attempt: int
    seed: int


class HarnessRunner(Protocol):
    """在给定任务和 Harness 上运行 Agent，并返回标准结果。"""

    def run(
        self,
        task_root: Path,
        harness_root: Path,
        context: RunContext,
    ) -> TaskResult:
        """运行环境并调用可信 Judge，不能接受 Agent 自报分数。"""
        ...
