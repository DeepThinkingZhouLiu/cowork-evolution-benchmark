"""所有 Evolution 算法必须实现的最小接口。"""

from typing import Protocol

from cowork_evo_bench.evolution.models import EvolutionRequest, MutationReport


class Evolver(Protocol):
    """读取反馈并只修改给定 Candidate 工作区。"""

    def evolve(self, request: EvolutionRequest) -> MutationReport:
        """生成 Harness 更新，不执行最终评测或晋升。"""
        ...
