"""具体确定性、视觉和语义 Judge 需要实现的最小接口。"""

from pathlib import Path
from typing import Protocol

from cowork_evo_bench.judging.models import JudgeResult


class TaskJudge(Protocol):
    """把一个任务的交付目录转换成可信 Judge 结果。"""

    def evaluate(self, task_root: Path, artifact_root: Path) -> JudgeResult:
        """读取隐藏评测配置和产物，但不修改 Agent 工作区。"""
        ...
