"""与具体 Agent 和 Judge 解耦的标准结果模型。"""

from dataclasses import dataclass
from enum import StrEnum


class TaskStatus(StrEnum):
    """任务运行状态，失败状态仍然保留在评测分母中。"""

    COMPLETED = "completed"
    ERROR = "error"
    TIMEOUT = "timeout"
    NOT_ATTEMPTED = "not_attempted"


@dataclass(frozen=True, slots=True)
class TaskResult:
    """一次 Harness 在一道题和一个随机种子上的标准结果。"""

    instance_id: str
    attempt: int
    seed: int
    status: TaskStatus
    reward: float
    passed: bool
    latency_ms: int = 0
    input_tokens: int | None = None
    output_tokens: int | None = None
    cost_usd: float | None = None

    def __post_init__(self) -> None:
        if not self.instance_id:
            raise ValueError("任务编号不能为空。")
        if self.attempt < 1:
            raise ValueError("运行次数必须从 1 开始。")
        if not 0 <= self.reward <= 1:
            raise ValueError("任务分数必须位于 [0, 1]。")
        if self.latency_ms < 0:
            raise ValueError("延迟不能小于零。")

    @property
    def pair_key(self) -> tuple[str, int, int]:
        """返回 H0 与 Candidate 配对所用的稳定键。"""

        return (self.instance_id, self.attempt, self.seed)
