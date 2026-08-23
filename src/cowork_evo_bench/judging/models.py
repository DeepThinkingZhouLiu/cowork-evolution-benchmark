"""Judge 输出模型，不包含具体 Office 文件实现。"""

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class CriterionResult:
    """一条 Rubric 的标准化评分与证据。"""

    criterion_id: str
    score: float
    evidence: str
    method: str

    def __post_init__(self) -> None:
        if not 0 <= self.score <= 1:
            raise ValueError("判据分数必须位于 [0, 1]。")


@dataclass(frozen=True, slots=True)
class JudgeResult:
    """一个任务的一次完整 Judge 结果。"""

    task_id: str
    rubric_version: str
    judge_version: str
    criterion_results: tuple[CriterionResult, ...]
    reward: float
    passed: bool
    hurdle_failures: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        if not 0 <= self.reward <= 1:
            raise ValueError("任务分数必须位于 [0, 1]。")
        if len(self.criterion_results) < 6:
            raise ValueError("Judge 结果至少需要六条判据。")
