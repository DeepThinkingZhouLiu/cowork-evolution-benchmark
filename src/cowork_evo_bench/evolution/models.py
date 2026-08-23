"""Evolution 算法与可信 Controller 之间的数据边界。"""

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class EvolutionRequest:
    """一次 Updater Session 可以读取的全部输入。"""

    run_id: str
    epoch_id: int
    candidate_root: Path
    feedback_packet: Path
    mutation_policy: Path


@dataclass(frozen=True, slots=True)
class MutationReport:
    """Evolution 算法完成修改后提交的可审计报告。"""

    summary: str
    hypothesis: str
    changed_paths: tuple[str, ...]
    evidence_refs: tuple[str, ...]
