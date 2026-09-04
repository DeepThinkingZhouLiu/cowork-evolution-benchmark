"""确定性分与视觉分的版本化聚合；基础设施失败永远不伪装成零分。"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from enum import StrEnum
from typing import Any

from cowork_evo_bench.judging.visual_config import RewardWeights


class CompositeStatus(StrEnum):
    COMPLETED = "completed"
    TASK_FAILURE = "task_failure"
    INFRASTRUCTURE_FAILURE = "infrastructure_failure"


@dataclass(frozen=True, slots=True)
class CompositeJudgeResult:
    schema_version: str
    task_id: str
    status: CompositeStatus
    visual_enabled: bool
    deterministic_reward: float
    visual_reward: float | None
    combined_reward: float | None
    weights: dict[str, float] | None
    failure: dict[str, str] | None = None

    def __post_init__(self) -> None:
        if self.schema_version != "1.0" or not self.task_id:
            raise ValueError("聚合结果的版本或任务编号无效。")
        if not 0 <= self.deterministic_reward <= 1:
            raise ValueError("deterministic_reward 必须位于 [0, 1]。")
        if self.status == CompositeStatus.INFRASTRUCTURE_FAILURE:
            if self.combined_reward is not None or self.visual_reward is not None:
                raise ValueError("基础设施失败不能产生视觉分或总分。")
            if self.failure is None:
                raise ValueError("基础设施失败必须包含机器可读原因。")
        else:
            if self.combined_reward is None or not 0 <= self.combined_reward <= 1:
                raise ValueError("完成或任务门槛失败必须有有效总分。")
        if self.visual_reward is not None and not 0 <= self.visual_reward <= 1:
            raise ValueError("visual_reward 必须位于 [0, 1]。")

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["status"] = self.status.value
        for key in ("visual_reward", "combined_reward", "weights", "failure"):
            if payload[key] is None:
                payload.pop(key)
        return payload


def combine_rewards(
    task_id: str,
    deterministic_reward: float,
    *,
    visual_enabled: bool,
    visual_reward: float | None = None,
    weights: RewardWeights | None = None,
    infrastructure_failure: dict[str, str] | None = None,
    task_gate_failure: dict[str, str] | None = None,
) -> CompositeJudgeResult:
    """按 v1 规则聚合：视觉题 70/30；非视觉题只用确定性分。"""

    if not 0 <= deterministic_reward <= 1:
        raise ValueError("deterministic_reward 必须位于 [0, 1]。")
    if infrastructure_failure is not None:
        return CompositeJudgeResult(
            schema_version="1.0",
            task_id=task_id,
            status=CompositeStatus.INFRASTRUCTURE_FAILURE,
            visual_enabled=visual_enabled,
            deterministic_reward=deterministic_reward,
            visual_reward=None,
            combined_reward=None,
            weights=None,
            failure=infrastructure_failure,
        )
    if task_gate_failure is not None:
        return CompositeJudgeResult(
            schema_version="1.0",
            task_id=task_id,
            status=CompositeStatus.TASK_FAILURE,
            visual_enabled=visual_enabled,
            deterministic_reward=deterministic_reward,
            visual_reward=None,
            combined_reward=0.0,
            weights=None,
            failure=task_gate_failure,
        )
    if not visual_enabled:
        return CompositeJudgeResult(
            schema_version="1.0",
            task_id=task_id,
            status=CompositeStatus.COMPLETED,
            visual_enabled=False,
            deterministic_reward=deterministic_reward,
            visual_reward=None,
            combined_reward=deterministic_reward,
            weights={"deterministic": 1.0, "visual": 0.0},
        )
    if visual_reward is None or weights is None:
        raise ValueError("启用视觉评分时必须提供 visual_reward 和 weights。")
    combined = weights.deterministic * deterministic_reward + weights.visual * visual_reward
    return CompositeJudgeResult(
        schema_version="1.0",
        task_id=task_id,
        status=CompositeStatus.COMPLETED,
        visual_enabled=True,
        deterministic_reward=deterministic_reward,
        visual_reward=visual_reward,
        combined_reward=combined,
        weights={"deterministic": weights.deterministic, "visual": weights.visual},
    )
