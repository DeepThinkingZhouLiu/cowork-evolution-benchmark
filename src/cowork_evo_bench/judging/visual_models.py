"""任务特定视觉 Rubric、模型响应和可信评分结果。"""

from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from enum import StrEnum
from pathlib import Path
from typing import Any


class VisualProtocolError(ValueError):
    """视觉 Rubric 或模型响应不符合协议。"""


class ReferenceMode(StrEnum):
    REFERENCE = "reference"
    REFERENCE_FREE = "reference_free"


class VisualJudgeStatus(StrEnum):
    COMPLETED = "completed"
    INFRASTRUCTURE_FAILURE = "infrastructure_failure"


@dataclass(frozen=True, slots=True)
class VisualArtifactSpec:
    path: str
    kind: str
    label: str = ""

    def __post_init__(self) -> None:
        candidate = Path(self.path)
        if candidate.is_absolute() or ".." in candidate.parts or not self.path:
            raise VisualProtocolError("视觉产物路径必须是根目录内的安全相对路径。")
        if self.kind not in {"pptx", "pdf"}:
            raise VisualProtocolError("视觉产物类型只能是 pptx 或 pdf。")
        if candidate.suffix.lower() != f".{self.kind}":
            raise VisualProtocolError("视觉产物扩展名必须与 kind 一致。")


@dataclass(frozen=True, slots=True)
class VisualCriterion:
    criterion_id: str
    description: str
    weight: float
    score_anchors: dict[str, str]
    evidence_requirement: str

    def __post_init__(self) -> None:
        if not re.fullmatch(r"V[0-9]{3}", self.criterion_id):
            raise VisualProtocolError("视觉判据编号必须采用 V001 形式。")
        if not self.description.strip() or not self.evidence_requirement.strip():
            raise VisualProtocolError("视觉判据描述和证据要求不能为空。")
        if isinstance(self.weight, bool) or self.weight <= 0:
            raise VisualProtocolError("视觉判据权重必须大于零。")
        if set(self.score_anchors) != {"0", "0.5", "1"}:
            raise VisualProtocolError("视觉判据必须给出 0、0.5、1 三个分数锚点。")
        if any(not isinstance(value, str) or not value.strip() for value in self.score_anchors.values()):
            raise VisualProtocolError("视觉判据分数锚点不能为空。")


@dataclass(frozen=True, slots=True)
class VisualRubric:
    task_id: str
    rubric_version: str
    enabled: bool
    reference_mode: ReferenceMode
    candidate_artifacts: tuple[VisualArtifactSpec, ...]
    reference_artifacts: tuple[VisualArtifactSpec, ...]
    criteria: tuple[VisualCriterion, ...]

    def __post_init__(self) -> None:
        if not self.task_id.strip() or not self.rubric_version.strip():
            raise VisualProtocolError("视觉 Rubric 的任务编号和版本不能为空。")
        if not self.candidate_artifacts:
            raise VisualProtocolError("视觉 Rubric 至少需要一个候选产物。")
        if self.reference_mode == ReferenceMode.REFERENCE and not self.reference_artifacts:
            raise VisualProtocolError("reference 模式必须声明参考产物。")
        if not self.criteria:
            raise VisualProtocolError("启用视觉评分时至少需要一条任务特定判据。")
        criterion_ids = [criterion.criterion_id for criterion in self.criteria]
        if len(criterion_ids) != len(set(criterion_ids)):
            raise VisualProtocolError("视觉判据编号不能重复。")


@dataclass(frozen=True, slots=True)
class VisualCriterionResult:
    criterion_id: str
    score: float
    evidence: str
    reason: str

    def __post_init__(self) -> None:
        if not re.fullmatch(r"V[0-9]{3}", self.criterion_id):
            raise VisualProtocolError("模型返回了非法视觉判据编号。")
        if isinstance(self.score, bool) or not 0 <= self.score <= 1:
            raise VisualProtocolError("视觉判据分数必须位于 [0, 1]。")
        if not self.evidence.strip() or not self.reason.strip():
            raise VisualProtocolError("视觉判据证据和原因不能为空。")


@dataclass(frozen=True, slots=True)
class ParsedModelResponse:
    criterion_results: tuple[VisualCriterionResult, ...]
    summary: str


@dataclass(frozen=True, slots=True)
class VisualJudgeResult:
    schema_version: str
    task_id: str
    rubric_version: str
    status: VisualJudgeStatus
    provider: str
    model: str
    codex_version: str
    reference_mode: ReferenceMode
    input_image_sha256: tuple[str, ...]
    criterion_results: tuple[VisualCriterionResult, ...]
    visual_reward: float | None
    summary: str
    runtime_ms: int
    attempt_count: int
    retry_count: int
    error: dict[str, str] | None = None

    def __post_init__(self) -> None:
        if self.schema_version != "1.0":
            raise VisualProtocolError("视觉结果 schema_version 必须为 1.0。")
        if self.runtime_ms < 0 or self.attempt_count < 0 or self.retry_count < 0:
            raise VisualProtocolError("视觉评分运行元数据不合法。")
        if len(self.input_image_sha256) != len(set(self.input_image_sha256)):
            raise VisualProtocolError("视觉输入图片哈希不能重复。")
        if any(not re.fullmatch(r"[a-f0-9]{64}", value) for value in self.input_image_sha256):
            raise VisualProtocolError("视觉输入图片哈希不合法。")
        if self.status == VisualJudgeStatus.COMPLETED:
            if self.visual_reward is None or not 0 <= self.visual_reward <= 1:
                raise VisualProtocolError("完成状态必须包含 [0, 1] 内的 visual_reward。")
            if not self.criterion_results or not self.summary.strip() or self.error is not None:
                raise VisualProtocolError("完成状态的结果字段不完整。")
        else:
            if self.visual_reward is not None or self.error is None:
                raise VisualProtocolError("基础设施失败不能产生视觉分，且必须带错误信息。")

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["status"] = self.status.value
        payload["reference_mode"] = self.reference_mode.value
        payload["input_image_sha256"] = list(self.input_image_sha256)
        payload["criterion_results"] = [
            asdict(result) for result in self.criterion_results
        ]
        if payload["visual_reward"] is None:
            payload.pop("visual_reward")
        if payload["error"] is None:
            payload.pop("error")
        return payload


def load_visual_rubric(path: Path) -> VisualRubric:
    """严格读取一题的视觉 Rubric，拒绝未知字段和通用空话。"""

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise VisualProtocolError(f"无法读取视觉 Rubric {path}: {error}") from error
    if not isinstance(payload, dict):
        raise VisualProtocolError("视觉 Rubric 根节点必须是对象。")
    allowed = {
        "task_id",
        "rubric_version",
        "enabled",
        "reference_mode",
        "candidate_artifacts",
        "reference_artifacts",
        "criteria",
    }
    unknown = set(payload) - allowed
    if unknown:
        raise VisualProtocolError(f"视觉 Rubric 包含未知字段：{sorted(unknown)}")
    try:
        mode = ReferenceMode(payload["reference_mode"])
        candidate_specs = tuple(
            _parse_artifact(item) for item in payload["candidate_artifacts"]
        )
        reference_specs = tuple(
            _parse_artifact(item) for item in payload.get("reference_artifacts", [])
        )
        criteria = tuple(_parse_criterion(item) for item in payload["criteria"])
        enabled = payload["enabled"]
    except (KeyError, TypeError, ValueError) as error:
        raise VisualProtocolError(f"视觉 Rubric 字段无效：{error}") from error
    if not isinstance(enabled, bool):
        raise VisualProtocolError("视觉 Rubric enabled 必须是布尔值。")
    return VisualRubric(
        task_id=str(payload["task_id"]),
        rubric_version=str(payload["rubric_version"]),
        enabled=enabled,
        reference_mode=mode,
        candidate_artifacts=candidate_specs,
        reference_artifacts=reference_specs,
        criteria=criteria,
    )


def _parse_artifact(payload: object) -> VisualArtifactSpec:
    if not isinstance(payload, dict) or set(payload) - {"path", "kind", "label"}:
        raise VisualProtocolError("视觉产物声明格式无效。")
    return VisualArtifactSpec(
        path=str(payload["path"]),
        kind=str(payload["kind"]),
        label=str(payload.get("label", "")),
    )


def _parse_criterion(payload: object) -> VisualCriterion:
    expected = {
        "criterion_id",
        "description",
        "weight",
        "score_anchors",
        "evidence_requirement",
    }
    if not isinstance(payload, dict) or set(payload) != expected:
        raise VisualProtocolError("每条视觉判据必须且只能包含协议规定字段。")
    anchors = payload["score_anchors"]
    if not isinstance(anchors, dict):
        raise VisualProtocolError("score_anchors 必须是对象。")
    return VisualCriterion(
        criterion_id=str(payload["criterion_id"]),
        description=str(payload["description"]),
        weight=float(payload["weight"]),
        score_anchors={str(key): str(value) for key, value in anchors.items()},
        evidence_requirement=str(payload["evidence_requirement"]),
    )


def parse_model_response(raw: str, rubric: VisualRubric) -> ParsedModelResponse:
    """验证 Codex 最终 JSON，并确保每条 Rubric 恰好评分一次。"""

    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as error:
        raise VisualProtocolError(f"模型最终响应不是合法 JSON：{error}") from error
    if not isinstance(payload, dict) or set(payload) != {"criterion_results", "summary"}:
        raise VisualProtocolError("模型响应字段必须严格匹配协议。")
    rows = payload["criterion_results"]
    if not isinstance(rows, list):
        raise VisualProtocolError("criterion_results 必须是数组。")
    results: list[VisualCriterionResult] = []
    for row in rows:
        if not isinstance(row, dict) or set(row) != {
            "criterion_id",
            "score",
            "evidence",
            "reason",
        }:
            raise VisualProtocolError("模型判据结果字段不符合协议。")
        results.append(
            VisualCriterionResult(
                criterion_id=str(row["criterion_id"]),
                score=float(row["score"]),
                evidence=str(row["evidence"]),
                reason=str(row["reason"]),
            )
        )
    expected_ids = [criterion.criterion_id for criterion in rubric.criteria]
    actual_ids = [result.criterion_id for result in results]
    if sorted(expected_ids) != sorted(actual_ids) or len(actual_ids) != len(set(actual_ids)):
        raise VisualProtocolError("模型必须对每条视觉判据恰好评分一次。")
    summary = payload["summary"]
    if not isinstance(summary, str) or not summary.strip():
        raise VisualProtocolError("模型响应 summary 不能为空。")
    by_id = {result.criterion_id: result for result in results}
    return ParsedModelResponse(
        criterion_results=tuple(by_id[criterion_id] for criterion_id in expected_ids),
        summary=summary.strip(),
    )


def compute_visual_reward(
    rubric: VisualRubric,
    results: tuple[VisualCriterionResult, ...],
) -> float:
    """只信任逐项分，由本地代码重新计算加权视觉分。"""

    scores = {result.criterion_id: result.score for result in results}
    expected = {criterion.criterion_id for criterion in rubric.criteria}
    if set(scores) != expected:
        raise VisualProtocolError("视觉结果与 Rubric 判据不完全对应。")
    total_weight = sum(criterion.weight for criterion in rubric.criteria)
    return sum(
        criterion.weight * scores[criterion.criterion_id]
        for criterion in rubric.criteria
    ) / total_weight
