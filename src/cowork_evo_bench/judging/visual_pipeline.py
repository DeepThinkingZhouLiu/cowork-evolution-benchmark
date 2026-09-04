"""单题 Codex-VLM 评分编排和 Harbor 兼容 artifacts 落盘。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from cowork_evo_bench.adapters.codex_vlm import CodexVLMJudge
from cowork_evo_bench.judging.composite import (
    CompositeJudgeResult,
    CompositeStatus,
    combine_rewards,
)
from cowork_evo_bench.judging.rendering import (
    RenderFailure,
    render_artifacts,
    validate_candidate_artifacts,
)
from cowork_evo_bench.judging.visual_config import VisualJudgeSettings
from cowork_evo_bench.judging.visual_models import (
    VisualJudgeResult,
    VisualJudgeStatus,
    VisualRubric,
)


class VisualPipelineError(RuntimeError):
    """调用方式或本地输入错误，不属于被测任务评分。"""


def load_deterministic_reward(path: Path) -> tuple[float, dict[str, Any]]:
    """接受现有 judge JSON，并保留原结果的可审计副本。"""

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise VisualPipelineError(f"无法读取确定性评分结果 {path}: {error}") from error
    if isinstance(payload, (int, float)) and not isinstance(payload, bool):
        reward = float(payload)
        normalized = {"reward": reward}
    elif isinstance(payload, dict):
        value = payload.get("deterministic_reward", payload.get("reward"))
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            raise VisualPipelineError("确定性结果必须包含 reward 或 deterministic_reward。")
        reward = float(value)
        normalized = payload
    else:
        raise VisualPipelineError("确定性结果必须是数值或 JSON 对象。")
    if not 0 <= reward <= 1:
        raise VisualPipelineError("确定性分数必须位于 [0, 1]。")
    return reward, normalized


def _write_json(path: Path, payload: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    temporary.replace(path)


def _write_final_outputs(output_root: Path, result: CompositeJudgeResult) -> None:
    _write_json(output_root / "combined" / "result.json", result.to_dict())
    reward_dir = output_root / "reward"
    reward_dir.mkdir(parents=True, exist_ok=True)
    _write_json(
        reward_dir / "status.json",
        {
            "schema_version": "1.0",
            "task_id": result.task_id,
            "status": result.status.value,
            "reward_available": result.combined_reward is not None,
        },
    )
    if result.combined_reward is not None:
        (reward_dir / "reward.txt").write_text(
            f"{result.combined_reward:.10f}\n", encoding="utf-8"
        )


def run_visual_judge_pipeline(
    *,
    task_dir: Path,
    candidate_dir: Path,
    deterministic_result_path: Path,
    visual_rubric: VisualRubric | None,
    settings: VisualJudgeSettings,
    api_key: str,
    output_root: Path,
) -> CompositeJudgeResult:
    """运行一题；失败语义明确区分任务门槛与评分基础设施。"""

    task_dir = task_dir.resolve()
    candidate_dir = candidate_dir.resolve()
    output_root = output_root.resolve()
    if output_root.exists():
        raise VisualPipelineError(f"输出目录已存在，拒绝覆盖历史评分：{output_root}")
    deterministic_reward, deterministic_payload = load_deterministic_reward(
        deterministic_result_path
    )
    output_root.mkdir(parents=True)
    deterministic_task_id = deterministic_payload.get("task_id")
    if not isinstance(deterministic_task_id, str) or not deterministic_task_id.strip():
        deterministic_task_id = task_dir.name
    task_id = (
        visual_rubric.task_id
        if visual_rubric is not None
        else deterministic_task_id
    )
    _write_json(
        output_root / "deterministic" / "result.json",
        {
            "schema_version": "1.0",
            "task_id": task_id,
            "deterministic_reward": deterministic_reward,
            "source_result": deterministic_payload,
        },
    )

    if visual_rubric is None or not visual_rubric.enabled:
        combined = combine_rewards(
            task_id,
            deterministic_reward,
            visual_enabled=False,
        )
        _write_final_outputs(output_root, combined)
        return combined

    try:
        validate_candidate_artifacts(visual_rubric.candidate_artifacts, candidate_dir)
    except RenderFailure as error:
        if not error.task_gate:
            raise
        combined = combine_rewards(
            task_id,
            deterministic_reward,
            visual_enabled=True,
            task_gate_failure={"code": error.code, "message": str(error)},
        )
        _write_final_outputs(output_root, combined)
        return combined

    try:
        candidate_bundle = render_artifacts(
            visual_rubric.candidate_artifacts,
            candidate_dir,
            output_root / "render" / "candidate",
            settings.rendering,
            role="candidate",
        )
        reference_bundle = None
        if visual_rubric.reference_mode.value == "reference":
            reference_bundle = render_artifacts(
                visual_rubric.reference_artifacts,
                task_dir,
                output_root / "render" / "reference",
                settings.rendering,
                role="reference",
            )
    except RenderFailure as error:
        if error.task_gate:
            combined = combine_rewards(
                task_id,
                deterministic_reward,
                visual_enabled=True,
                task_gate_failure={"code": error.code, "message": str(error)},
            )
            _write_final_outputs(output_root, combined)
            return combined
        visual_failure = VisualJudgeResult(
            schema_version="1.0",
            task_id=task_id,
            rubric_version=visual_rubric.rubric_version,
            status=VisualJudgeStatus.INFRASTRUCTURE_FAILURE,
            provider=settings.provider.name,
            model=settings.codex.model,
            codex_version="not_started",
            reference_mode=visual_rubric.reference_mode,
            input_image_sha256=(),
            criterion_results=(),
            visual_reward=None,
            summary="",
            runtime_ms=0,
            attempt_count=0,
            retry_count=0,
            error={"code": error.code, "message": str(error)},
        )
        _write_json(output_root / "visual" / "result.json", visual_failure.to_dict())
        combined = combine_rewards(
            task_id,
            deterministic_reward,
            visual_enabled=True,
            infrastructure_failure=visual_failure.error,
        )
        _write_final_outputs(output_root, combined)
        return combined

    judge = CodexVLMJudge(settings, api_key)
    instruction_path = task_dir / "instruction.md"
    try:
        instruction = instruction_path.read_text(encoding="utf-8")
    except OSError as error:
        raise VisualPipelineError(f"无法读取任务指令 {instruction_path}: {error}") from error
    visual_result = judge.evaluate(
        visual_rubric,
        instruction,
        candidate_bundle,
        reference_bundle,
        output_root / "visual" / "codex",
    )
    _write_json(output_root / "visual" / "result.json", visual_result.to_dict())
    if visual_result.status == VisualJudgeStatus.INFRASTRUCTURE_FAILURE:
        combined = combine_rewards(
            task_id,
            deterministic_reward,
            visual_enabled=True,
            infrastructure_failure=visual_result.error,
        )
    else:
        combined = combine_rewards(
            task_id,
            deterministic_reward,
            visual_enabled=True,
            visual_reward=visual_result.visual_reward,
            weights=settings.weights,
        )
    _write_final_outputs(output_root, combined)
    return combined
