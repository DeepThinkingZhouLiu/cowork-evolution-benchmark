"""单题视觉编排的 reference 隔离、门槛和失败语义测试。"""

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from cowork_evo_bench.judging.rendering import RenderBundle
from cowork_evo_bench.judging.visual_config import load_visual_judge_settings
from cowork_evo_bench.judging.visual_models import (
    ReferenceMode,
    VisualArtifactSpec,
    VisualCriterion,
    VisualCriterionResult,
    VisualJudgeResult,
    VisualJudgeStatus,
    VisualRubric,
)
from cowork_evo_bench.judging.visual_pipeline import run_visual_judge_pipeline


def _rubric(mode: ReferenceMode = ReferenceMode.REFERENCE_FREE) -> VisualRubric:
    return VisualRubric(
        task_id="ppt/example",
        rubric_version="visual-v1",
        enabled=True,
        reference_mode=mode,
        candidate_artifacts=(VisualArtifactSpec("answer.pptx", "pptx"),),
        # 故意声明不存在的隐藏参考；reference_free 下绝不能读取它。
        reference_artifacts=(VisualArtifactSpec("solution/must-not-read.pptx", "pptx"),),
        criteria=(
            VisualCriterion(
                "V001",
                "标题层级清楚。",
                1,
                {"0": "不清楚", "0.5": "部分清楚", "1": "清楚"},
                "指出页码和标题区域。",
            ),
        ),
    )


class VisualPipelineTest(unittest.TestCase):
    def _paths(self, root: Path):
        task = root / "task"
        candidate = root / "candidate"
        task.mkdir()
        candidate.mkdir()
        (task / "instruction.md").write_text("制作一页演示文稿。", encoding="utf-8")
        deterministic = root / "deterministic.json"
        deterministic.write_text(
            json.dumps({"task_id": "cowork-evo/example", "reward": 0.8}),
            encoding="utf-8",
        )
        config = Path(__file__).parents[2] / "configs/judging/codex-vlm.json"
        settings = load_visual_judge_settings(config, {})
        return task, candidate, deterministic, settings

    def test_reference_free_never_renders_or_reads_reference(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            task, candidate, deterministic, settings = self._paths(root)
            visual_result = VisualJudgeResult(
                schema_version="1.0",
                task_id="ppt/example",
                rubric_version="visual-v1",
                status=VisualJudgeStatus.COMPLETED,
                provider="zcloud",
                model="gpt-5.6-terra",
                codex_version="0.148.0",
                reference_mode=ReferenceMode.REFERENCE_FREE,
                input_image_sha256=(),
                criterion_results=(VisualCriterionResult("V001", 0.5, "第1页", "层级一般"),),
                visual_reward=0.5,
                summary="一般",
                runtime_ms=1,
                attempt_count=1,
                retry_count=0,
            )
            fake_judge = Mock()
            fake_judge.evaluate.return_value = visual_result
            with (
                patch("cowork_evo_bench.judging.visual_pipeline.validate_candidate_artifacts"),
                patch(
                    "cowork_evo_bench.judging.visual_pipeline.render_artifacts",
                    return_value=RenderBundle("candidate", (), root / "manifest.json"),
                ) as render,
                patch(
                    "cowork_evo_bench.judging.visual_pipeline.CodexVLMJudge",
                    return_value=fake_judge,
                ),
            ):
                result = run_visual_judge_pipeline(
                    task_dir=task,
                    candidate_dir=candidate,
                    deterministic_result_path=deterministic,
                    visual_rubric=_rubric(),
                    settings=settings,
                    api_key="local-only",
                    output_root=root / "output",
                )
            self.assertEqual(render.call_count, 1)
            self.assertEqual(render.call_args.kwargs["role"], "candidate")
            self.assertIsNone(fake_judge.evaluate.call_args.args[3])
            self.assertAlmostEqual(result.combined_reward, 0.71)
            self.assertFalse((root / "output/render/reference").exists())

    def test_reference_mode_renders_reference_and_passes_it_to_judge(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            task, candidate, deterministic, settings = self._paths(root)
            visual_result = VisualJudgeResult(
                schema_version="1.0",
                task_id="ppt/example",
                rubric_version="visual-v1",
                status=VisualJudgeStatus.COMPLETED,
                provider="zcloud",
                model="gpt-5.6-terra",
                codex_version="0.148.0",
                reference_mode=ReferenceMode.REFERENCE,
                input_image_sha256=(),
                criterion_results=(VisualCriterionResult("V001", 1, "第1页", "一致"),),
                visual_reward=1,
                summary="一致",
                runtime_ms=1,
                attempt_count=1,
                retry_count=0,
            )
            fake_judge = Mock()
            fake_judge.evaluate.return_value = visual_result

            def fake_render(*args, **kwargs):
                return RenderBundle(kwargs["role"], (), root / f"{kwargs['role']}-manifest.json")

            with (
                patch("cowork_evo_bench.judging.visual_pipeline.validate_candidate_artifacts"),
                patch(
                    "cowork_evo_bench.judging.visual_pipeline.render_artifacts",
                    side_effect=fake_render,
                ) as render,
                patch(
                    "cowork_evo_bench.judging.visual_pipeline.CodexVLMJudge",
                    return_value=fake_judge,
                ),
            ):
                run_visual_judge_pipeline(
                    task_dir=task,
                    candidate_dir=candidate,
                    deterministic_result_path=deterministic,
                    visual_rubric=_rubric(ReferenceMode.REFERENCE),
                    settings=settings,
                    api_key="local-only",
                    output_root=root / "output",
                )
            self.assertEqual(
                [call.kwargs["role"] for call in render.call_args_list],
                ["candidate", "reference"],
            )
            self.assertIsNotNone(fake_judge.evaluate.call_args.args[3])

    def test_no_visual_task_keeps_deterministic_reward_without_api(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            task, candidate, deterministic, settings = self._paths(root)
            result = run_visual_judge_pipeline(
                task_dir=task,
                candidate_dir=candidate,
                deterministic_result_path=deterministic,
                visual_rubric=None,
                settings=settings,
                api_key="",
                output_root=root / "output",
            )
            self.assertFalse(result.visual_enabled)
            self.assertEqual(result.task_id, "cowork-evo/example")
            self.assertEqual(result.combined_reward, 0.8)

    def test_missing_core_candidate_is_task_zero_not_infrastructure_failure(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            task, candidate, deterministic, settings = self._paths(root)
            result = run_visual_judge_pipeline(
                task_dir=task,
                candidate_dir=candidate,
                deterministic_result_path=deterministic,
                visual_rubric=_rubric(),
                settings=settings,
                api_key="unused",
                output_root=root / "output",
            )
            self.assertEqual(result.status.value, "task_failure")
            self.assertEqual(result.combined_reward, 0)
            self.assertEqual((root / "output/reward/reward.txt").read_text().strip(), "0.0000000000")

    def test_damaged_core_candidate_is_task_zero(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            task, candidate, deterministic, settings = self._paths(root)
            (candidate / "answer.pptx").write_bytes(b"not-an-ooxml-file")
            result = run_visual_judge_pipeline(
                task_dir=task,
                candidate_dir=candidate,
                deterministic_result_path=deterministic,
                visual_rubric=_rubric(),
                settings=settings,
                api_key="unused",
                output_root=root / "output",
            )
            self.assertEqual(result.status.value, "task_failure")
            self.assertEqual(result.combined_reward, 0)

    def test_vlm_failure_has_no_final_reward_file(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            task, candidate, deterministic, settings = self._paths(root)
            visual_failure = VisualJudgeResult(
                schema_version="1.0",
                task_id="ppt/example",
                rubric_version="visual-v1",
                status=VisualJudgeStatus.INFRASTRUCTURE_FAILURE,
                provider="zcloud",
                model="gpt-5.6-terra",
                codex_version="0.148.0",
                reference_mode=ReferenceMode.REFERENCE_FREE,
                input_image_sha256=(),
                criterion_results=(),
                visual_reward=None,
                summary="",
                runtime_ms=1,
                attempt_count=2,
                retry_count=1,
                error={"code": "api_timeout", "message": "超时"},
            )
            fake_judge = Mock()
            fake_judge.evaluate.return_value = visual_failure
            with (
                patch("cowork_evo_bench.judging.visual_pipeline.validate_candidate_artifacts"),
                patch(
                    "cowork_evo_bench.judging.visual_pipeline.render_artifacts",
                    return_value=RenderBundle("candidate", (), root / "manifest.json"),
                ),
                patch(
                    "cowork_evo_bench.judging.visual_pipeline.CodexVLMJudge",
                    return_value=fake_judge,
                ),
            ):
                result = run_visual_judge_pipeline(
                    task_dir=task,
                    candidate_dir=candidate,
                    deterministic_result_path=deterministic,
                    visual_rubric=_rubric(),
                    settings=settings,
                    api_key="local-only",
                    output_root=root / "output",
                )
            self.assertEqual(result.status.value, "infrastructure_failure")
            self.assertIsNone(result.combined_reward)
            self.assertFalse((root / "output/reward/reward.txt").exists())


if __name__ == "__main__":
    unittest.main()
