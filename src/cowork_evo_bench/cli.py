"""项目命令行入口，只负责参数解析和调用领域模块。"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

from cowork_evo_bench.benchmarking.validation import validate_benchmark_file
from cowork_evo_bench.judging.composite import CompositeStatus
from cowork_evo_bench.judging.visual_config import (
    load_env_file,
    load_visual_judge_settings,
    resolve_api_key,
)
from cowork_evo_bench.judging.visual_models import load_visual_rubric
from cowork_evo_bench.judging.visual_pipeline import run_visual_judge_pipeline
from cowork_evo_bench.tasks.assets import discover_task_roots, validate_task_assets
from cowork_evo_bench.tasks.validation import validate_task_package
from cowork_evo_bench.validation import ValidationIssue


def _print_issues(issues: Sequence[ValidationIssue]) -> int:
    if not issues:
        print("校验通过。")
        return 0

    for issue in issues:
        print(f"[{issue.code}] {issue.location}: {issue.message}")
    print(f"共发现 {len(issues)} 个问题。")
    return 1


def build_parser() -> argparse.ArgumentParser:
    """构造命令行参数，不读取文件或执行任务。"""

    parser = argparse.ArgumentParser(description="CoworkEvoBench 协议校验工具")
    subparsers = parser.add_subparsers(dest="command", required=True)

    task_parser = subparsers.add_parser("validate-task", help="校验 Harbor 单题目录")
    task_parser.add_argument("path", type=Path, help="单题目录")

    asset_parser = subparsers.add_parser(
        "validate-task-assets", help="严格校验一个公开端到端任务的输入、Golden 和哈希"
    )
    asset_parser.add_argument("path", type=Path, help="单题目录")

    public_parser = subparsers.add_parser(
        "validate-public-tasks", help="严格校验 tasks 下发现的全部公开端到端任务"
    )
    public_parser.add_argument("path", type=Path, help="tasks 根目录")

    benchmark_parser = subparsers.add_parser(
        "validate-benchmark", help="校验 Benchmark Release 清单"
    )
    benchmark_parser.add_argument("path", type=Path, help="benchmark.json 路径")

    visual_parser = subparsers.add_parser(
        "judge-visual", help="运行单题确定性 + Codex-VLM 混合评分"
    )
    visual_parser.add_argument("--task-dir", type=Path, required=True, help="单题根目录")
    visual_parser.add_argument(
        "--candidate-dir", type=Path, required=True, help="受测 Harness 的交付目录"
    )
    visual_parser.add_argument(
        "--deterministic-result",
        type=Path,
        required=True,
        help="现有代码 Judge 的 JSON 结果",
    )
    visual_parser.add_argument(
        "--config",
        type=Path,
        default=Path("configs/judging/codex-vlm.json"),
        help="公共 Codex-VLM 非秘密配置",
    )
    visual_parser.add_argument(
        "--env-file",
        type=Path,
        default=Path(".env"),
        help="本地秘密文件；必须被 Git 忽略",
    )
    visual_parser.add_argument(
        "--visual-rubric",
        type=Path,
        help="任务特定视觉 Rubric；省略时尝试 <task-dir>/visual_rubric.json",
    )
    visual_parser.add_argument(
        "--output-dir", type=Path, required=True, help="新的评分 artifacts 目录"
    )

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """执行命令行并返回进程退出码。"""

    args = build_parser().parse_args(argv)
    if args.command == "validate-task":
        return _print_issues(validate_task_package(args.path))
    if args.command == "validate-task-assets":
        return _print_issues(validate_task_assets(args.path))
    if args.command == "validate-public-tasks":
        roots = discover_task_roots(args.path)
        if not roots:
            print("没有发现公开端到端任务。")
            return 1
        issues = [issue for root in roots for issue in validate_task_assets(root)]
        if not issues:
            print(f"校验通过：{len(roots)} 个公开端到端任务。")
            return 0
        return _print_issues(issues)
    if args.command == "validate-benchmark":
        return _print_issues(validate_benchmark_file(args.path))
    if args.command == "judge-visual":
        settings = load_visual_judge_settings(args.config)
        rubric_path = args.visual_rubric or (args.task_dir / "visual_rubric.json")
        rubric = load_visual_rubric(rubric_path) if rubric_path.is_file() else None
        api_key = ""
        if rubric is not None and rubric.enabled:
            api_key = resolve_api_key(settings, load_env_file(args.env_file))
        result = run_visual_judge_pipeline(
            task_dir=args.task_dir,
            candidate_dir=args.candidate_dir,
            deterministic_result_path=args.deterministic_result,
            visual_rubric=rubric,
            settings=settings,
            api_key=api_key,
            output_root=args.output_dir,
        )
        if result.combined_reward is None:
            print(f"评分基础设施失败；详见 {args.output_dir / 'combined/result.json'}")
            return 2
        print(f"评分完成：{result.combined_reward:.6f}")
        return 0 if result.status != CompositeStatus.INFRASTRUCTURE_FAILURE else 2
    raise AssertionError(f"未处理的命令：{args.command}")


if __name__ == "__main__":
    raise SystemExit(main())
