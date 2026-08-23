"""项目命令行入口，只负责参数解析和调用领域模块。"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

from cowork_evo_bench.benchmarking.validation import validate_benchmark_file
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

    benchmark_parser = subparsers.add_parser(
        "validate-benchmark", help="校验 Benchmark Release 清单"
    )
    benchmark_parser.add_argument("path", type=Path, help="benchmark.json 路径")

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """执行命令行并返回进程退出码。"""

    args = build_parser().parse_args(argv)
    if args.command == "validate-task":
        return _print_issues(validate_task_package(args.path))
    if args.command == "validate-benchmark":
        return _print_issues(validate_benchmark_file(args.path))
    raise AssertionError(f"未处理的命令：{args.command}")
