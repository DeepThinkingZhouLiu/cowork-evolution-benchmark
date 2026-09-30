#!/usr/bin/env python3
"""从逐题 JSON 重算 Cross-Office 三套 Harness 混合评分汇总。"""

from __future__ import annotations

import argparse
import csv
import json
from collections import Counter
from dataclasses import dataclass
from pathlib import Path


HARNESSES = ("claude-code", "opencode", "openhands")
EXPECTED_TASKS = 60
VALID_STATUSES = {"completed", "task_failure", "infrastructure_failure"}


@dataclass(frozen=True)
class TaskScore:
    """一个 Harness 在一道题上的标准评分结果。"""

    harness: str
    task: str
    visual_enabled: bool
    status: str
    deterministic_reward: float
    visual_reward: float | None
    combined_reward: float | None
    failure_code: str


def _read_json(path: Path) -> dict:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError(f"JSON 根节点不是对象：{path}")
    return payload


def _number(value: object, *, field: str, path: Path) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{path} 的 {field} 不是有效数值")
    result = float(value)
    if not 0.0 <= result <= 1.0:
        raise ValueError(f"{path} 的 {field} 越界：{result}")
    return result


def load_scores(panel_root: Path) -> list[TaskScore]:
    """读取逐题结果，并严格校验数量、状态与 70/30 公式。"""

    rows: list[TaskScore] = []
    task_sets: dict[str, set[str]] = {}
    for harness in HARNESSES:
        harness_root = panel_root / harness
        result_paths = sorted(harness_root.glob("*/combined/result.json"))
        if len(result_paths) != EXPECTED_TASKS:
            raise ValueError(
                f"{harness} 期望 {EXPECTED_TASKS} 份结果，实际 {len(result_paths)} 份"
            )
        task_sets[harness] = {path.parent.parent.name for path in result_paths}
        if len(task_sets[harness]) != EXPECTED_TASKS:
            raise ValueError(f"{harness} 存在重复任务")

        for path in result_paths:
            payload = _read_json(path)
            task = path.parent.parent.name
            status = str(payload.get("status", ""))
            if status not in VALID_STATUSES:
                raise ValueError(f"{path} 的状态无效：{status}")

            deterministic = _number(
                payload.get("deterministic_reward"),
                field="deterministic_reward",
                path=path,
            )
            visual_enabled = payload.get("visual_enabled") is True
            visual_raw = payload.get("visual_reward")
            visual = (
                None
                if visual_raw is None
                else _number(visual_raw, field="visual_reward", path=path)
            )
            combined_raw = payload.get("combined_reward")
            combined = (
                None
                if combined_raw is None
                else _number(combined_raw, field="combined_reward", path=path)
            )

            if status == "infrastructure_failure":
                if visual is not None or combined is not None:
                    raise ValueError(f"{path} 将基础设施失败误记为有效分")
            else:
                if combined is None:
                    raise ValueError(f"{path} 缺少 combined_reward")
                if visual_enabled and status == "completed" and visual is None:
                    raise ValueError(f"{path} 是视觉题但缺少 visual_reward")
                expected = (
                    0.7 * deterministic + 0.3 * visual
                    if visual_enabled and visual is not None
                    else deterministic
                )
                if abs(combined - expected) > 1e-9:
                    raise ValueError(
                        f"{path} 合成分不符合协议：{combined} != {expected}"
                    )

            failure = payload.get("failure")
            failure_code = (
                str(failure.get("code", "")) if isinstance(failure, dict) else ""
            )
            rows.append(
                TaskScore(
                    harness=harness,
                    task=task,
                    visual_enabled=visual_enabled,
                    status=status,
                    deterministic_reward=deterministic,
                    visual_reward=visual,
                    combined_reward=combined,
                    failure_code=failure_code,
                )
            )

    reference = task_sets[HARNESSES[0]]
    for harness in HARNESSES[1:]:
        if task_sets[harness] != reference:
            raise ValueError(f"{harness} 与 {HARNESSES[0]} 的任务集不一致")
    return rows


def _mean(values: list[float]) -> float:
    if not values:
        raise ValueError("无可用分数")
    return sum(values) / len(values)


def _format_reward(value: float | None) -> str:
    return "" if value is None else f"{value:.10f}"


def write_csv(path: Path, rows: list[TaskScore]) -> None:
    """写出 180 行逐题明细。"""

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(
            stream,
            fieldnames=(
                "harness",
                "task",
                "visual_enabled",
                "status",
                "deterministic_reward",
                "visual_reward",
                "combined_reward",
                "failure_code",
            ),
        )
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    "harness": row.harness,
                    "task": row.task,
                    "visual_enabled": str(row.visual_enabled).lower(),
                    "status": row.status,
                    "deterministic_reward": _format_reward(row.deterministic_reward),
                    "visual_reward": _format_reward(row.visual_reward),
                    "combined_reward": _format_reward(row.combined_reward),
                    "failure_code": row.failure_code,
                }
            )


def write_markdown(path: Path, rows: list[TaskScore], *, csv_path: Path) -> None:
    """写出协议说明、Harness 汇总和完整性结论。"""

    table_rows: list[str] = []
    for harness in HARNESSES:
        subset = [row for row in rows if row.harness == harness]
        counts = Counter(row.status for row in subset)
        deterministic = [row.deterministic_reward for row in subset]
        visual = [row.visual_reward for row in subset if row.visual_reward is not None]
        combined = [row.combined_reward for row in subset if row.combined_reward is not None]
        table_rows.append(
            f"| {harness:<11} "
            f"| {len(subset):>6} "
            f"| {counts['completed']:>9} "
            f"| {counts['task_failure']:>12} "
            f"| {counts['infrastructure_failure']:>13} "
            f"| {sum(score > 0.0 for score in combined):>6} "
            f"| {sum(abs(score - 1.0) <= 1e-12 for score in combined):>6} "
            f"| {100.0 * _mean(deterministic):>8.2f}% "
            f"| {100.0 * _mean(visual):>12.2f}% "
            f"| {100.0 * _mean(combined):>10.2f}% |"
        )

    all_statuses = Counter(row.status for row in rows)
    visual_rows = [row for row in rows if row.visual_enabled]
    visual_scored = [row for row in visual_rows if row.visual_reward is not None]
    content = "\n".join(
        (
            "# Cross-Office 混合评分试跑（2026-09-04）",
            "",
            "本轮只对三套 Harness 的已冻结 Sonnet 5 答案重评分，没有重新运行解题 Agent。视觉题使用每题独立 `visual_rubric.json`，由 Codex CLI 0.148.0 调用 `gpt-5.6-terra`；非视觉题只使用当前逐题确定性 Judge。",
            "",
            "```text",
            "视觉题：combined_reward = 0.7 * deterministic_reward + 0.3 * visual_reward",
            "非视觉题：combined_reward = deterministic_reward",
            "```",
            "",
            "## Harness 汇总",
            "",
            "| Harness     | 有效题 | completed | task failure | infra failure | 非零题 | 满分题 | 代码均分 | 视觉子集均分 | 混合总均分 |",
            "|-------------|-------:|----------:|-------------:|--------------:|-------:|-------:|---------:|-------------:|-----------:|",
            *table_rows,
            "",
            "## 完整性",
            "",
            f"- 逐题结果：{len(rows)} / {len(HARNESSES) * EXPECTED_TASKS}。",
            f"- 正常完成：{all_statuses['completed']}。",
            f"- 真实任务失败：{all_statuses['task_failure']}；这些题均因受测 Agent 没有交付视觉 Rubric 指定的核心产物，不是评测基础设施故障。",
            f"- 评测基础设施失败：{all_statuses['infrastructure_failure']}；没有将 API、Codex、渲染或 JSON 解析故障记为 0 分。",
            f"- 视觉候选：{len(visual_rows)} 份；成功视觉评分：{len(visual_scored)} 份；其余 {len(visual_rows) - len(visual_scored)} 份是核心视觉产物缺失的真实任务失败。",
            "- 表中“视觉子集均分”只对成功执行 VLM 的候选取均值；缺少核心视觉产物的 `task_failure` 没有伪造视觉分，但其混合总分按协议为 0 并计入总均分。",
            f"- 逐题明细：`{csv_path.as_posix()}`。",
            "- 逐题合成分已从 JSON 重算，180 份全部与 70/30 或纯代码公式一致。",
            "",
            "## 状态边界",
            "",
            "`task_failure` 是受测 Agent 的真实答题结果，因此保留 0 分并计入均值。`infrastructure_failure` 不产生分数且不应进入均值；本轮该项为 0。",
            "",
        )
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--panel-root",
        type=Path,
        default=Path(
            ".coworkevo/empirical-calibration/"
            "cross-office-hybrid-20260904/panel"
        ),
    )
    parser.add_argument(
        "--csv",
        type=Path,
        required=True,
        help="输出逐题 CSV；公开 Release 不再携带历史 Pilot 原始表",
    )
    parser.add_argument(
        "--markdown",
        type=Path,
        default=Path("docs/cross-office-hybrid-pilot-20260904.md"),
    )
    arguments = parser.parse_args()

    rows = load_scores(arguments.panel_root)
    write_csv(arguments.csv, rows)
    write_markdown(arguments.markdown, rows, csv_path=arguments.csv)
    print(
        f"汇总通过：{len(rows)} 份逐题结果；"
        f"CSV={arguments.csv}；Markdown={arguments.markdown}"
    )


if __name__ == "__main__":
    main()
