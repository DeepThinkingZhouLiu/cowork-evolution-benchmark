#!/usr/bin/env python3
"""汇总 Claude Code 顺序运行 60 道 PPT 任务的分数、Token、成本和异常回放。"""

from __future__ import annotations

import argparse
import csv
import json
import statistics
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from run_claude_ppt60 import REPO_ROOT, ordered_tasks


DEFAULT_OUTPUT_ROOT = REPO_ROOT / "jobs/claude-sonnet5-ppt60-20260830"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-root", type=Path, default=DEFAULT_OUTPUT_ROOT)
    return parser.parse_args()


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def parse_timestamp(value: str | None) -> datetime | None:
    if not value:
        return None
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def source_family(task_name: str) -> str:
    return task_name.split("-", 1)[0]


def trial_dir(job_dir: Path) -> Path:
    candidates = sorted(
        path
        for path in job_dir.iterdir()
        if path.is_dir() and path.name != "runner-logs"
    )
    if len(candidates) != 1:
        raise RuntimeError(f"{job_dir}: 预期唯一 trial 目录，实际 {len(candidates)} 个")
    return candidates[0]


def final_agent_event(log_path: Path) -> dict[str, Any]:
    result: dict[str, Any] | None = None
    for line in log_path.read_text(encoding="utf-8").splitlines():
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        if event.get("type") == "result":
            result = event
    if result is None:
        raise RuntimeError(f"{log_path}: 缺少 Claude Code result 事件")
    return result


def recovered_interrupted_result(index: int, trial: Path) -> dict[str, Any]:
    """恢复 Harbor 最后写 JSON 前被中断的第 55 题，不重跑 Agent。"""
    reward_path = trial / "verifier/reward.txt"
    event = final_agent_event(trial / "agent/claude-code.txt")
    usage = event.get("usage") or {}
    cache_tokens = int(usage.get("cache_creation_input_tokens") or 0) + int(
        usage.get("cache_read_input_tokens") or 0
    )
    return {
        "reward": float(reward_path.read_text(encoding="utf-8").strip()),
        "exception": None,
        "status": "recovered_after_runner_interruption",
        "replayed": False,
        "input_tokens": int(usage.get("input_tokens") or 0) + cache_tokens,
        "cache_tokens": cache_tokens,
        "output_tokens": int(usage.get("output_tokens") or 0),
        "cost_usd": float(event.get("total_cost_usd") or 0.0),
        "duration_seconds": float(event.get("duration_ms") or 0) / 1000.0,
        "started_at": None,
        "finished_at": None,
        "trial_name": trial.name,
        "score_source": str(reward_path),
        "recovery_note": f"第 {index} 题 Agent、产物和 Judge 已完成，仅 Harbor 末尾汇总 JSON 未及写入。",
    }


def normal_or_replayed_result(index: int, trial: Path) -> dict[str, Any]:
    payload = load_json(trial / "result.json")
    verifier = payload.get("verifier_result") or {}
    reward = (verifier.get("rewards") or {}).get("reward")
    exception = payload.get("exception_info")
    status = "completed"
    replayed = False
    score_source = str(trial / "result.json")
    recovery_note = None

    if reward is None and index == 51:
        replay_reward = trial / "verifier/replay/reward.txt"
        replay_result = trial / "verifier/replay/result.json"
        if not replay_reward.is_file() or not replay_result.is_file():
            raise RuntimeError("第 51 题缺少同产物 Judge 回放结果")
        reward = float(replay_reward.read_text(encoding="utf-8").strip())
        status = "same_artifact_judge_replay"
        replayed = True
        score_source = str(replay_result)
        recovery_note = (
            "Agent 产物含非法浮点 EMU 坐标；Judge 改为将无效坐标判为布局失败后，"
            "对同一原始产物回放，未重跑 Agent。"
        )

    if reward is None:
        raise RuntimeError(f"第 {index} 题缺少有效 reward")

    agent_result = payload.get("agent_result") or {}
    started = parse_timestamp(payload.get("started_at"))
    finished = parse_timestamp(payload.get("finished_at"))
    duration_seconds = (
        (finished - started).total_seconds() if started and finished else None
    )
    return {
        "reward": float(reward),
        "exception": exception,
        "status": status,
        "replayed": replayed,
        "input_tokens": agent_result.get("n_input_tokens"),
        "cache_tokens": agent_result.get("n_cache_tokens"),
        "output_tokens": agent_result.get("n_output_tokens"),
        "cost_usd": agent_result.get("cost_usd"),
        "duration_seconds": duration_seconds,
        "started_at": payload.get("started_at"),
        "finished_at": payload.get("finished_at"),
        "trial_name": payload.get("trial_name") or trial.name,
        "score_source": score_source,
        "recovery_note": recovery_note,
    }


def task_result(index: int, task: Path, output_root: Path) -> dict[str, Any]:
    job_dir = output_root / f"{index:02d}-{task.name}"
    if not job_dir.is_dir():
        raise RuntimeError(f"第 {index} 题缺少 job 目录：{job_dir}")
    trial = trial_dir(job_dir)
    result_path = trial / "result.json"
    if result_path.is_file():
        details = normal_or_replayed_result(index, trial)
    elif index == 55 and (trial / "verifier/reward.txt").is_file():
        details = recovered_interrupted_result(index, trial)
    else:
        raise RuntimeError(f"第 {index} 题没有可恢复的结果：{trial}")

    return {
        "index": index,
        "task": task.name,
        "source_family": source_family(task.name),
        "job_dir": str(job_dir),
        "trial_dir": str(trial),
        **details,
    }


def aggregate(rows: list[dict[str, Any]]) -> dict[str, Any]:
    rewards = [float(row["reward"]) for row in rows]
    by_source: dict[str, list[float]] = defaultdict(list)
    for row in rows:
        by_source[row["source_family"]].append(float(row["reward"]))

    return {
        "task_count": len(rows),
        "reward_sum": sum(rewards),
        "mean_reward": statistics.fmean(rewards),
        "ppt_score_100": statistics.fmean(rewards) * 100.0,
        "median_reward": statistics.median(rewards),
        "min_reward": min(rewards),
        "max_reward": max(rewards),
        "full_score_count": sum(value == 1.0 for value in rewards),
        "pass_at_0_8_count": sum(value >= 0.8 for value in rewards),
        "nonzero_count": sum(value > 0 for value in rewards),
        "zero_count": sum(value == 0 for value in rewards),
        "same_artifact_replay_count": sum(bool(row["replayed"]) for row in rows),
        "interrupted_finalization_recovery_count": sum(
            row["status"] == "recovered_after_runner_interruption" for row in rows
        ),
        "input_tokens": sum(int(row["input_tokens"] or 0) for row in rows),
        "cache_tokens": sum(int(row["cache_tokens"] or 0) for row in rows),
        "output_tokens": sum(int(row["output_tokens"] or 0) for row in rows),
        "cost_usd": sum(float(row["cost_usd"] or 0.0) for row in rows),
        "agent_wall_seconds": sum(
            float(row["duration_seconds"] or 0.0) for row in rows
        ),
        "by_source": {
            family: {
                "count": len(values),
                "mean_reward": statistics.fmean(values),
                "score_100": statistics.fmean(values) * 100.0,
                "nonzero_count": sum(value > 0 for value in values),
                "full_score_count": sum(value == 1.0 for value in values),
            }
            for family, values in sorted(by_source.items())
        },
    }


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    fields = [
        "index",
        "task",
        "source_family",
        "reward",
        "status",
        "replayed",
        "input_tokens",
        "cache_tokens",
        "output_tokens",
        "cost_usd",
        "duration_seconds",
        "started_at",
        "finished_at",
        "trial_name",
        "job_dir",
        "score_source",
        "recovery_note",
    ]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def write_markdown(
    path: Path, rows: list[dict[str, Any]], summary: dict[str, Any]
) -> None:
    lines = [
        "# Claude Code Sonnet 5 — PPT 60 题基线结果",
        "",
        f"- 总题数：{summary['task_count']}",
        f"- PPT Score：{summary['ppt_score_100']:.4f}/100",
        f"- Mean Reward：{summary['mean_reward']:.6f}",
        f"- 满分题：{summary['full_score_count']}",
        f"- Reward >= 0.8：{summary['pass_at_0_8_count']}",
        f"- 非零题：{summary['nonzero_count']}",
        f"- 零分题：{summary['zero_count']}",
        f"- 总成本：${summary['cost_usd']:.4f}",
        "",
        "| # | 题目 | 来源 | Reward | 结果状态 |",
        "|--:|:-----|:-----|-------:|:---------|",
    ]
    for row in rows:
        lines.append(
            f"| {row['index']} | `{row['task']}` | {row['source_family']} | "
            f"{row['reward']:.6f} | {row['status']} |"
        )
    lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    args = parse_args()
    output_root = args.output_root.resolve()
    tasks = ordered_tasks(REPO_ROOT / "tasks/ppt")
    rows = [
        task_result(index, task, output_root)
        for index, task in enumerate(tasks, start=1)
    ]
    if len(rows) != 60:
        raise RuntimeError(f"预期 60 个结果，实际 {len(rows)} 个")

    summary = aggregate(rows)
    report_dir = output_root / "reports"
    report_dir.mkdir(parents=True, exist_ok=True)
    write_csv(report_dir / "per-task-results.csv", rows)
    (report_dir / "results.json").write_text(
        json.dumps({"summary": summary, "tasks": rows}, ensure_ascii=False, indent=2)
        + "\n",
        encoding="utf-8",
    )
    (report_dir / "scorecard.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    write_markdown(report_dir / "README.md", rows, summary)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
