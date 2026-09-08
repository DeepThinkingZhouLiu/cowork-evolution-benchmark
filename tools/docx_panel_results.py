"""Word Panel 的历史证据恢复、预算耗尽判定和分数四档汇总。"""

from __future__ import annotations

import csv
import json
import re
import tomllib
from pathlib import Path


def recover_source(row, source_run):
    path = row.get("source_result_path")
    if path and Path(path).is_file():
        return Path(path)
    candidates = list(Path(source_run).glob(f"trials/*-{row['task']}-a*/*/result.json"))
    if len(candidates) != 1:
        raise ValueError(f"旧结果无法唯一定位：{row['task']}，候选={len(candidates)}")
    payload = json.loads(candidates[0].read_text())
    if not payload.get("finished_at") or not (payload.get("verifier_result") or {}).get("rewards"):
        raise ValueError(f"旧试验未完成：{row['task']}")
    return candidates[0]


def budget_timeout_record(job_dir, record):
    """SDK 超时可能未落盘 usage；有实测用量日志时按真实预算耗尽记分。"""
    exception = record.get("exception") or {}
    if (not record.get("infrastructure_failure") or record.get("reward") is None
            or exception.get("exception_type") != "AgentTimeoutError"):
        return record
    paths = list(job_dir.glob("*/agent/openhands_sdk.txt"))
    for path in paths:
        text = path.read_text(errors="replace")
        if "Agent Action" in text and re.search(r"Tokens:.*input\s+[1-9]", text):
            return {**record, "infrastructure_failure": False,
                    "classification": "agent_budget_exhausted", "classification_evidence": str(path)}
    return record


def active_harbor_pids(output):
    """接续已有单题进程，避免主调度器恢复时重复发起付费作答。"""
    target = str((output / "trials").resolve())
    result = []
    for path in Path("/proc").glob("[0-9]*/cmdline"):
        try:
            args = [part.decode(errors="replace") for part in path.read_bytes().split(b"\0") if part]
            if "--jobs-dir" in args and args[args.index("--jobs-dir") + 1] == target and "--job-name" in args:
                result.append(int(path.parent.name))
        except (OSError, IndexError):
            continue
    return result


def difficulty_report(root, output, harnesses):
    panels = {h: json.loads((output / h / "results.json").read_text()) for h in harnesses}
    if any(p["valid_count"] != 60 for p in panels.values()):
        return
    maps = {h: {r["task"]: r for r in p["tasks"]} for h, p in panels.items()}
    names = set(next(iter(maps.values())))
    if len(names) != 60 or any(set(m) != names for m in maps.values()):
        raise ValueError("三套 Harness 结果没有严格逐题配对")
    rows = []
    for name in names:
        scores = {h: maps[h][name]["reward"] for h in harnesses}
        metadata = tomllib.loads((root / "tasks/docx" / name / "task.toml").read_text())["metadata"]
        mean = sum(scores.values()) / 3
        rows.append({"task": name, **scores, "panel_mean": mean, "empirical_difficulty": 1 - mean,
                     "estimated_minutes": int(metadata.get("estimated_minutes", 0))})
    # 主排序仅由分数决定；同分时按既有预计工作量与任务编号稳定排序。
    rows.sort(key=lambda r: (-round(r["panel_mean"], 12), r["estimated_minutes"], r["task"]))
    groups, offset = [], 0
    for label, count in (("Easy", 12), ("Medium", 18), ("Hard", 18), ("Expert", 12)):
        selected = rows[offset:offset + count]
        for i, row in enumerate(selected, offset + 1):
            row.update(rank=i, pilot_tier=label)
        groups.append({"tier": label, "count": count,
                       **{h: sum(r[h] for r in selected) / count for h in harnesses},
                       "panel_mean": sum(r["panel_mean"] for r in selected) / count})
        offset += count
    ties = [{"after_rank": boundary, "panel_mean": rows[boundary - 1]["panel_mean"]}
            for boundary in (12, 30, 48)
            if abs(rows[boundary - 1]["panel_mean"] - rows[boundary]["panel_mean"]) < 1e-12]
    data = {"method": "Panel=(Claude+OpenCode+OpenHands)/3；经验难度=1-Panel；按分数降序分12/18/18/12档。",
            "tie_rule": "边界同分时按原预计工作量升序、任务编号排序；同分跨档不代表实测能力差异。",
            "scope": "同一Sonnet 5模型、每Harness每题一次有效作答的相对分位档，不改task.toml正式标签。",
            "boundary_ties": ties, "tiers": groups, "tasks": rows}
    (output / "difficulty.json").write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n")
    fields = ["rank", "task", "pilot_tier", *harnesses, "panel_mean", "empirical_difficulty", "estimated_minutes"]
    with (output / "difficulty.csv").open("w", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader(); writer.writerows(rows)
    lines = ["Word 60 题三 Harness 分数与经验难度", "", data["method"], "",
             "| 难度   | 题数 | Claude Code | OpenCode | OpenHands | Panel 均分 | 经验难度 |",
             "|--------|-----:|------------:|---------:|----------:|-----------:|---------:|"]
    for g in groups:
        lines.append(f"| {g['tier']:<6} | {g['count']:4} | {g['claude']*100:10.2f}% | {g['opencode']*100:7.2f}% | {g['openhands']*100:8.2f}% | {g['panel_mean']*100:9.2f}% | {(1-g['panel_mean'])*100:7.2f}% |")
    means = {h: sum(r[h] for r in rows) / 60 for h in harnesses}
    panel_mean = sum(means.values()) / 3
    lines.append(f"| 全部   |   60 | {means['claude']*100:10.2f}% | {means['opencode']*100:7.2f}% | {means['openhands']*100:8.2f}% | {panel_mean*100:9.2f}% | {(1-panel_mean)*100:7.2f}% |")
    lines.extend(["", data["scope"], "", f"同分边界：{ties or '无'}。" + data["tie_rule"]])
    (output / "difficulty.md").write_text("\n".join(lines) + "\n")
