#!/usr/bin/env python3
"""冻结 Word 重构版本：保留其他模块分区，最少移动 Word 题消除来源泄漏。"""

from __future__ import annotations

import json
import tempfile
from collections import Counter
from pathlib import Path

import freeze_benchmark_release as release


RENAMES = {
    "docx/officebench-1-16-test-sentence": "docx/officeval-003-coordination-minutes-audit",
    "docx/officebench-1-3-rename-docx": "docx/officeval-088-product-release-dossier",
}


def repair_assignments(tasks, components, previous):
    """动态规划精确满足容量；代价是相对旧版移动的 Word 题数。"""
    states = {(0, 0, 0): (0, [])}
    units = []
    for component in components:
        members = [i for i in component if tasks[i].domain == "docx"]
        if not members:
            continue
        fixed = {previous[tasks[i].instance_id] for i in component if tasks[i].domain != "docx"}
        if len(fixed) > 1:
            raise ValueError("其他模块的固定分区冲突，不能只调整 Word 修复")
        options = sorted(fixed) if fixed else list(release.PARTITIONS)
        units.append(members)
        next_states = {}
        for counts, (cost, choices) in states.items():
            for partition in options:
                updated = list(counts)
                updated[release.PARTITIONS.index(partition)] += len(members)
                if any(n > cap for n, cap in zip(updated, (30, 10, 20))):
                    continue
                extra = sum(previous[tasks[i].instance_id] != partition for i in members)
                key = tuple(updated)
                candidate = (cost + extra, choices + [partition])
                if key not in next_states or candidate < next_states[key]:
                    next_states[key] = candidate
        states = next_states
    if (30, 10, 20) not in states:
        raise ValueError("保持其他模块不变时无法满足 Word 30/10/20 容量")
    cost, choices = states[(30, 10, 20)]
    result = dict(previous)
    for members, partition in zip(units, choices):
        for i in members:
            result[tasks[i].instance_id] = partition
    return result, cost


def main():
    root = Path(__file__).resolve().parents[1]
    target = root / "benchmarks/cowork-evo-240-v0.3"
    if target.exists():
        raise FileExistsError("禁止覆盖已有 Release")
    tasks = release._discover_tasks(root)
    components = release._build_components(tasks)
    previous = release._load_frozen_assignments(root / "benchmarks/cowork-evo-240-v0.2")
    for old, new in RENAMES.items():
        previous[new] = previous.pop(old)
    if set(previous) != {t.instance_id for t in tasks}:
        raise ValueError("任务替换映射未覆盖当前任务集")
    repaired, cost = repair_assignments(tasks, components, previous)
    with tempfile.TemporaryDirectory(prefix="cowork-release-mapping-") as temporary:
        staging = Path(temporary)
        (staging / "splits").mkdir()
        for partition, filename in release.SPLIT_FILENAMES.items():
            release._write_jsonl(staging / "splits" / filename, (
                {"instance_id": key} for key in sorted(repaired) if repaired[key] == partition
            ))
        release.generate_release(root, target, target.name, source_release_root=staging)
    release._write_json(target / "migration.json", {
        "source_release": "cowork-evo-240-v0.2",
        "task_replacements": RENAMES,
        "policy": "保持其他三模块逐题分区不变；Word 按泄漏组件整体分配、精确30/10/20、最小化移动题数；不使用模型分数。",
        "minimum_moved_docx_tasks": cost,
        "moves": [{"instance_id": key, "before": previous[key], "after": repaired[key]}
                  for key in sorted(repaired) if repaired[key] != previous[key]],
        "counts": dict(Counter(repaired.values())),
    })
    release.validate_release(root, target)
    print(f"Word 最少移动 {cost} 题；其他模块分区保持不变。")


if __name__ == "__main__":
    main()
