"""Benchmark Manifest 的跨字段语义校验。"""

from __future__ import annotations

import json
from pathlib import Path

from cowork_evo_bench.validation import ValidationIssue

_EXPECTED_VISIBILITY = {
    "feedback": "detailed",
    "selection": "aggregate_only",
    "final": "sealed",
}


def validate_benchmark_payload(payload: object, location: str = "benchmark") -> list[ValidationIssue]:
    """检查三段式划分、数量和跨分区重复。"""

    issues: list[ValidationIssue] = []
    if not isinstance(payload, dict):
        return [ValidationIssue("benchmark.root_invalid", location, "根节点必须是对象。")]

    spec = payload.get("spec")
    if not isinstance(spec, dict):
        return [ValidationIssue("benchmark.spec_missing", location, "缺少 spec 对象。")]

    partitions = spec.get("partitions")
    if not isinstance(partitions, dict):
        return [
            ValidationIssue("benchmark.partitions_missing", location, "缺少 partitions 对象。")
        ]

    all_ids: list[str] = []
    for name, visibility in _EXPECTED_VISIBILITY.items():
        partition_location = f"{location}#spec.partitions.{name}"
        partition = partitions.get(name)
        if not isinstance(partition, dict):
            issues.append(
                ValidationIssue("benchmark.partition_missing", partition_location, "缺少分区。")
            )
            continue
        if partition.get("visibility") != visibility:
            issues.append(
                ValidationIssue(
                    "benchmark.visibility_invalid",
                    partition_location,
                    f"可见性必须是 {visibility}。",
                )
            )
        instance_ids = partition.get("instance_ids")
        if not isinstance(instance_ids, list) or not instance_ids:
            issues.append(
                ValidationIssue(
                    "benchmark.instances_missing",
                    partition_location,
                    "分区至少需要一个任务编号。",
                )
            )
            continue
        if any(not isinstance(instance_id, str) or not instance_id for instance_id in instance_ids):
            issues.append(
                ValidationIssue(
                    "benchmark.instance_id_invalid",
                    partition_location,
                    "任务编号必须是非空字符串。",
                )
            )
        if len(instance_ids) != len(set(instance_ids)):
            issues.append(
                ValidationIssue(
                    "benchmark.partition_duplicates",
                    partition_location,
                    "同一分区内不能重复任务。",
                )
            )
        expected_count = partition.get("expected_count")
        if expected_count != len(instance_ids):
            issues.append(
                ValidationIssue(
                    "benchmark.partition_count_mismatch",
                    partition_location,
                    "expected_count 与任务编号数量不一致。",
                )
            )
        all_ids.extend(instance_id for instance_id in instance_ids if isinstance(instance_id, str))

    if len(all_ids) != len(set(all_ids)):
        issues.append(
            ValidationIssue(
                "benchmark.cross_partition_leakage",
                location,
                "任务编号出现在多个分区，存在直接泄漏。",
            )
        )
    if spec.get("expected_total") != len(all_ids):
        issues.append(
            ValidationIssue(
                "benchmark.total_count_mismatch",
                location,
                "expected_total 与全部分区任务数量不一致。",
            )
        )
    return issues


def validate_benchmark_file(path: Path) -> list[ValidationIssue]:
    """读取 JSON Manifest 并执行语义校验。"""

    path = path.resolve()
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        return [
            ValidationIssue(
                "benchmark.invalid_json",
                str(path),
                f"无法读取 Benchmark JSON：{error}",
            )
        ]
    return validate_benchmark_payload(payload, str(path))
