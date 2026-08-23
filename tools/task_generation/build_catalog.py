#!/usr/bin/env python3
"""校验公开种子题元数据，并生成稳定的任务目录与覆盖统计。"""

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator


REPO_ROOT = Path(__file__).resolve().parents[2]
TASKS_ROOT = REPO_ROOT / "tasks"
TAXONOMY_PATH = REPO_ROOT / "configs/taxonomy/cowork-office-v0.1.json"
SCHEMA_PATH = REPO_ROOT / "schemas/task-metadata.schema.json"


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _task_roots() -> list[Path]:
    return sorted(path.parent.parent for path in TASKS_ROOT.glob("*/*/metadata/task.json"))


def _count_files(root: Path) -> tuple[int, int]:
    input_count = sum(path.is_file() for path in (root / "data/input_files").rglob("*"))
    metadata = _read_json(root / "metadata/task.json")
    output_count = len(metadata["output_filenames"])
    return input_count, output_count


def _validate_and_collect() -> tuple[list[dict[str, Any]], dict[str, Any]]:
    taxonomy = _read_json(TAXONOMY_PATH)
    schema = _read_json(SCHEMA_PATH)
    validator = Draft202012Validator(schema)
    records: list[dict[str, Any]] = []
    seen_ids: set[str] = set()

    for root in _task_roots():
        metadata_path = root / "metadata/task.json"
        metadata = _read_json(metadata_path)
        errors = sorted(validator.iter_errors(metadata), key=lambda error: list(error.path))
        if errors:
            details = "; ".join(
                f"{'.'.join(map(str, error.path)) or '<root>'}: {error.message}"
                for error in errors
            )
            raise ValueError(f"{metadata_path} 不符合 Schema：{details}")
        if metadata["task_id"] in seen_ids:
            raise ValueError(f"任务编号重复：{metadata['task_id']}")
        seen_ids.add(metadata["task_id"])
        expected_domain = root.parent.name
        if metadata["domain"] != expected_domain:
            raise ValueError(
                f"{metadata_path} 的 domain={metadata['domain']} 与目录 {expected_domain} 不一致"
            )
        expected_families = taxonomy["task_families"].get(metadata["domain"], [])
        if metadata["task_family"] not in expected_families:
            raise ValueError(f"未登记的任务家族：{metadata['task_family']}")
        input_count, output_count = _count_files(root)
        if input_count < 1 or output_count < 1:
            raise ValueError(f"{root} 缺少真实输入或声明输出")
        records.append(
            {
                **metadata,
                "task_path": str(root.relative_to(REPO_ROOT)),
                "input_file_count": input_count,
                "output_file_count": output_count,
            }
        )

    family_counts = Counter((item["domain"], item["task_family"]) for item in records)
    coverage_errors: list[str] = []
    for domain, families in taxonomy["task_families"].items():
        for family in families:
            count = family_counts[(domain, family)]
            if not 1 <= count <= 2:
                coverage_errors.append(f"{domain}/{family}={count}，应为 1–2 题")
    if coverage_errors:
        raise ValueError("任务家族覆盖不完整：" + "；".join(coverage_errors))

    summary = {
        "taxonomy_version": taxonomy["taxonomy_version"],
        "task_count": len(records),
        "family_count": len(family_counts),
        "domain_counts": dict(sorted(Counter(item["domain"] for item in records).items())),
        "difficulty_counts": dict(
            sorted(Counter(item["difficulty_tier"] for item in records).items())
        ),
        "capability_counts": dict(
            sorted(Counter(capability for item in records for capability in item["capability_ids"]).items())
        ),
        "supplier_capability_counts": dict(
            sorted(
                Counter(
                    capability
                    for item in records
                    for capability in item["supplier_capability_ids"]
                ).items()
            )
        ),
        "input_file_count": sum(item["input_file_count"] for item in records),
        "output_file_count": sum(item["output_file_count"] for item in records),
    }
    return records, summary


def _write_outputs(records: list[dict[str, Any]], summary: dict[str, Any]) -> None:
    catalog_path = TASKS_ROOT / "catalog.jsonl"
    catalog_path.write_text(
        "".join(json.dumps(item, ensure_ascii=False, sort_keys=True) + "\n" for item in records),
        encoding="utf-8",
    )
    (TASKS_ROOT / "catalog-summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )

    headers = ["领域", "任务家族", "任务编号", "难度", "输入/输出", "能力标签"]
    rows: list[list[str]] = []
    for item in records:
        short_id = item["task_id"].split("/", 1)[1]
        rows.append(
            [
                item["domain"],
                f"`{item['task_family']}`",
                f"`{short_id}`",
                item["difficulty_tier"],
                f"{item['input_file_count']}/{item['output_file_count']}",
                ", ".join(item["capability_ids"]),
            ]
        )
    widths = [
        max(len(headers[index]), *(len(row[index]) for row in rows))
        for index in range(len(headers))
    ]

    def table_line(cells: list[str]) -> str:
        return "| " + " | ".join(
            cell.ljust(widths[index]) for index, cell in enumerate(cells)
        ) + " |"

    lines = [
        "# 公开种子题目录",
        "",
        "> 本文件由 `tools/task_generation/build_catalog.py` 生成。当前内容是公开开发种子题，不等同于论文密封测试集。",
        "",
        f"共 {summary['task_count']} 题，覆盖 {summary['family_count']} 个任务家族；"
        f"包含 {summary['input_file_count']} 个输入文件和 {summary['output_file_count']} 个 Golden 交付物。",
        "",
        table_line(headers),
        "| " + " | ".join("-" * width for width in widths) + " |",
    ]
    lines.extend(table_line(row) for row in rows)
    (TASKS_ROOT / "CATALOG.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="只校验，不覆盖目录文件")
    args = parser.parse_args()
    records, summary = _validate_and_collect()
    if not args.check:
        _write_outputs(records, summary)
    print(
        f"公开任务目录校验通过：{summary['task_count']} 题，"
        f"{summary['family_count']} 个家族，{summary['input_file_count']} 个输入文件。"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
