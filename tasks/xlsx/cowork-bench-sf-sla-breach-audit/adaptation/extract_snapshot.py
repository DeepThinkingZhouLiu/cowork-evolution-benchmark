#!/usr/bin/env python3
"""从 Cowork-Bench 的 PostgreSQL dump 中逐行抽取本题使用的两张上游表。

该脚本只做原始表的无损静态化，不生成业务数据，也不改写任何字段。
"""

from __future__ import annotations

import argparse
import csv
import gzip
from pathlib import Path


TABLES = {
    "SUPPORT_CENTER__PUBLIC__TICKETS": "support_tickets.csv",
    "SUPPORT_CENTER__PUBLIC__SLA_POLICIES": "sla_policies.csv",
}


def postgres_unescape(value: str) -> str:
    """还原 PostgreSQL COPY 文本格式中的常见转义。"""
    if value == r"\N":
        return ""
    return (
        value.replace(r"\t", "\t")
        .replace(r"\n", "\n")
        .replace(r"\r", "\r")
        .replace(r"\\", "\\")
    )


def parse_copy_header(line: str) -> tuple[str, list[str]] | None:
    prefix = 'COPY sf_data."'
    marker = '" ('
    suffix = ") FROM stdin;\n"
    if not line.startswith(prefix) or marker not in line or not line.endswith(suffix):
        return None
    table, column_text = line[len(prefix) : -len(suffix)].split(marker, 1)
    if table not in TABLES:
        return None
    columns = [item.strip().strip('"') for item in column_text.split(",")]
    return table, columns


def extract(dump_path: Path, output_dir: Path) -> dict[str, int]:
    output_dir.mkdir(parents=True, exist_ok=True)
    handles: dict[str, object] = {}
    writers: dict[str, csv.writer] = {}
    counts = {table: 0 for table in TABLES}
    active_table: str | None = None
    active_columns: list[str] = []

    try:
        with gzip.open(dump_path, "rt", encoding="utf-8", newline="") as source:
            for line in source:
                if active_table is None:
                    parsed = parse_copy_header(line)
                    if parsed is None:
                        continue
                    active_table, active_columns = parsed
                    handle = (output_dir / TABLES[active_table]).open("w", encoding="utf-8", newline="")
                    writer = csv.writer(handle, lineterminator="\n")
                    writer.writerow(active_columns)
                    handles[active_table] = handle
                    writers[active_table] = writer
                    continue

                if line == "\\.\n":
                    handles[active_table].close()
                    active_table = None
                    active_columns = []
                    if all(table in writers for table in TABLES):
                        break
                    continue

                values = [postgres_unescape(value) for value in line.rstrip("\n").split("\t")]
                if len(values) != len(active_columns):
                    raise ValueError(
                        f"{active_table} 列数不匹配：期望 {len(active_columns)}，实际 {len(values)}"
                    )
                writers[active_table].writerow(values)
                counts[active_table] += 1
    finally:
        for handle in handles.values():
            if not handle.closed:
                handle.close()

    missing = [table for table in TABLES if table not in writers]
    if missing:
        raise RuntimeError(f"在 dump 中未找到表：{missing}")
    return counts


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("dump_path", type=Path)
    parser.add_argument("output_dir", type=Path)
    args = parser.parse_args()
    for table, count in extract(args.dump_path, args.output_dir).items():
        print(f"{table}: {count} 行")


if __name__ == "__main__":
    main()
