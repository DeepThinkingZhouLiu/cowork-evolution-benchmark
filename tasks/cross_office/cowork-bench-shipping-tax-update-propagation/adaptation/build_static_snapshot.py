#!/usr/bin/env python3
"""从 Cowork-Bench 原始 PostgreSQL dump 忠实抽取本题所需的六张表。

脚本不改写、补全或推断任何业务值；它只解析 COPY ... FROM stdin
区块，将 ``\\N`` 投影为 CSV 空值，并保留原始行顺序。
"""

from __future__ import annotations

import argparse
import csv
import gzip
import re
from pathlib import Path


TABLE_OUTPUTS = {
    "shipping_zones": "shipping_zones.csv",
    "shipping_zone_locations": "shipping_zone_locations.csv",
    "shipping_zone_methods": "shipping_zone_methods.csv",
    "shipping_methods": "shipping_methods.csv",
    "tax_classes": "tax_classes.csv",
    "tax_rates": "tax_rates.csv",
}

COPY_RE = re.compile(r'^COPY wc\.([a-z_]+) \((.+)\) FROM stdin;$')


def parse_dump(path: Path) -> tuple[dict[str, tuple[list[str], list[list[str]]]], list[str]]:
    selected: dict[str, tuple[list[str], list[list[str]]]] = {}
    raw_lines: list[str] = []
    active_table: str | None = None
    active_columns: list[str] = []
    active_rows: list[list[str]] = []

    with gzip.open(path, "rt", encoding="utf-8", newline="") as handle:
        for physical_line in handle:
            line = physical_line.rstrip("\r\n")
            if active_table is None:
                match = COPY_RE.match(line)
                if match and match.group(1) in TABLE_OUTPUTS:
                    active_table = match.group(1)
                    active_columns = [item.strip().strip('"') for item in match.group(2).split(",")]
                    active_rows = []
                    raw_lines.append(line)
                continue

            raw_lines.append(line)
            if line == r"\.":
                selected[active_table] = (active_columns, active_rows)
                raw_lines.append("")
                active_table = None
                active_columns = []
                active_rows = []
                continue

            values = ["" if value == r"\N" else value for value in line.split("\t")]
            if len(values) != len(active_columns):
                raise ValueError(
                    f"{active_table} 列数错误：期望 {len(active_columns)}，实际 {len(values)}"
                )
            active_rows.append(values)

    missing = sorted(set(TABLE_OUTPUTS) - set(selected))
    if missing:
        raise ValueError(f"原始 dump 缺少表：{', '.join(missing)}")
    return selected, raw_lines


def write_snapshot(
    tables: dict[str, tuple[list[str], list[list[str]]]],
    output_dir: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    for table, filename in TABLE_OUTPUTS.items():
        columns, rows = tables[table]
        target = output_dir / filename
        with target.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.writer(handle, lineterminator="\n")
            writer.writerow(columns)
            writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db-init", required=True, type=Path, help="Cowork-Bench/db/init.sql.gz")
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--raw-snapshot", type=Path)
    args = parser.parse_args()

    tables, raw_lines = parse_dump(args.db_init.resolve())
    write_snapshot(tables, args.output_dir.resolve())
    if args.raw_snapshot:
        target = args.raw_snapshot.resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("\n".join(raw_lines).rstrip() + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
