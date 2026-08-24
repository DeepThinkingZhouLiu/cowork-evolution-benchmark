#!/usr/bin/env python3
"""从 Cowork-Bench 固定 SQL dump 中抽取本题的 Q4 2024 离线快照。

这不是通用数据生成器：它只识别本题上游 Query 指定的两张表，
保留原始订单金额、日期、客户 ID、地区和客户分群，并在写出前
用上游 Evaluator 的真值校验总数。
"""

from __future__ import annotations

import argparse
import csv
import gzip
from collections import defaultdict
from pathlib import Path


CUSTOMER_COPY_PREFIX = 'COPY sf_data."SALES_DW__PUBLIC__CUSTOMERS"'
ORDER_COPY_PREFIX = 'COPY sf_data."SALES_DW__PUBLIC__ORDERS"'

# SQL dump 的集中俄语化标签。键是 dump 中的原始值。
EXPECTED_RUSSIAN_REGIONS = {
    "Азиатско-Тихоокеанский регион": (538, 78879.15),
    "Ближний Восток": (486, 88724.86),
    "Европа": (528, 84306.43),
    "Латинская Америка": (472, 78432.77),
    "Северная Америка": (555, 81513.97),
}


def read_copy_sections(dump_path: Path) -> tuple[list[list[str]], list[list[str]]]:
    customers: list[list[str]] = []
    orders: list[list[str]] = []
    active: list[list[str]] | None = None
    with gzip.open(dump_path, "rt", encoding="utf-8") as stream:
        for raw_line in stream:
            if raw_line.startswith(CUSTOMER_COPY_PREFIX):
                active = customers
                continue
            if raw_line.startswith(ORDER_COPY_PREFIX):
                active = orders
                continue
            if active is None:
                continue
            line = raw_line.rstrip("\n")
            if line == "\\.":
                active = None
                continue
            active.append(line.split("\t"))
    if len(customers) != 2000 or len(orders) != 20000:
        raise ValueError(f"上游表行数不符：customers={len(customers)}, orders={len(orders)}")
    return customers, orders


def write_snapshot(dump_path: Path, output_dir: Path) -> None:
    customers, orders = read_copy_sections(dump_path)
    customer_by_id = {row[0]: row for row in customers}
    q4_orders = [row for row in orders if "2024-10-01" <= row[2] < "2025-01-01"]
    relevant_customer_ids = sorted({row[1] for row in q4_orders}, key=int)

    region_totals: dict[str, list[float]] = defaultdict(lambda: [0, 0.0])
    for order in q4_orders:
        customer = customer_by_id[order[1]]
        region_totals[customer[4]][0] += 1
        region_totals[customer[4]][1] += float(order[9])
    actual = {region: (int(values[0]), round(values[1], 2)) for region, values in region_totals.items()}
    if actual != EXPECTED_RUSSIAN_REGIONS:
        raise ValueError(f"地区聚合不符合上游真值：{actual}")
    if len(q4_orders) != 2579 or len(relevant_customer_ids) != 1451:
        raise ValueError(f"Q4 记录数异常：orders={len(q4_orders)}, customers={len(relevant_customer_ids)}")

    output_dir.mkdir(parents=True, exist_ok=True)
    orders_path = output_dir / "q4_orders_2024.csv"
    with orders_path.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.writer(stream)
        writer.writerow(["ORDER_ID", "CUSTOMER_ID", "ORDER_DATE", "TOTAL_AMOUNT"])
        for row in sorted(q4_orders, key=lambda item: int(item[0])):
            writer.writerow([row[0], row[1], row[2], row[9]])

    customers_path = output_dir / "customer_dimension.csv"
    with customers_path.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.writer(stream)
        writer.writerow(["CUSTOMER_ID", "SEGMENT", "REGION"])
        for customer_id in relevant_customer_ids:
            row = customer_by_id[customer_id]
            writer.writerow([row[0], row[3], row[4]])

    print(f"{orders_path}: {len(q4_orders)} 行")
    print(f"{customers_path}: {len(relevant_customer_ids)} 行")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dump", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    write_snapshot(args.dump, args.output_dir)


if __name__ == "__main__":
    main()
