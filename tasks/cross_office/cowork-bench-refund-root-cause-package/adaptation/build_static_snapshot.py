#!/usr/bin/env python3
"""将 Cowork-Bench 动态 PostgreSQL 数据忠实静态化为本题输入。

脚本只做可审计的数据投影：

- 从上游 ``db/init.sql.gz`` 的 ``COPY wc.*`` 数据块读取真实记录；
- 逐字执行上游 preprocess 对订单 39/49/110 的全额退款对齐规则；
- 使用上游 ``wc_reviews_parts`` 中按 id 给出的俄文评论正文，复现动态环境的俄语化后状态；
- 不创建、抽样、删减或改写任何退款、订单、商品或评论事实。

用法：

    python build_static_snapshot.py \
      --cowork-bench-root /path/to/Cowork-Bench \
      --output-dir ../data/input_files
"""

from __future__ import annotations

import argparse
import csv
import glob
import gzip
import json
from decimal import Decimal
from pathlib import Path


TABLES = {"orders", "product_reviews", "products", "refunds"}


def parse_copy_blocks(seed_path: Path) -> dict[str, tuple[list[str], list[list[str]]]]:
    """解析 PostgreSQL plain-text COPY 块。

上游生成器也使用制表符分列；JSON 字段中的制表符会以转义形式保存，
因此此处与上游 ``wc_gen_seed_sql.py`` 的解析方法保持一致。
"""
    with gzip.open(seed_path, "rt", encoding="utf-8", errors="strict") as handle:
        lines = handle.read().splitlines()

    blocks: dict[str, tuple[list[str], list[list[str]]]] = {}
    index = 0
    while index < len(lines):
        line = lines[index]
        if line.startswith("COPY wc.") and "FROM stdin;" in line:
            table = line[len("COPY wc.") : line.index("(")].strip().strip('"')
            columns = [
                column.strip().strip('"')
                for column in line[line.index("(") + 1 : line.rindex(")")].split(",")
            ]
            rows: list[list[str]] = []
            index += 1
            while index < len(lines) and lines[index] != r"\.":
                rows.append(lines[index].split("\t"))
                index += 1
            if table in TABLES:
                blocks[table] = (columns, rows)
        index += 1

    missing = TABLES.difference(blocks)
    if missing:
        raise RuntimeError(f"种子数据缺少表：{sorted(missing)}")
    return blocks


def records(block: tuple[list[str], list[list[str]]]) -> list[dict[str, str]]:
    columns, rows = block
    return [dict(zip(columns, row, strict=True)) for row in rows]


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, object]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def load_russian_reviews(root: Path) -> dict[int, str]:
    merged: dict[int, str] = {}
    pattern = root / "scripts" / "wc_reviews_parts" / "part_*.json"
    for source in sorted(glob.glob(str(pattern))):
        with Path(source).open(encoding="utf-8") as handle:
            payload = json.load(handle)
        for review_id, body in payload.items():
            merged[int(review_id)] = body["review"]
    return merged


def build(root: Path, output_dir: Path) -> None:
    blocks = parse_copy_blocks(root / "db" / "init.sql.gz")
    refunds = records(blocks["refunds"])
    orders = records(blocks["orders"])
    products = records(blocks["products"])
    reviews = records(blocks["product_reviews"])

    order_by_id = {int(row["id"]): row for row in orders}
    refund_count_by_order: dict[int, int] = {}
    for row in refunds:
        order_id = int(row["order_id"])
        refund_count_by_order[order_id] = refund_count_by_order.get(order_id, 0) + 1

    # 与上游 preprocess/main.py::fix_full_refund_amounts 逐条对齐。
    for row in refunds:
        order_id = int(row["order_id"])
        order = order_by_id[order_id]
        if (
            order_id in {39, 49, 110}
            and order["status"] == "refunded"
            and refund_count_by_order[order_id] == 1
            and Decimal(row["amount"]) != Decimal(order["total"])
        ):
            row["amount"] = order["total"]

    refund_order_ids = {int(row["order_id"]) for row in refunds}
    related_orders = [row for row in orders if int(row["id"]) in refund_order_ids]
    related_orders.sort(key=lambda row: int(row["id"]))

    output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(
        output_dir / "refunds.csv",
        ["id", "order_id", "amount", "reason", "refunded_by", "date_created"],
        [
            {
                "id": int(row["id"]),
                "order_id": int(row["order_id"]),
                "amount": f'{Decimal(row["amount"]):.2f}',
                "reason": row["reason"],
                "refunded_by": int(row["refunded_by"]),
                "date_created": row["date_created"],
            }
            for row in sorted(refunds, key=lambda item: int(item["id"]))
        ],
    )

    order_projection = []
    for row in related_orders:
        order_projection.append(
            {
                "id": int(row["id"]),
                "status": row["status"],
                "currency": row["currency"],
                "total": f'{Decimal(row["total"]):.2f}',
                "line_items": json.loads(row["line_items"]),
                "date_created": row["date_created"],
                "date_modified": row["date_modified"],
            }
        )
    (output_dir / "refund_orders.json").write_text(
        json.dumps(order_projection, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    write_csv(
        output_dir / "products.csv",
        ["id", "name", "sku", "average_rating", "rating_count"],
        [
            {
                "id": int(row["id"]),
                "name": row["name"],
                "sku": row["sku"],
                "average_rating": row["average_rating"],
                "rating_count": int(row["rating_count"]),
            }
            for row in sorted(products, key=lambda item: int(item["id"]))
        ],
    )

    russian_reviews = load_russian_reviews(root)
    missing_review_ids = sorted(
        int(row["id"]) for row in reviews if int(row["id"]) not in russian_reviews
    )
    if missing_review_ids:
        raise RuntimeError(f"上游俄文评论投影缺失 id：{missing_review_ids[:20]}")
    write_csv(
        output_dir / "product_reviews.csv",
        [
            "id",
            "product_id",
            "reviewer",
            "reviewer_email",
            "review",
            "rating",
            "verified",
            "status",
            "date_created",
        ],
        [
            {
                "id": int(row["id"]),
                "product_id": int(row["product_id"]),
                "reviewer": row["reviewer"],
                "reviewer_email": row["reviewer_email"],
                "review": russian_reviews[int(row["id"])],
                "rating": int(row["rating"]),
                "verified": row["verified"],
                "status": row["status"],
                "date_created": row["date_created"],
            }
            for row in sorted(reviews, key=lambda item: int(item["id"]))
        ],
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--cowork-bench-root", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args()
    build(args.cowork_bench_root.resolve(), args.output_dir.resolve())


if __name__ == "__main__":
    main()
