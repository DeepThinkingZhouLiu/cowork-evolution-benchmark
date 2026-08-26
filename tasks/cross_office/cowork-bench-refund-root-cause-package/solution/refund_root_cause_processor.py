#!/usr/bin/env python3
"""退款根因分析处理器。

该脚本只从命令行指定的输入目录读取数据，并生成可回放的
``analysis_results.json``；不包含本题固定答案。
"""

from __future__ import annotations

import argparse
import csv
import json
from collections import Counter, defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any


MONEY = Decimal("0.01")

ROOT_CAUSE_MAP = {
    "Product defective": "Manufacturing Defect",
    "Product arrived damaged": "Shipping Damage",
    "Wrong item received": "Wrong Item",
    "Changed my mind": "Customer Expectations",
    "Item not as described": "Customer Expectations",
    "Duplicate order": "Other",
}


def money(value: Decimal) -> Decimal:
    return value.quantize(MONEY, rounding=ROUND_HALF_UP)


def as_json_number(value: Decimal) -> float:
    return float(money(value))


def load_inputs(input_dir: Path) -> tuple[list[dict[str, str]], dict[int, dict[str, Any]], dict[int, list[int]]]:
    with (input_dir / "refunds.csv").open(encoding="utf-8", newline="") as handle:
        refunds = list(csv.DictReader(handle))
    orders_payload = json.loads((input_dir / "refund_orders.json").read_text(encoding="utf-8"))
    orders = {int(row["id"]): row for row in orders_payload}
    ratings: dict[int, list[int]] = defaultdict(list)
    with (input_dir / "product_reviews.csv").open(encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            ratings[int(row["product_id"])].append(int(row["rating"]))
    return refunds, orders, ratings


def analyze(input_dir: Path) -> dict[str, Any]:
    refunds, orders, ratings = load_inputs(input_dir)
    details: list[dict[str, Any]] = []
    impact: dict[str, dict[str, Any]] = defaultdict(
        lambda: {"refund_count": 0, "total_refund_amount": Decimal("0"), "product_id": None}
    )
    severities: Counter[str] = Counter()
    reasons: Counter[str] = Counter()
    root_causes: Counter[str] = Counter()

    for refund in sorted(refunds, key=lambda row: int(row["id"])):
        refund_id = int(refund["id"])
        order_id = int(refund["order_id"])
        if order_id not in orders:
            raise ValueError(f"退款 {refund_id} 找不到订单 {order_id}")
        items = orders[order_id].get("line_items") or []
        if not items:
            raise ValueError(f"订单 {order_id} 没有 line_items")

        amount = Decimal(refund["amount"])
        severity = "Critical" if amount > 50 else "Major" if amount >= 20 else "Minor"
        reason = refund["reason"]
        product_names = [str(item["name"]) for item in items]
        details.append(
            {
                "Refund_ID": refund_id,
                "Order_ID": order_id,
                "Product_Name": "; ".join(product_names),
                "Refund_Amount": as_json_number(amount),
                "Severity": severity,
                "Refund_Reason": reason,
                "Date": refund["date_created"][:10],
            }
        )
        severities[severity] += 1
        reasons[reason] += 1
        root_causes[ROOT_CAUSE_MAP.get(reason, "Other")] += 1

        per_product = amount / Decimal(len(items))
        for item in items:
            entry = impact[str(item["name"])]
            entry["refund_count"] += 1
            entry["total_refund_amount"] += per_product
            entry["product_id"] = int(item["product_id"])

    product_rows: list[dict[str, Any]] = []
    for product_name, item in impact.items():
        product_ratings = ratings.get(int(item["product_id"]), [])
        if product_ratings:
            average = money(Decimal(sum(product_ratings)) / Decimal(len(product_ratings)))
            issue_type = "Quality Issue" if average < 3 else "Service Issue"
            average_value: float | None = float(average)
        else:
            average_value = None
            issue_type = "No Reviews"
        product_rows.append(
            {
                "Product_Name": product_name,
                "Refund_Count": int(item["refund_count"]),
                "Total_Refund_Amount": as_json_number(item["total_refund_amount"]),
                "Avg_Review_Rating": average_value,
                "Review_Count": len(product_ratings),
                "Issue_Type": issue_type,
                "Requires_Investigation": "Yes" if item["refund_count"] >= 2 else "No",
            }
        )
    product_rows.sort(key=lambda row: (-row["Refund_Count"], row["Product_Name"]))

    summary = {
        "Total_Refunds": len(details),
        "Total_Refund_Amount": as_json_number(sum(Decimal(row["amount"]) for row in refunds)),
        "Critical_Count": severities["Critical"],
        "Major_Count": severities["Major"],
        "Minor_Count": severities["Minor"],
        "Products_Affected": len(product_rows),
        "Quality_Issues": sum(row["Issue_Type"] == "Quality Issue" for row in product_rows),
        "Service_Issues": sum(row["Issue_Type"] == "Service Issue" for row in product_rows),
        "Products_Requiring_Investigation": sum(
            row["Requires_Investigation"] == "Yes" for row in product_rows
        ),
    }

    reason_rows = [
        {"Refund_Reason": reason, "Refund_Count": count}
        for reason, count in sorted(reasons.items(), key=lambda item: (-item[1], item[0]))
    ]
    root_rows = [
        {"Root_Cause": cause, "Refund_Count": count}
        for cause, count in sorted(root_causes.items(), key=lambda item: (-item[1], item[0]))
    ]
    investigation_rows = [
        row for row in product_rows if row["Requires_Investigation"] == "Yes"
    ]
    return {
        "refund_details": details,
        "product_impact": product_rows,
        "summary": summary,
        "refund_reason_breakdown": reason_rows,
        "root_cause_breakdown": root_rows,
        "investigation_products": investigation_rows,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    result = analyze(args.input_dir.resolve())
    args.output.resolve().parent.mkdir(parents=True, exist_ok=True)
    args.output.resolve().write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


if __name__ == "__main__":
    main()
