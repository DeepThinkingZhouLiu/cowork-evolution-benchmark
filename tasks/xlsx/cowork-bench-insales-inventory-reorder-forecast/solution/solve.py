#!/usr/bin/env python3
"""Cowork-Bench 库存快照题的逐题参考解法。"""

from __future__ import annotations

import argparse
import json
import math
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook


OUTPUT_NAME = "Inventory_Forecast.xlsx"


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def solve(input_dir: Path, output_dir: Path) -> None:
    products = load_json(input_dir / "store_products.json")
    orders = load_json(input_dir / "store_orders.json")
    lead_times = load_json(input_dir / "supplier_lead_times.json")
    snapshot_date = date.fromisoformat((input_dir / "snapshot_date.txt").read_text(encoding="utf-8").strip())

    eligible_orders = [order for order in orders if order["status"] in {"completed", "processing"}]
    order_dates = [datetime.fromisoformat(order["order_date"]).date() for order in eligible_orders]
    history_days = (max(order_dates) - min(order_dates)).days
    if history_days != 361:
        raise ValueError(f"上游固定快照应覆盖 361 天，实际为 {history_days}")

    total_sold = defaultdict(int)
    for order in eligible_orders:
        for item in order["line_items"]:
            total_sold[int(item["product_id"])] += int(item["quantity"])

    analysis_rows = []
    for product in sorted(products, key=lambda item: int(item["product_id"])):
        if product["status"] != "publish" or int(product["current_stock"]) <= 0:
            continue
        product_id = int(product["product_id"])
        stock = int(product["current_stock"])
        category = product["categories"][0]["name"]
        average = round(total_sold[product_id] / history_days, 4)
        days_until = 999 if average == 0 else round(max(stock - 5, 0) / average, 1)
        needs_reorder = total_sold[product_id] > 0 and days_until <= 90
        if needs_reorder:
            computed = snapshot_date + timedelta(days=math.floor(days_until) - int(lead_times[category]))
            reorder_date = max(snapshot_date + timedelta(days=1), computed).isoformat()
        else:
            reorder_date = "N/A"
        analysis_rows.append(
            (
                product_id,
                product["product_name"][:80],
                category,
                stock,
                average,
                days_until,
                int(lead_times[category]),
                "Yes" if needs_reorder else "No",
                reorder_date,
            )
        )

    workbook = Workbook()
    stock_sheet = workbook.active
    stock_sheet.title = "Stock Analysis"
    stock_sheet.append([
        "Product_ID", "Product_Name", "Category", "Current_Stock", "Avg_Daily_Sales",
        "Days_Until_Safety_Stock", "Lead_Time", "Needs_Reorder", "Reorder_Date",
    ])
    for row in analysis_rows:
        stock_sheet.append(row)

    schedule_sheet = workbook.create_sheet("Reorder Schedule")
    schedule_sheet.append(["Product_Name", "Category", "Current_Stock", "Reorder_Date", "Lead_Time_Days"])
    for row in analysis_rows:
        if row[7] == "Yes":
            schedule_sheet.append((row[1][:60], row[2], row[3], row[8], row[6]))

    output_dir.mkdir(parents=True, exist_ok=True)
    workbook.save(output_dir / OUTPUT_NAME)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
