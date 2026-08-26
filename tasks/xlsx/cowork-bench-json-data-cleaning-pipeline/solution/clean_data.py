#!/usr/bin/env python3
"""清洗 Cowork-Bench 的销售与库存 API JSON，并生成可回放中间结果。"""

from __future__ import annotations

import argparse
import json
from collections import OrderedDict
from pathlib import Path
from typing import Any


def load_records(path: Path) -> list[dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    records = payload.get("data") if isinstance(payload, dict) else payload
    if not isinstance(records, list) or not all(isinstance(row, dict) for row in records):
        raise ValueError(f"{path.name} 的 data 必须是对象数组")
    return records


def clean_first(records: list[dict[str, Any]], key: str) -> list[dict[str, Any]]:
    seen: set[str] = set()
    cleaned: list[dict[str, Any]] = []
    for source in records:
        identifier = str(source.get(key, ""))
        if not identifier or identifier in seen:
            continue
        seen.add(identifier)
        row = dict(source)
        row["product_name"] = str(row.get("product_name", "")).strip().title()
        cleaned.append(row)
    return cleaned


def analyze(sales: list[dict[str, Any]], inventory: list[dict[str, Any]], raw_sales: int, raw_inventory: int) -> dict[str, Any]:
    sales_agg: OrderedDict[str, dict[str, Any]] = OrderedDict()
    for row in sales:
        name = str(row["product_name"])
        item = sales_agg.setdefault(name, {"quantity": 0, "revenue": 0.0, "prices": []})
        quantity = int(row["quantity"])
        price = float(row["unit_price"])
        item["quantity"] += quantity
        item["revenue"] += quantity * price
        item["prices"].append(price)
    sales_analysis = [
        {
            "Product_Name": name,
            "Total_Quantity_Sold": item["quantity"],
            "Total_Revenue": round(item["revenue"], 2),
            "Avg_Unit_Price": round(sum(item["prices"]) / len(item["prices"]), 2),
        }
        for name, item in sales_agg.items()
    ]

    inventory_status = []
    inventory_by_name: OrderedDict[str, dict[str, Any]] = OrderedDict()
    for row in inventory:
        name = str(row["product_name"])
        status = "LOW" if int(row["current_stock"]) <= int(row["reorder_point"]) else "OK"
        item = {
            "Product_Name": name,
            "Current_Stock": int(row["current_stock"]),
            "Reorder_Point": int(row["reorder_point"]),
            "Status": status,
        }
        inventory_status.append(item)
        inventory_by_name[name] = item

    names = list(sales_agg)
    names.extend(name for name in inventory_by_name if name not in sales_agg)
    combined = []
    for name in names:
        sales_item = sales_agg.get(name)
        inventory_item = inventory_by_name.get(name)
        combined.append({
            "Product_Name": name,
            "Total_Sold": sales_item["quantity"] if sales_item else 0,
            "Current_Stock": inventory_item["Current_Stock"] if inventory_item else 0,
            "Revenue": round(sales_item["revenue"], 2) if sales_item else 0.0,
            "Stock_Status": inventory_item["Status"] if inventory_item else "NO INVENTORY",
        })
    return {
        "summary": {
            "raw_sales_count": raw_sales,
            "unique_sales_count": len(sales),
            "sales_duplicates_removed": raw_sales - len(sales),
            "raw_inventory_count": raw_inventory,
            "unique_inventory_count": len(inventory),
            "inventory_duplicates_removed": raw_inventory - len(inventory),
            "total_quantity_sold": sum(item["quantity"] for item in sales_agg.values()),
            "total_revenue": round(sum(item["revenue"] for item in sales_agg.values()), 2),
            "low_stock_products": sum(item["Status"] == "LOW" for item in inventory_status),
        },
        "sales_analysis": sales_analysis,
        "inventory_status": inventory_status,
        "combined_view": combined,
    }


def process(sales_path: Path, inventory_path: Path) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    raw_sales = load_records(sales_path)
    raw_inventory = load_records(inventory_path)
    sales = clean_first(raw_sales, "order_id")
    inventory = clean_first(raw_inventory, "product_id")
    cleaned_sales = {"sales": sales, "count": len(sales)}
    cleaned_inventory = {"inventory": inventory, "count": len(inventory)}
    analysis = analyze(sales, inventory, len(raw_sales), len(raw_inventory))
    return cleaned_sales, cleaned_inventory, analysis


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sales", type=Path, required=True)
    parser.add_argument("--inventory", type=Path, required=True)
    parser.add_argument("--sales-output", type=Path, required=True)
    parser.add_argument("--inventory-output", type=Path, required=True)
    parser.add_argument("--analysis-output", type=Path, required=True)
    args = parser.parse_args()
    sales, inventory, analysis = process(args.sales, args.inventory)
    write_json(args.sales_output, sales)
    write_json(args.inventory_output, inventory)
    write_json(args.analysis_output, analysis)


if __name__ == "__main__":
    main()
