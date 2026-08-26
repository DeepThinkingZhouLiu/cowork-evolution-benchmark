#!/usr/bin/env python3
"""Cowork-Bench 食堂菜单题的可回放成本处理器。"""

from __future__ import annotations

import argparse
import json
import re
from datetime import date, timedelta
from html.parser import HTMLParser
from pathlib import Path


STORES = ["FreshMart", "ValueGrocer", "BulkBarn"]
ALIASES = {
    "Говядина (вырезка)": "Говядина",
    "Филе рыбы (судак/треска)": "Филе рыбы",
    "Фарш мясной (свинина+говядина)": "Фарш мясной",
}
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
WEEKLY_BUDGET = 12000.0


class TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.rows: list[list[str]] = []
        self.current_row: list[str] | None = None
        self.current_cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() == "tr":
            self.current_row = []
        elif tag.lower() in {"th", "td"} and self.current_row is not None:
            self.current_cell = []

    def handle_data(self, data: str) -> None:
        if self.current_cell is not None:
            self.current_cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        lowered = tag.lower()
        if lowered in {"th", "td"} and self.current_row is not None and self.current_cell is not None:
            self.current_row.append(" ".join("".join(self.current_cell).split()))
            self.current_cell = None
        elif lowered == "tr" and self.current_row:
            self.rows.append(self.current_row)
            self.current_row = None


def parse_number(value: str) -> float:
    match = re.search(r"-?\d+(?:[.,]\d+)?", value.replace("\xa0", " "))
    if not match:
        raise ValueError(f"无法从报价中解析数字: {value!r}")
    return float(match.group(0).replace(",", "."))


def parse_prices(path: Path) -> list[dict]:
    parser = TableParser()
    parser.feed(path.read_text(encoding="utf-8"))
    if not parser.rows:
        raise ValueError("价格 HTML 中没有表格行")
    header = parser.rows[0]
    if len(header) < 5 or header[2:5] != STORES:
        raise ValueError(f"价格表表头不符合预期: {header}")
    rows = []
    for raw in parser.rows[1:]:
        if len(raw) < 5:
            continue
        store_prices = {store: parse_number(raw[index + 2]) for index, store in enumerate(STORES)}
        minimum = min(store_prices.values())
        cheapest = next(store for store in STORES if store_prices[store] == minimum)
        rows.append(
            {
                "Item": raw[0],
                "Unit": raw[1],
                **{f"{store}_Price": store_prices[store] for store in STORES},
                "Cheapest_Store": cheapest,
                "Cheapest_Price": minimum,
            }
        )
    if len(rows) != 22 or len({row["Item"] for row in rows}) != 22:
        raise ValueError(f"预期 22 条不重复商品，实际 {len(rows)}")
    return sorted(rows, key=lambda row: row["Item"])


def analyze_recipes(recipes_path: Path, price_rows: list[dict]) -> list[dict]:
    recipes = json.loads(recipes_path.read_text(encoding="utf-8"))
    if not isinstance(recipes, list) or len(recipes) != 50:
        raise ValueError("配方库必须是上游完整 50 条快照")
    by_item = {row["Item"]: row for row in price_rows}
    candidates = []
    for recipe in recipes:
        if recipe.get("category") != "горячее":
            continue
        matched = []
        unpriced = []
        seen_items = set()
        for ingredient in recipe.get("ingredients", []):
            recipe_name = str(ingredient.get("name") or "").strip()
            price_item = ALIASES.get(recipe_name, recipe_name)
            if price_item in by_item and price_item not in seen_items:
                row = by_item[price_item]
                matched.append(
                    {
                        "Recipe_Ingredient": recipe_name,
                        "Price_Item": price_item,
                        "Cheapest_Store": row["Cheapest_Store"],
                        "Quoted_Price": row["Cheapest_Price"],
                    }
                )
                seen_items.add(price_item)
            elif price_item not in by_item:
                unpriced.append(recipe_name)
        if len(matched) < 3:
            continue
        total = round(sum(item["Quoted_Price"] for item in matched), 2)
        store_mix = [store for store in STORES if any(item["Cheapest_Store"] == store for item in matched)]
        candidates.append(
            {
                "Recipe_ID": recipe["id"],
                "Recipe": recipe["name"],
                "Matched_Ingredient_Count": len(matched),
                "Unpriced_Ingredient_Count": len(unpriced),
                "Matched_Ingredients": matched,
                "Unpriced_Ingredients": unpriced,
                "Total_Cost": total,
                "Cheapest_Store_Mix": store_mix,
                "Cost_Per_Serving": round(total / 4.0, 2),
            }
        )
    candidates.sort(key=lambda row: (row["Total_Cost"], row["Recipe_ID"]))
    if len(candidates) < 5:
        raise ValueError("符合至少 3 个已报价原料的热菜不足 5 个")
    return candidates[:5]


def build_plan(selected: list[dict]) -> list[dict]:
    plan = []
    start = date(2026, 3, 16)
    dinner = list(reversed(selected))
    for index, day in enumerate(DAYS):
        current = start + timedelta(days=index)
        for meal, recipe in (("Lunch", selected[index]), ("Dinner", dinner[index])):
            plan.append(
                {
                    "Date": current.isoformat(),
                    "Day": day,
                    "Meal": meal,
                    "Recipe_ID": recipe["Recipe_ID"],
                    "Recipe": recipe["Recipe"],
                    "Estimated_Cost": recipe["Total_Cost"],
                }
            )
    return plan


def build_budget(price_rows: list[dict], selected: list[dict], plan: list[dict]) -> dict:
    price_by_item = {row["Item"]: row for row in price_rows}
    recipe_by_id = {row["Recipe_ID"]: row for row in selected}
    weekly_total = round(sum(row["Estimated_Cost"] for row in plan), 2)
    single_store_totals = {}
    for store in STORES:
        total = 0.0
        for meal in plan:
            for item in recipe_by_id[meal["Recipe_ID"]]["Matched_Ingredients"]:
                total += price_by_item[item["Price_Item"]][f"{store}_Price"]
        single_store_totals[store] = round(total, 2)
    best_store = min(STORES, key=lambda store: (single_store_totals[store], STORES.index(store)))
    best_total = single_store_totals[best_store]
    savings = round(best_total - weekly_total, 2)
    return {
        "Weekly_Total": weekly_total,
        "Average_Daily": round(weekly_total / 5.0, 2),
        "Weekly_Budget": WEEKLY_BUDGET,
        "Budget_Headroom": round(WEEKLY_BUDGET - weekly_total, 2),
        "Best_Single_Store": best_store,
        "Best_Single_Store_Total": best_total,
        "Savings_Vs_Best_Single_Store": savings,
        "Savings_Pct": round(savings / best_total * 100.0, 2),
        "Single_Store_Totals": single_store_totals,
    }


def optimize(prices_path: Path, recipes_path: Path) -> dict:
    price_rows = parse_prices(prices_path)
    selected = analyze_recipes(recipes_path, price_rows)
    plan = build_plan(selected)
    return {
        "price_rows": price_rows,
        "selected_recipes": selected,
        "weekly_meal_plan": plan,
        "budget_summary": build_budget(price_rows, selected, plan),
        "limitations": [
            "Total_Cost is a quoted-basket cost: one quoted unit per distinct matched ingredient; no quantity conversion or imputation.",
            "Protein >=20 g per serving cannot be verified because the supplied inputs contain no nutrition data.",
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--prices", type=Path, required=True)
    parser.add_argument("--recipes", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = optimize(args.prices, args.recipes)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
