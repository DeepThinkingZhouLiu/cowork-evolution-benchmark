#!/usr/bin/env python3
"""Cowork-Bench 区域销售预测题的逐题参考解法。"""

from __future__ import annotations

import argparse
import json
from collections import defaultdict
from pathlib import Path

from docx import Document
from openpyxl import Workbook


REGION_ORDER = [
    "Азиатско-Тихоокеанский регион",
    "Европа",
    "Латинская Америка",
    "Ближний Восток",
    "Северная Америка",
]


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def calculate(input_dir: Path):
    historical = load_json(input_dir / "historical_sales.json")
    economics = load_json(input_dir / "economic_indicators.json")["regions"]
    by_region = defaultdict(list)
    for row in historical:
        by_region[row["region"]].append(float(row["revenue"]))

    forecasts = {}
    for region in REGION_ORDER:
        raw_average = sum(by_region[region]) / 3
        trailing_average = round(raw_average, 2)
        gdp = float(economics[region]["gdp_growth"])
        adjustment = round(1 + gdp / 100, 4)
        adjusted_monthly = round(trailing_average * adjustment, 2)
        quarterly = round(adjusted_monthly * 3, 2)
        forecasts[region] = {
            "raw_average": raw_average,
            "trailing_average": trailing_average,
            "gdp": gdp,
            "adjustment": adjustment,
            "adjusted_monthly": adjusted_monthly,
            "quarterly": quarterly,
            "unrounded_quarterly": raw_average * adjustment * 3,
        }
    return historical, forecasts


def write_workbook(output_path: Path, historical: list[dict], forecasts: dict) -> dict:
    workbook = Workbook()
    historical_sheet = workbook.active
    historical_sheet.title = "Historical"
    historical_sheet.append(["Region", "Month", "Orders", "Revenue"])
    lookup = {(row["region"], row["month"]): row for row in historical}
    for region in REGION_ORDER:
        for month in ("2025-10", "2025-11", "2025-12"):
            row = lookup[(region, month)]
            historical_sheet.append((region, month, int(row["orders"]), round(float(row["revenue"]), 2)))

    forecast_sheet = workbook.create_sheet("Forecast")
    forecast_sheet.append([
        "Region", "Trailing_3M_Avg", "GDP_Growth_Pct", "Economic_Adjustment",
        "Adjusted_Monthly_Forecast", "Next_Quarter_Forecast",
    ])
    for region in REGION_ORDER:
        item = forecasts[region]
        forecast_sheet.append((region, item["trailing_average"], item["gdp"], item["adjustment"], item["adjusted_monthly"], item["quarterly"]))

    total_q4 = round(sum(float(row["revenue"]) for row in historical), 2)
    total_q1 = round(sum(round(item["unrounded_quarterly"], 2) for item in forecasts.values()), 2)
    highest = max(forecasts, key=lambda region: forecasts[region]["unrounded_quarterly"])
    lowest = min(forecasts, key=lambda region: forecasts[region]["unrounded_quarterly"])
    highest_gdp = max(forecasts, key=lambda region: forecasts[region]["gdp"])
    summary = {
        "Total_Q4_2025_Revenue": total_q4,
        "Total_Monthly_Average": round(total_q4 / 3, 2),
        "Total_Q1_2026_Forecast": total_q1,
        "QoQ_Growth_Pct": round((total_q1 / total_q4 - 1) * 100, 1),
        "Highest_Forecast_Region": highest,
        "Highest_Forecast_Amount": round(forecasts[highest]["unrounded_quarterly"], 2),
        "Lowest_Forecast_Region": lowest,
        "Lowest_Forecast_Amount": round(forecasts[lowest]["unrounded_quarterly"], 2),
        "Highest_GDP_Growth_Region": highest_gdp,
        "Regions_Count": len(forecasts),
    }
    summary_sheet = workbook.create_sheet("Summary")
    summary_sheet.append(["Metric", "Value"])
    for key, value in summary.items():
        summary_sheet.append((key, value))
    workbook.save(output_path)
    return summary


def write_report(output_path: Path, forecasts: dict, summary: dict) -> None:
    document = Document()
    document.add_heading("Региональный прогноз продаж", 0)
    document.add_heading("Краткое резюме", level=1)
    document.add_paragraph(
        f"Прогноз совокупной выручки на I квартал 2026 года составляет ${summary['Total_Q1_2026_Forecast']:,.2f} "
        f"против ${summary['Total_Q4_2025_Revenue']:,.2f} в IV квартале 2025 года; рост равен {summary['QoQ_Growth_Pct']:.1f}%."
    )
    document.add_heading("Методология", level=1)
    document.add_paragraph(
        "Для каждого региона использована скользящая средняя выручка за октябрь–декабрь 2025 года. "
        "Месячный прогноз = средняя за 3 месяца × (1 + темп роста ВВП / 100); квартальный прогноз равен трём месячным."
    )
    document.add_heading("Региональный анализ", level=1)
    for region in REGION_ORDER:
        item = forecasts[region]
        document.add_paragraph(
            f"{region}: средняя за 3 месяца ${item['trailing_average']:,.2f}, рост ВВП {item['gdp']:.1f}%, "
            f"прогноз I квартала ${item['quarterly']:,.2f}."
        )
    document.add_heading("Ключевые допущения", level=1)
    document.add_paragraph(
        "Модель предполагает стабильные цены, продуктовый портфель, охват команды продаж и экономические условия. "
        "Она не моделирует новые промоакции, запуски или рыночные шоки и должна обновляться ежемесячно."
    )
    document.add_heading("Рекомендации", level=1)
    document.add_paragraph(
        "Увеличить коммерческий фокус на Азиатско-Тихоокеанском регионе как на лидере по прогнозу и росту ВВП. "
        "Для Европы применить таргетированные кампании, а все регионы сверять с фактом ежемесячно."
    )
    document.save(output_path)


def solve(input_dir: Path, output_dir: Path) -> None:
    historical, forecasts = calculate(input_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    summary = write_workbook(output_dir / "Sales_Forecast.xlsx", historical, forecasts)
    write_report(output_dir / "Forecast_Report.docx", forecasts, summary)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
