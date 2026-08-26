#!/usr/bin/env python3
"""Cowork-Bench MOEX 行业轮动题的逐题参考解法。"""

from __future__ import annotations

import argparse
import html
import json
import re
from pathlib import Path

from openpyxl import Workbook


SYMBOLS = ["SBER.ME", "GAZP.ME", "LKOH.ME", "MGNT.ME", "MTSS.ME"]


def parse_portal(path: Path):
    source = path.read_text(encoding="utf-8")
    rows = [
        [html.unescape(re.sub(r"<[^>]+>", "", cell)).strip() for cell in re.findall(r"<td>(.*?)</td>", row, flags=re.S)]
        for row in re.findall(r"<tr>(.*?)</tr>", source, flags=re.S)
    ]
    rows = [row for row in rows if row]
    benchmarks = {row[0]: float(row[1]) for row in rows if len(row) == 2}
    ratings = {row[0]: {"sector": row[1], "rating": row[2], "target": float(row[3])} for row in rows if len(row) == 4}
    return benchmarks, ratings


def calculate(input_dir: Path):
    snapshot = json.loads((input_dir / "moex_snapshot.json").read_text(encoding="utf-8"))
    benchmarks, ratings = parse_portal(input_dir / "research_portal.html")
    prices = {(row["symbol"], row["date"]): float(row["close"]) for row in snapshot["prices"]}
    statements = {row["symbol"]: row for row in snapshot["annual_income_statements"]}
    performance = []
    financials = []
    for symbol in SYMBOLS:
        baseline = round(prices[(symbol, "2026-02-25")], 4)
        current = round(prices[(symbol, "2026-05-26")], 2)
        sector = ratings[symbol]["sector"]
        returned = round((current - baseline) / baseline * 100, 2)
        benchmark = benchmarks[sector]
        alpha = round(returned - benchmark, 2)
        target = ratings[symbol]["target"]
        upside = round((target - current) / current * 100, 2)
        performance.append((symbol, sector, baseline, current, returned, benchmark, alpha, ratings[symbol]["rating"], int(target), upside))
        statement = statements[symbol]
        revenue = round(float(statement["total_revenue"]) / 1_000_000, 2)
        net_income = round(float(statement["net_income"]) / 1_000_000, 2)
        margin = round(net_income / revenue * 100, 2)
        financials.append((symbol, revenue, net_income, margin))
    returns = {row[0]: row[4] for row in performance}
    alphas = {row[0]: row[6] for row in performance}
    summary = [
        ("Best_Performer", max(returns, key=returns.get)),
        ("Worst_Performer", min(returns, key=returns.get)),
        ("Avg_Alpha", round(sum(alphas.values()) / len(alphas), 2)),
        ("Stocks_Above_Benchmark", sum(value > 0 for value in alphas.values())),
        ("Stocks_Below_Benchmark", sum(value < 0 for value in alphas.values())),
    ]
    return performance, financials, summary


def solve(input_dir: Path, output_dir: Path) -> None:
    performance, financials, summary = calculate(input_dir)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Performance"
    sheet.append(["Symbol", "Sector", "Price_1Y_Ago", "Current_Price", "Return_Pct", "Benchmark_Return_Pct", "Alpha", "Analyst_Rating", "Target_Price", "Upside_Pct"])
    for row in performance:
        sheet.append(row)
    sheet = workbook.create_sheet("Financials")
    sheet.append(["Symbol", "Revenue_Latest_Q", "Net_Income_Latest_Q", "Profit_Margin_Pct"])
    for row in financials:
        sheet.append(row)
    sheet = workbook.create_sheet("Summary")
    for row in summary:
        sheet.append(row)
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook.save(output_dir / "sector_rotation_report.xlsx")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
