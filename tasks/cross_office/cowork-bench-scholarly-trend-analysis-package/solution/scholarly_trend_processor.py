#!/usr/bin/env python3
"""抽取 Cowork-Bench 学术趋势快照，按领域关联并输出 JSON。"""

from __future__ import annotations

import argparse
import json
from decimal import Decimal, ROUND_HALF_UP
from html.parser import HTMLParser
from pathlib import Path


class TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.in_cell = False
        self.cell_parts: list[str] = []
        self.current_row: list[str] = []
        self.rows: list[list[str]] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() in ("td", "th"):
            self.in_cell = True
            self.cell_parts = []

    def handle_data(self, data: str) -> None:
        if self.in_cell:
            self.cell_parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        lowered = tag.lower()
        if lowered in ("td", "th") and self.in_cell:
            self.current_row.append(" ".join("".join(self.cell_parts).split()))
            self.in_cell = False
            self.cell_parts = []
        elif lowered == "tr" and self.current_row:
            self.rows.append(self.current_row)
            self.current_row = []


def parse_dashboard(path: Path) -> dict[str, dict]:
    parser = TableParser()
    parser.feed(path.read_text(encoding="utf-8"))
    required = ["Research_Area", "Publication_Count_2025", "Avg_Citations", "Trending_Score"]
    for index, row in enumerate(parser.rows):
        if row == required:
            records = {}
            for values in parser.rows[index + 1 :]:
                if len(values) != len(required):
                    break
                item = dict(zip(required, values))
                records[item["Research_Area"]] = {
                    "Publication_Count_2025": int(item["Publication_Count_2025"]),
                    "Avg_Citations": float(item["Avg_Citations"]),
                    "Trending_Score": float(item["Trending_Score"]),
                }
            if records:
                return records
    raise ValueError("未在 HTML 中找到指定学术趋势表格")


def round_integer(value: float) -> int:
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def process(papers_path: Path, dashboard_path: Path) -> dict:
    papers = json.loads(papers_path.read_text(encoding="utf-8"))
    dashboard = parse_dashboard(dashboard_path)
    rows = []
    for paper in papers:
        area = str(paper["Area"])
        if area not in dashboard:
            raise ValueError(f"面板缺少论文领域：{area}")
        benchmark = dashboard[area]
        rows.append(
            {
                "Paper_ID": str(paper["Paper_ID"]),
                "Title": str(paper["Title"]),
                "Area": area,
                "Citations": benchmark["Avg_Citations"],
                "Relevance_Score": benchmark["Trending_Score"],
            }
        )
    rows.sort(key=lambda item: item["Area"])
    top_rows = sorted(rows, key=lambda item: (-item["Relevance_Score"], -item["Citations"], item["Area"]))
    metrics = {
        "Total_Papers": len(rows),
        "Avg_Citations": round_integer(sum(item["Citations"] for item in rows) / len(rows)),
        "Top_Area": top_rows[0]["Area"],
        "Avg_Relevance": round_integer(sum(item["Relevance_Score"] for item in rows) / len(rows)),
    }
    recommendations = [
        {
            "Priority": "High",
            "Action": "Prioritize quarterly literature review and replication planning",
            "Area": top_rows[0]["Area"],
        },
        {
            "Priority": "Medium",
            "Action": "Track benchmark movement and prepare a focused methods scan",
            "Area": top_rows[1]["Area"],
        },
    ]
    return {"data_analysis": rows, "metrics": metrics, "recommendations": recommendations}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--papers", type=Path, required=True)
    parser.add_argument("--dashboard", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = process(args.papers, args.dashboard)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
