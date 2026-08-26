#!/usr/bin/env python3
"""从 Cowork-Bench 静态化的 Arxiv/API JSON 中重算文献筛选与趋势匹配结果。"""

from __future__ import annotations

import argparse
import json
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any


MATCH_RULES = (
    ("language model", "LLMs"),
    ("vision transformer", "Computer Vision"),
    ("reinforcement learning", "RL"),
)


def round_half_up(value: float, digits: int = 2) -> float:
    quantum = Decimal("1").scaleb(-digits)
    return float(Decimal(str(value)).quantize(quantum, rounding=ROUND_HALF_UP))


def load_json(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def validate_inputs(papers: Any, benchmark: Any) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    if not isinstance(papers, list) or not papers:
        raise ValueError("arxiv_papers.json 必须是非空数组")
    required_paper_fields = {"id", "title", "summary", "primary_category", "published"}
    ids: set[str] = set()
    for index, record in enumerate(papers):
        if not isinstance(record, dict) or not required_paper_fields.issubset(record):
            raise ValueError(f"论文记录 {index} 缺少必需字段")
        paper_id = str(record["id"])
        if paper_id in ids:
            raise ValueError(f"重复 Paper_ID: {paper_id}")
        ids.add(paper_id)

    topics = benchmark.get("trending_topics") if isinstance(benchmark, dict) else None
    if not isinstance(topics, list) or not topics:
        raise ValueError("benchmark_data.json 缺少 trending_topics")
    seen_areas: set[str] = set()
    for index, topic in enumerate(topics):
        if not isinstance(topic, dict) or "area" not in topic or "growth_pct" not in topic:
            raise ValueError(f"benchmark 记录 {index} 缺少 area/growth_pct")
        area = str(topic["area"])
        if area in seen_areas:
            raise ValueError(f"重复 benchmark area: {area}")
        float(topic["growth_pct"])
        seen_areas.add(area)
    return papers, topics


def infer_benchmark_area(record: dict[str, Any]) -> str | None:
    evidence = f"{record.get('title', '')} {record.get('summary', '')}".lower()
    for keyword, area in MATCH_RULES:
        if keyword in evidence:
            return area
    return None


def build_result(papers: list[dict[str, Any]], benchmark: dict[str, Any]) -> dict[str, Any]:
    papers, topics = validate_inputs(papers, benchmark)
    growth_by_area = {str(row["area"]): float(row["growth_pct"]) for row in topics}

    included: list[dict[str, Any]] = []
    excluded: list[dict[str, Any]] = []
    for record in papers:
        paper_id = str(record["id"])
        category = str(record["primary_category"])
        if not category.startswith("cs."):
            excluded.append(
                {
                    "Paper_ID": paper_id,
                    "Title": str(record["title"]),
                    "Primary_Category": category,
                    "Exclusion_Reason": "primary_category does not start with cs.",
                }
            )
            continue

        area = infer_benchmark_area(record)
        matched = area is not None and area in growth_by_area
        included.append(
            {
                "Paper_ID": paper_id,
                "Title": str(record["title"]),
                "Primary_Category": category,
                "Published": str(record["published"]),
                "Benchmark_Area": area if matched else None,
                "Growth_Pct": growth_by_area[area] if matched else None,
                "Match_Status": "Matched" if matched else "Unmatched",
            }
        )

    included.sort(key=lambda row: row["Paper_ID"])
    excluded.sort(key=lambda row: row["Paper_ID"])
    matched_rows = [row for row in included if row["Match_Status"] == "Matched"]
    matched_growth = [float(row["Growth_Pct"]) for row in matched_rows]
    ranked = sorted(matched_rows, key=lambda row: (-float(row["Growth_Pct"]), row["Benchmark_Area"]))

    top_area = ranked[0]["Benchmark_Area"] if ranked else None
    average_growth = round_half_up(sum(matched_growth) / len(matched_growth), 2) if matched_growth else None
    coverage = round_half_up(100 * len(matched_rows) / len(included), 2) if included else 0.0

    priority_labels = ("High", "Medium", "Monitor")
    action_labels = (
        "Prioritize an in-depth literature synthesis for this matched area.",
        "Schedule a focused comparative review for this matched area.",
        "Maintain a monitoring queue for this matched area.",
    )
    recommendations = []
    for index, row in enumerate(ranked, start=1):
        recommendations.append(
            {
                "Rank": index,
                "Priority": priority_labels[min(index - 1, len(priority_labels) - 1)],
                "Action": action_labels[min(index - 1, len(action_labels) - 1)],
                "Area": row["Benchmark_Area"],
                "Growth_Pct": float(row["Growth_Pct"]),
            }
        )

    return {
        "sources": [
            {"file": "arxiv_papers.json", "role": "local Arxiv snapshot"},
            {"file": "benchmark_data.json", "role": "external topic-growth snapshot"},
        ],
        "source_summary": {
            "Input_Papers": len(papers),
            "Relevant_Papers": len(included),
            "Excluded_Papers": len(excluded),
            "Benchmark_Topics": len(topics),
        },
        "data_analysis": included,
        "excluded_papers": excluded,
        "metrics": {
            "Total_Papers": len(included),
            "Excluded_Noise_Papers": len(excluded),
            "Benchmark_Matched_Papers": len(matched_rows),
            "Avg_Matched_Growth_Pct": average_growth,
            "Top_Growth_Area": top_area,
            "Benchmark_Coverage_Pct": coverage,
        },
        "recommendations": recommendations,
        "limitations": [
            "The inputs provide no paper citation counts.",
            "The inputs provide no paper-level relevance scores.",
            "No internal baseline is available, so a benchmark gap cannot be calculated.",
            "No time series or observation period is supplied; growth_pct is reproduced only as an external snapshot.",
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--papers", required=True, type=Path)
    parser.add_argument("--benchmark", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()

    papers = load_json(args.papers)
    benchmark = load_json(args.benchmark)
    result = build_result(papers, benchmark)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8") as handle:
        json.dump(result, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


if __name__ == "__main__":
    main()
