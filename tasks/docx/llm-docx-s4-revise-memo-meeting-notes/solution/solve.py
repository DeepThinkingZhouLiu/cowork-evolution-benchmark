#!/usr/bin/env python3
"""LLM-DOCX-Editing S4 的逐题参考解法。

替换内容逐段取自上游公开、passed=true 的 S4 编辑结果；这个脚本只用于
从该题固定输入重放答案，不是通用改写器，也不用于批量生成其他任务。
"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document


INPUT_NAME = "legal-memo.docx"
OUTPUT_NAME = "legal-memo-revised.docx"

# 段落编号和内容由该题原输入与上游公开输出逐段 diff 得到。
REPLACEMENTS = {
    6: (
        "This memorandum analyzes the Company's legal options following the data breach by "
        "Westfield Technologies Inc. (the \"Respondent\"). After reviewing the applicable "
        "statutes, contractual provisions, and available evidence, we recommend pursuing "
        "Option B: regulatory filing with the Federal Trade Commission and California Attorney "
        "General's Office. This approach leverages government enforcement resources and aligns "
        "with current regulatory enforcement trends while maintaining federal litigation as a "
        "fallback option."
    ),
    7: (
        "The estimated timeline for Option B is completion by Q3 2026, with an estimated budget "
        "of $200,000 to $350,000 for regulatory filings and coordination. Federal litigation "
        "remains available as a fallback option if regulatory action proves insufficient."
    ),
    17: (
        "The regulatory approach involves filing complaints with the Federal Trade Commission "
        "(FTC) and the California Attorney General's Office, leveraging recent enforcement trends "
        "that favor aggressive action against data privacy violations. The FTC has significantly "
        "increased enforcement actions under Section 5 of the FTC Act for deceptive data "
        "practices, with recent settlements in the hundreds of millions for similar breaches. "
        "Notable precedents include the $5 billion Facebook settlement (2019) and the $170 million "
        "YouTube settlement (2019) for privacy violations."
    ),
    18: (
        "The California AG has demonstrated particularly robust enforcement under the CCPA, with "
        "recent actions resulting in substantial penalties and operational reforms. This approach "
        "has manageable direct costs (estimated at $200,000 to $350,000) and regulatory "
        "investigations typically conclude within 12-18 months. While monetary penalties flow to "
        "government entities, regulatory action often triggers private settlements and creates "
        "powerful leverage for subsequent negotiations. Additionally, government enforcement "
        "establishes public precedent that deters similar conduct industry-wide."
    ),
    20: (
        "Both options have merit, but they serve different strategic objectives. Regulatory action "
        "leverages government resources and enforcement momentum while establishing broader "
        "deterrence at lower cost. Litigation provides direct compensation and control but "
        "requires significantly higher investment. Given current regulatory enforcement trends "
        "and the strength of our evidence, regulatory action offers the most efficient path "
        "forward while preserving litigation options."
    ),
    23: (
        "Based on our analysis and the client meeting of March 12, 2026, we recommend pursuing "
        "Option B (regulatory filing) as the primary enforcement strategy. We will file coordinated "
        "complaints with the FTC and California Attorney General's Office to maximize regulatory "
        "pressure and leverage current enforcement trends. The estimated budget for this approach "
        "is $200,000 to $350,000, well within the approved budget ceiling of $500,000, with an "
        "expected timeline of completion by Q3 2026. Federal litigation remains available as a "
        "fallback option if regulatory action proves insufficient."
    ),
    24: (
        "We are prepared to begin drafting the regulatory complaints immediately upon your "
        "authorization. Please let us know how you would like to proceed."
    ),
}


def replace_paragraph(paragraph, text: str) -> None:
    """替换正文，同时保留段落样式和段落级版式。"""

    paragraph.clear()
    paragraph.add_run(text)


def solve(input_path: Path, output_path: Path) -> None:
    document = Document(input_path)
    if len(document.paragraphs) != 25:
        raise ValueError(f"S4 固定输入应有 25 段，实际为 {len(document.paragraphs)} 段")
    if document.paragraphs[0].text.strip() != "MEMORANDUM":
        raise ValueError("输入不是预期的 S4 法律备忘录")
    for index, text in REPLACEMENTS.items():
        replace_paragraph(document.paragraphs[index], text)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir / INPUT_NAME, args.output_dir / OUTPUT_NAME)


if __name__ == "__main__":
    main()
