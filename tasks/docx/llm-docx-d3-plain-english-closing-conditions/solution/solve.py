#!/usr/bin/env python3
"""LLM-DOCX-Editing D3 的逐题参考解法。"""

from __future__ import annotations

import argparse
from pathlib import Path

from docx import Document


INPUT_NAME = "cap-conditional-edit.docx"
OUTPUT_NAME = "cap-conditional-edit-plain-english.docx"
PLAIN_ENGLISH_CLOSING_CONDITIONS = (
    "Before the sale can be completed, the following conditions must be met: "
    "(a) all the promises and statements made by both parties must still be true and accurate, "
    "(b) government antitrust approval must be obtained, "
    "(c) no significant negative changes to the business can have occurred, "
    "(d) important third-party approvals listed in Schedule 3 must be received, and "
    "(e) the Buyer must have secured its financing."
)


def solve(input_path: Path, output_path: Path) -> None:
    document = Document(input_path)
    paragraphs = document.paragraphs
    if len(paragraphs) != 22 or paragraphs[16].text.strip() != "Closing Conditions":
        raise ValueError("输入不是预期的 LLM-DOCX-Editing D3 fixture")
    paragraphs[17].clear()
    paragraphs[17].add_run(PLAIN_ENGLISH_CLOSING_CONDITIONS)
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
