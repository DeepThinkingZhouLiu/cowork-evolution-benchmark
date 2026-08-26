#!/usr/bin/env python3
"""对本题 Golden 执行六类可复现负例，确认 Judge 能拒绝。"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True

from openpyxl import load_workbook
from pptx import Presentation

from judge import OUTPUT_PROCESSOR, OUTPUT_PPTX, OUTPUT_XLSX, evaluate


def find_text_shape(slide, text: str):
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False) and shape.text.strip() == text:
            return shape
    raise ValueError(f"未找到文本：{text}")


def missing_file(case: Path) -> None:
    (case / OUTPUT_XLSX).unlink()


def wrong_fact(case: Path) -> None:
    path = case / OUTPUT_PPTX
    presentation = Presentation(path)
    find_text_shape(
        presentation.slides[9],
        "Match five stages of the story with frames",
    ).text = "Match four stages of the story with frames"
    presentation.save(path)


def hardcoded_processor(case: Path) -> None:
    script = """#!/usr/bin/env python3
import argparse
from pathlib import Path
p=argparse.ArgumentParser()
p.add_argument('--requirements')
p.add_argument('--source-pptx')
p.add_argument('--output', type=Path, required=True)
a=p.parse_args()
a.output.write_bytes(Path(__file__).with_name('revision_plan.json').read_bytes())
"""
    (case / OUTPUT_PROCESSOR).write_text(script, encoding="utf-8")


def omitted_review_comment(case: Path) -> None:
    path = case / OUTPUT_XLSX
    workbook = load_workbook(path)
    sheet = workbook["响应记录"]
    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row, 1).value == "R11":
            sheet.delete_rows(row)
            break
    workbook.save(path)


def cross_file_conflict(case: Path) -> None:
    path = case / OUTPUT_XLSX
    workbook = load_workbook(path)
    sheet = workbook["响应记录"]
    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row, 1).value == "R03":
            sheet.cell(row, 2).value = "4"
            break
    workbook.save(path)


def unsupported_assertion(case: Path) -> None:
    path = case / OUTPUT_XLSX
    workbook = load_workbook(path)
    sheet = workbook["响应记录"]
    sheet.append(["R13", "1", "已获得校长批准", "已完成", "pptx:slide=1;check=approved"])
    workbook.save(path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--golden-dir", type=Path, required=True)
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    arguments = parser.parse_args()
    mutators = {
        "missing_file": missing_file,
        "wrong_fact": wrong_fact,
        "hardcoded_processor": hardcoded_processor,
        "omitted_review_comment": omitted_review_comment,
        "cross_file_conflict": cross_file_conflict,
        "unsupported_assertion": unsupported_assertion,
    }
    results: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="officeval-080-negatives-") as temporary:
        root = Path(temporary)
        for name, mutator in mutators.items():
            case = root / name
            shutil.copytree(arguments.golden_dir, case)
            mutator(case)
            outcome = evaluate(case, arguments.input_dir, run_render=False)
            rejected = not outcome["passed"] and outcome["score"] < 1.0
            results.append(
                {
                    "case": name,
                    "rejected": rejected,
                    "score": outcome["score"],
                    "failed_checks": [
                        check["id"]
                        for check in outcome["checks"]
                        if not check["passed"]
                    ],
                }
            )
    payload = {
        "passed": all(item["rejected"] for item in results),
        "cases": results,
    }
    arguments.result.parent.mkdir(parents=True, exist_ok=True)
    arguments.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    if not payload["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
