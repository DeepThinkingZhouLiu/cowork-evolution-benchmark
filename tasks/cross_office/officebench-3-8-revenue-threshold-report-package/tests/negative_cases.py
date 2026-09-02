#!/usr/bin/env python3
"""构造并回放本题的任务专属负例。"""

from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

from judge import DOCX_NAME, REQUIRED_FILES, XLSX_NAME, evaluate


CASES = (
    "missing-docx",
    "formulas-flattened",
    "source-row-changed",
    "retained-missing-year",
    "removed-wrong-shortfall",
    "summary-wrong-total",
    "threshold-rule-conflict",
    "word-wrong-retained-count",
    "word-cross-file-total-conflict",
    "unsupported-causal-claim",
    "corrupt-docx",
    "extra-debug-file",
    "native-chart-removed",
)


def copy_golden(golden: Path, case_dir: Path) -> None:
    case_dir.mkdir(parents=True, exist_ok=True)
    for name in REQUIRED_FILES:
        shutil.copy2(golden / name, case_dir / name)


def replace_text(path: Path, old: str, new: str) -> None:
    document = Document(path)
    locations = list(document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                locations.extend(cell.paragraphs)
    for paragraph in locations:
        if old not in paragraph.text:
            continue
        if any(old in run.text for run in paragraph.runs):
            for run in paragraph.runs:
                run.text = run.text.replace(old, new)
        else:
            paragraph.text = paragraph.text.replace(old, new)
    document.save(path)


def mutate(case: str, case_dir: Path) -> None:
    workbook_path = case_dir / XLSX_NAME
    document_path = case_dir / DOCX_NAME
    if case == "missing-docx":
        document_path.unlink()
    elif case == "formulas-flattened":
        formulas = load_workbook(workbook_path, data_only=False)
        values = load_workbook(workbook_path, data_only=True)
        for formula_sheet, value_sheet in zip(formulas.worksheets, values.worksheets, strict=True):
            for row in formula_sheet.iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        cell.value = value_sheet[cell.coordinate].value
        formulas.save(workbook_path)
    elif case == "source-row-changed":
        workbook = load_workbook(workbook_path, data_only=False)
        workbook["Source_Data"]["B4"] = 2_193_225
        workbook.save(workbook_path)
    elif case == "retained-missing-year":
        workbook = load_workbook(workbook_path, data_only=False)
        sheet = workbook["Retained_Years"]
        sheet.delete_rows(18, 1)
        workbook.save(workbook_path)
    elif case == "removed-wrong-shortfall":
        workbook = load_workbook(workbook_path, data_only=False)
        workbook["Removed_Years"]["C2"] = "=2000000-B3"
        workbook.save(workbook_path)
    elif case == "summary-wrong-total":
        workbook = load_workbook(workbook_path, data_only=False)
        workbook["Summary"]["B6"] = "=56100000"
        workbook.save(workbook_path)
    elif case == "threshold-rule-conflict":
        replace_text(document_path, "greater than or equal to $2,000,000", "strictly greater than $2,000,000")
    elif case == "word-wrong-retained-count":
        replace_text(document_path, "17 years", "16 years")
    elif case == "word-cross-file-total-conflict":
        replace_text(document_path, "$56,172,635.00", "$56,000,000.00")
    elif case == "unsupported-causal-claim":
        document = Document(document_path)
        document.add_paragraph("Revenue fell because of inflation and market competition.")
        document.save(document_path)
    elif case == "corrupt-docx":
        document_path.write_bytes(b"not-a-docx")
    elif case == "extra-debug-file":
        (case_dir / "debug.txt").write_text("temporary", encoding="utf-8")
    elif case == "native-chart-removed":
        workbook = load_workbook(workbook_path, data_only=False)
        workbook["Summary"]._charts = []
        workbook.save(workbook_path)
    else:
        raise ValueError(case)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--golden", type=Path, required=True)
    parser.add_argument("--case", choices=CASES, action="append", required=True)
    args = parser.parse_args()
    failures: list[dict[str, object]] = []
    for case in args.case:
        with tempfile.TemporaryDirectory(prefix=f"revenue-negative-{case}-") as tmp:
            case_dir = Path(tmp)
            copy_golden(args.golden, case_dir)
            mutate(case, case_dir)
            result = evaluate(case_dir)
            print(
                f"{case}\treward={result['reward']:.6f}\tpassed={result['passed']}"
                f"\thurdles={','.join(result['hurdle_failures'])}"
            )
            if result["passed"] or float(result["reward"]) >= 0.8:
                failures.append({"case": case, "result": result})
    if failures:
        raise SystemExit(f"负例错误通过：{[item['case'] for item in failures]}")


if __name__ == "__main__":
    main()
