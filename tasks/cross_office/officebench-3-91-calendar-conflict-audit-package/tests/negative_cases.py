#!/usr/bin/env python3
"""构造并回放 OfficeBench 3-91 的任务专属负例。"""

from __future__ import annotations

import argparse
import re
import shutil
import tempfile
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

from judge import DOCX_NAME, ICS_NAME, REQUIRED_FILES, XLSX_NAME, evaluate


CASES = (
    "missing-docx",
    "formulas-flattened",
    "event-time-changed",
    "event-owner-changed",
    "internal-conflict-fabricated",
    "missing-pair",
    "wrong-pair-overlap",
    "wrong-merged-window",
    "wrong-free-window",
    "wrong-summary-total",
    "merged-ics-drop-event",
    "merged-ics-change-time",
    "merged-ics-add-event",
    "merged-ics-wrong-owner",
    "word-cross-file-conflict",
    "invented-timezone-and-meeting",
    "corrupt-xlsx",
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
    replaced = False
    for paragraph in locations:
        if old not in paragraph.text:
            continue
        replaced = True
        if any(old in run.text for run in paragraph.runs):
            for run in paragraph.runs:
                run.text = run.text.replace(old, new)
        else:
            paragraph.text = paragraph.text.replace(old, new)
    if not replaced:
        raise ValueError(f"未找到待替换文本：{old}")
    document.save(path)


def mutate_ics_event(path: Path, uid: str, old: str, new: str) -> None:
    text = path.read_text(encoding="utf-8")
    pattern = re.compile(r"BEGIN:VEVENT\r?\n(?:(?!END:VEVENT).)*?END:VEVENT\r?\n", flags=re.DOTALL)
    blocks = pattern.findall(text)
    target = next((block for block in blocks if f"UID:{uid}" in block), None)
    if target is None or old not in target:
        raise ValueError(f"未找到 ICS 事件或字段：{uid} {old}")
    replacement = target.replace(old, new)
    path.write_bytes(text.replace(target, replacement, 1).encode("utf-8"))


def mutate(case: str, case_dir: Path) -> None:
    workbook_path = case_dir / XLSX_NAME
    calendar_path = case_dir / ICS_NAME
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
    elif case == "event-time-changed":
        workbook = load_workbook(workbook_path)
        workbook["Source_Events"]["D2"] = "=DATE(2024,5,1)+TIME(8,30,0)"
        workbook.save(workbook_path)
    elif case == "event-owner-changed":
        workbook = load_workbook(workbook_path)
        workbook["Source_Events"]["B2"] = "Alice"
        workbook.save(workbook_path)
    elif case == "internal-conflict-fabricated":
        workbook = load_workbook(workbook_path)
        workbook["Source_Events"]["I2"] = "=1"
        workbook.save(workbook_path)
    elif case == "missing-pair":
        workbook = load_workbook(workbook_path)
        workbook["Pairwise_Conflicts"].delete_rows(6)
        workbook.save(workbook_path)
    elif case == "wrong-pair-overlap":
        workbook = load_workbook(workbook_path)
        workbook["Pairwise_Conflicts"]["H2"] = "=60"
        workbook.save(workbook_path)
    elif case == "wrong-merged-window":
        workbook = load_workbook(workbook_path)
        workbook["Conflict_Windows"]["C3"] = "=DATE(2024,5,1)+TIME(13,0,0)"
        workbook.save(workbook_path)
    elif case == "wrong-free-window":
        workbook = load_workbook(workbook_path)
        workbook["Joint_Free_Windows"]["C3"] = "=DATE(2024,5,1)+TIME(15,0,0)"
        workbook.save(workbook_path)
    elif case == "wrong-summary-total":
        workbook = load_workbook(workbook_path)
        workbook["Summary"]["B8"] = "=400"
        workbook.save(workbook_path)
    elif case == "merged-ics-drop-event":
        text = calendar_path.read_text(encoding="utf-8")
        pattern = re.compile(r"BEGIN:VEVENT\r?\n(?:(?!END:VEVENT).)*?UID:13e17887-2bbc-42a2-a881-5aabcdfda201@13e1\.org(?:(?!END:VEVENT).)*?END:VEVENT\r?\n", flags=re.DOTALL)
        updated, count = pattern.subn("", text, count=1)
        if count != 1:
            raise ValueError("未找到待删除的 dinner 事件")
        calendar_path.write_bytes(updated.encode("utf-8"))
    elif case == "merged-ics-change-time":
        mutate_ics_event(calendar_path, "2fc598fd-361b-46c0-8b12-426c4d7939af@2fc5.org", "DTSTART:20240501T160000Z", "DTSTART:20240501T163000Z")
    elif case == "merged-ics-add-event":
        text = calendar_path.read_text(encoding="utf-8")
        extra = (
            "BEGIN:VEVENT\r\nUID:invented@evocowork\r\nDTSTAMP:20260901T063000Z\r\n"
            "DTSTART:20240501T150000Z\r\nDTEND:20240501T153000Z\r\nSUMMARY:invented meeting\r\n"
            "X-EVOCOWORK-OWNER:Bob\r\nX-EVOCOWORK-SOURCE:Bob.ics\r\nEND:VEVENT\r\n"
        )
        calendar_path.write_bytes(text.replace("END:VCALENDAR\r\n", extra + "END:VCALENDAR\r\n").encode("utf-8"))
    elif case == "merged-ics-wrong-owner":
        mutate_ics_event(calendar_path, "7cf33664-81fa-44da-860a-6076c22bbf70@7cf3.org", "X-EVOCOWORK-OWNER:Bob", "X-EVOCOWORK-OWNER:Tom")
    elif case == "word-cross-file-conflict":
        replace_text(document_path, "420 minutes", "400 minutes")
    elif case == "invented-timezone-and-meeting":
        document = Document(document_path)
        document.add_paragraph("The schedule was converted to EST and a new meeting was scheduled in Room 7.")
        document.save(document_path)
    elif case == "corrupt-xlsx":
        workbook_path.write_bytes(b"not-an-xlsx")
    elif case == "extra-debug-file":
        (case_dir / "debug.txt").write_text("temporary", encoding="utf-8")
    elif case == "native-chart-removed":
        workbook = load_workbook(workbook_path)
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
        with tempfile.TemporaryDirectory(prefix=f"calendar-negative-{case}-") as tmp:
            case_dir = Path(tmp)
            copy_golden(args.golden, case_dir)
            mutate(case, case_dir)
            result = evaluate(case_dir)
            print(
                f"{case}\treward={result['reward']:.6f}\tpassed={result['passed']}"
                f"\thurdles={','.join(result['hurdle_failures'])}",
                flush=True,
            )
            if result["passed"] or float(result["reward"]) >= 0.8:
                failures.append({"case": case, "result": result})
    if failures:
        raise SystemExit(f"负例错误通过：{[item['case'] for item in failures]}")


if __name__ == "__main__":
    main()
