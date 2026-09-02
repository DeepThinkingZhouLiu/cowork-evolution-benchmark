#!/usr/bin/env python3
"""构造并回放 OfficeBench 3-19/3-20 的任务专属负例。"""

from __future__ import annotations

import argparse
import shutil
import tempfile
from email import policy
from email.generator import BytesGenerator
from email.parser import BytesParser
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

from judge import DOCX_NAME, EML_NAME, ICS_NAME, REQUIRED_FILES, XLSX_NAME, evaluate


CASES = (
    "missing-docx",
    "formulas-flattened",
    "policy-text-changed",
    "wrong-policy-status",
    "wrong-word-count",
    "fabricated-second-reference",
    "decision-yes",
    "deadline-invented",
    "submission-on-time-invented",
    "word-decision-conflict",
    "email-wrong-recipient",
    "email-drop-attachment",
    "email-modify-attachment",
    "email-decision-conflict",
    "ics-change-time",
    "ics-add-dtend",
    "ics-add-location-attendee",
    "ics-decision-conflict",
    "corrupt-xlsx",
    "extra-debug-file",
    "native-chart-removed",
)


def copy_golden(golden: Path, case_dir: Path) -> None:
    case_dir.mkdir(parents=True, exist_ok=True)
    for filename in REQUIRED_FILES:
        shutil.copy2(golden / filename, case_dir / filename)


def replace_docx_text(path: Path, old: str, new: str) -> None:
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
        for run in paragraph.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)
        if old in paragraph.text:
            paragraph.text = paragraph.text.replace(old, new)
    if not replaced:
        raise ValueError(f"未找到 Word 待替换文本：{old}")
    document.save(path)


def save_message(message, path: Path) -> None:
    with path.open("wb") as handle:
        BytesGenerator(handle, policy=policy.SMTP, mangle_from_=False).flatten(message)


def mutate(case: str, case_dir: Path) -> None:
    workbook_path = case_dir / XLSX_NAME
    document_path = case_dir / DOCX_NAME
    email_path = case_dir / EML_NAME
    calendar_path = case_dir / ICS_NAME
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
    elif case == "policy-text-changed":
        workbook = load_workbook(workbook_path)
        workbook["Policy_Criteria"]["B2"] = "The applicant's name is optional."
        workbook.save(workbook_path)
    elif case == "wrong-policy-status":
        workbook = load_workbook(workbook_path)
        workbook["Policy_Criteria"]["D5"] = "PASS"
        workbook.save(workbook_path)
    elif case == "wrong-word-count":
        workbook = load_workbook(workbook_path)
        workbook["Application_Evidence"]["B5"] = 99
        workbook.save(workbook_path)
    elif case == "fabricated-second-reference":
        workbook = load_workbook(workbook_path)
        workbook["Application_Evidence"]["B6"] = 2
        workbook["Application_Evidence"]["B7"] = "Dr. Jane Doe, Professor, University Y"
        workbook.save(workbook_path)
    elif case == "decision-yes":
        workbook = load_workbook(workbook_path)
        workbook["Decision_Summary"]["B5"] = "YES"
        workbook.save(workbook_path)
    elif case == "deadline-invented":
        workbook = load_workbook(workbook_path)
        workbook["Application_Evidence"]["B8"] = "2024-05-15"
        workbook.save(workbook_path)
    elif case == "submission-on-time-invented":
        document = Document(document_path)
        document.add_paragraph("The application was submitted on time before the deadline.")
        document.save(document_path)
    elif case == "word-decision-conflict":
        replace_docx_text(document_path, "Decision: NO.", "Decision: YES.")
    elif case == "email-wrong-recipient":
        message = BytesParser(policy=policy.default).parse(email_path.open("rb"))
        message.replace_header("To", "Fellowship Officer <officer@school.edu>")
        save_message(message, email_path)
    elif case == "email-drop-attachment":
        message = BytesParser(policy=policy.default).parse(email_path.open("rb"))
        if not message.is_multipart() or len(message.get_payload()) < 2:
            raise ValueError("邮件没有可删除附件")
        del message.get_payload()[1:]
        save_message(message, email_path)
    elif case == "email-modify-attachment":
        message = BytesParser(policy=policy.default).parse(email_path.open("rb"))
        attachment = next(message.iter_attachments())
        attachment.set_content(b"changed-pdf", maintype="application", subtype="pdf", cte="base64")
        attachment.add_header("Content-Disposition", "attachment", filename="fellowship_application.pdf")
        save_message(message, email_path)
    elif case == "email-decision-conflict":
        raw = email_path.read_bytes()
        if b"Decision: NO." not in raw:
            raise ValueError("邮件正文中未找到决策")
        email_path.write_bytes(raw.replace(b"Decision: NO.", b"Decision: YES.", 1))
    elif case == "ics-change-time":
        raw = calendar_path.read_text(encoding="utf-8")
        calendar_path.write_bytes(raw.replace("DTSTART:20240531T130000Z", "DTSTART:20240531T140000Z").encode("utf-8"))
    elif case == "ics-add-dtend":
        raw = calendar_path.read_text(encoding="utf-8")
        calendar_path.write_bytes(raw.replace("DTSTART:20240531T130000Z\n", "DTSTART:20240531T130000Z\nDTEND:20240531T140000Z\n").encode("utf-8"))
    elif case == "ics-add-location-attendee":
        raw = calendar_path.read_text(encoding="utf-8")
        added = "LOCATION:Room 7\nATTENDEE:mailto:officer@school.edu\n"
        calendar_path.write_bytes(raw.replace("STATUS:TENTATIVE\n", "STATUS:TENTATIVE\n" + added).encode("utf-8"))
    elif case == "ics-decision-conflict":
        raw = calendar_path.read_text(encoding="utf-8")
        calendar_path.write_bytes(raw.replace("X-EVOCOWORK-DECISION:NO", "X-EVOCOWORK-DECISION:YES").encode("utf-8"))
    elif case == "corrupt-xlsx":
        workbook_path.write_bytes(b"not-an-xlsx")
    elif case == "extra-debug-file":
        (case_dir / "debug.txt").write_text("temporary", encoding="utf-8")
    elif case == "native-chart-removed":
        workbook = load_workbook(workbook_path)
        workbook["Decision_Summary"]._charts = []
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
        with tempfile.TemporaryDirectory(prefix=f"fellowship-negative-{case}-") as temporary:
            case_dir = Path(temporary)
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
