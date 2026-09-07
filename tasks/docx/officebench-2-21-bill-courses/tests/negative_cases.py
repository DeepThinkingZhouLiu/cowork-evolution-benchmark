#!/usr/bin/env python3
"""Build task-specific calibration cases for Bill's course-history audit."""

from __future__ import annotations

import argparse
import base64
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from lxml import etree


OUTPUT_NAME = "bill_course_history_audit.docx"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def clean(value: str) -> str:
    return " ".join((value or "").split()).strip()


def table(document: Document, header: list[str]):
    matches = [item for item in document.tables if [clean(cell.text) for cell in item.rows[0].cells] == header]
    if len(matches) != 1:
        raise ValueError(f"Expected one table: {header}; got {len(matches)}")
    return matches[0]


def save(document: Document, root: Path, name: str) -> None:
    folder = root / name
    folder.mkdir(parents=True, exist_ok=True)
    document.save(folder / OUTPUT_NAME)


def remove_row(row) -> None:
    row._tr.getparent().remove(row._tr)


def mutate_zip(source: Path, target: Path, mutate) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(target, "w") as outgoing:
        for info in incoming.infolist():
            outgoing.writestr(info, mutate(info.filename, incoming.read(info.filename)))


def unwrap_controls(name: str, payload: bytes) -> bytes:
    if name != "word/document.xml":
        return payload
    root = etree.fromstring(payload)
    for control in root.xpath("//w:sdt", namespaces=NS):
        parent = control.getparent()
        position = parent.index(control)
        for child in control.xpath("./w:sdtContent/*", namespaces=NS):
            parent.insert(position, child)
            position += 1
        parent.remove(control)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def lock_controls(name: str, payload: bytes) -> bytes:
    if name != "word/document.xml":
        return payload
    root = etree.fromstring(payload)
    for properties in root.xpath("//w:sdt/w:sdtPr", namespaces=NS):
        node = etree.Element(f"{{{W_NS}}}lock")
        node.set(f"{{{W_NS}}}val", "sdtLocked")
        properties.append(node)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--golden", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=True)

    ledger_header = ["Source row", "Term", "Year", "Course", "Grade", "Points", "Source cells"]
    missing = Document(args.golden)
    remove_row(table(missing, ledger_header).rows[6])
    save(missing, args.output_root, "missing_chronological_record")

    wrong_order = Document(args.golden)
    ledger = table(wrong_order, ledger_header)
    ledger.rows[2].cells[3].text, ledger.rows[3].cells[3].text = ledger.rows[3].cells[3].text, ledger.rows[2].cells[3].text
    save(wrong_order, args.output_root, "wrong_source_order")

    wrong_points = Document(args.golden)
    table(wrong_points, ledger_header).rows[1].cells[5].text = "3.0"
    save(wrong_points, args.output_root, "wrong_mapped_grade_points")

    wrong_locator = Document(args.golden)
    table(wrong_locator, ledger_header).rows[8].cells[6].text = "Sheet1!A8:D8"
    save(wrong_locator, args.output_root, "wrong_row_evidence_locator")

    wrong_term = Document(args.golden)
    term = table(wrong_term, ["Term", "Courses", "N", "Points", "Term GPA", "Cum. N", "Cum. points", "Cum. GPA"])
    term.rows[3].cells[4].text = "3.67"
    save(wrong_term, args.output_root, "wrong_term_gpa")

    wrong_cumulative = Document(args.golden)
    term = table(wrong_cumulative, ["Term", "Courses", "N", "Points", "Term GPA", "Cum. N", "Cum. points", "Cum. GPA"])
    term.rows[4].cells[6].text = "42.00"
    term.rows[4].cells[7].text = "3.50"
    save(wrong_cumulative, args.output_root, "wrong_cumulative_reconciliation")

    wrong_subject = Document(args.golden)
    subject = table(wrong_subject, ["Subject", "N", "Source-order course sequence", "Points", "GPA"])
    subject.rows[2].cells[3].text = "17.40"
    subject.rows[2].cells[4].text = "3.48"
    save(wrong_subject, args.output_root, "wrong_math_subject_summary")

    wrong_sequence = Document(args.golden)
    subject = table(wrong_sequence, ["Subject", "N", "Source-order course sequence", "Points", "GPA"])
    subject.rows[1].cells[2].text = "CS 101 → CS 103 → CS 102 → CS 104 → CS 105"
    save(wrong_sequence, args.output_root, "wrong_subject_source_sequence")

    wrong_jump = Document(args.golden)
    continuity = table(wrong_jump, ["Subject", "Course", "No.", "Prior no.", "Delta", "Classification", "Course cell"])
    continuity.rows[9].cells[4].text = "+1"
    continuity.rows[9].cells[5].text = "CONTIGUOUS +1"
    save(wrong_jump, args.output_root, "wrong_math_jump_classification")

    wrong_continuity_locator = Document(args.golden)
    continuity = table(wrong_continuity_locator, ["Subject", "Course", "No.", "Prior no.", "Delta", "Classification", "Course cell"])
    continuity.rows[15].cells[6].text = "Sheet1!C15"
    save(wrong_continuity_locator, args.output_root, "wrong_continuity_source_cell")

    wrong_distribution = Document(args.golden)
    distribution = table(wrong_distribution, ["Grade", "Points", "Count", "Share", "Point contribution"])
    distribution.rows[1].cells[2].text = "4"
    distribution.rows[1].cells[3].text = "26.7%"
    save(wrong_distribution, args.output_root, "wrong_grade_distribution")

    medium = Document(args.golden)
    closure = table(medium, ["Control", "Result", "Evidence / rule"])
    remove_row(closure.rows[-1])
    save(medium, args.output_root, "medium_missing_review_closure")

    mutate_zip(args.golden, args.output_root / "missing_native_review_controls" / OUTPUT_NAME, unwrap_controls)
    mutate_zip(args.golden, args.output_root / "locked_review_controls" / OUTPUT_NAME, lock_controls)

    invented = Document(args.golden)
    invented.add_paragraph("Bill has completed all degree requirements and is approved to graduate.")
    save(invented, args.output_root, "invented_degree_claim")

    extra = Document(args.golden)
    save(extra, args.output_root, "unexpected_extra_docx")
    Document().save(args.output_root / "unexpected_extra_docx" / "extra.docx")

    flattened = Document()
    flattened.add_heading("Bill Course History Audit Dossier", level=0)
    pixel = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
    with tempfile.TemporaryDirectory(prefix="bill-audit-flat-") as temp_dir:
        image = Path(temp_dir) / "flattened.png"
        image.write_bytes(pixel)
        flattened.add_picture(str(image), width=Inches(7), height=Inches(8.5))
        save(flattened, args.output_root, "flattened_screenshot_only")

    corrupt = args.output_root / "corrupt_docx"
    corrupt.mkdir(exist_ok=True)
    (corrupt / OUTPUT_NAME).write_bytes(b"not a DOCX")
    (args.output_root / "empty_output").mkdir(exist_ok=True)


if __name__ == "__main__":
    main()
