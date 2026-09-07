#!/usr/bin/env python3
"""Build calibration cases unique to the editable OCR reconstruction task."""

from __future__ import annotations

import argparse
import base64
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches
from lxml import etree


OUTPUT_NAME = "homework.docx"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
NS = {"w": W_NS, "m": M_NS, "cp": CP_NS}


def clean(value: object) -> str:
    return " ".join(str(value or "").split()).strip()


def table(document: Document, header: list[str]):
    matches = [item for item in document.tables if [clean(cell.text) for cell in item.rows[0].cells] == header]
    if len(matches) != 1:
        raise ValueError(f"Expected one table {header}; got {len(matches)}")
    return matches[0]


def paragraph(document: Document, exact: str):
    matches = [item for item in document.paragraphs if clean(item.text) == clean(exact)]
    if len(matches) != 1:
        raise ValueError(f"Expected one paragraph {exact!r}; got {len(matches)}")
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


def corrupt_segment_sha(name: str, payload: bytes) -> bytes:
    if name != "word/document.xml":
        return payload
    original = b"571be5a360ea"
    if payload.count(original) != 1:
        raise ValueError("Question 2 segment fingerprint not unique")
    return payload.replace(original, b"000000000000")


def wrong_equation(name: str, payload: bytes) -> bytes:
    if name != "word/document.xml":
        return payload
    root = etree.fromstring(payload)
    values = root.xpath("//m:oMath//m:t", namespaces=NS)
    if len(values) != 1:
        raise ValueError("Native equation not unique")
    values[0].text = "11 * 1.75 = 19.25ns < 14 * 1.21 = 16.94ns"
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def drop_property(name: str, payload: bytes) -> bytes:
    if name != "docProps/custom.xml":
        return payload
    root = etree.fromstring(payload)
    for prop in root.xpath("//cp:property[@name='StudentID']", namespaces=NS):
        prop.getparent().remove(prop)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def break_link(name: str, payload: bytes) -> bytes:
    if name != "word/document.xml":
        return payload
    root = etree.fromstring(payload)
    links = root.xpath("//w:hyperlink[@w:anchor='question_2']", namespaces=NS)
    if len(links) != 1:
        raise ValueError("Question 2 link not unique")
    links[0].set(qn("w:anchor"), "missing_question_2")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def drop_repeating_header(name: str, payload: bytes) -> bytes:
    if name != "word/document.xml":
        return payload
    root = etree.fromstring(payload)
    markers = root.xpath("//w:tblHeader", namespaces=NS)
    if len(markers) < 7:
        raise ValueError("Expected seven repeating headers")
    markers[0].getparent().remove(markers[0])
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--golden", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=True)

    wrong_identity = Document(args.golden)
    paragraph(wrong_identity, "Ryan Wang ID:A012345").text = "Ryan Wang ID:A01234S"
    save(wrong_identity, args.output_root, "wrong_student_id_character")

    missing_bullet = Document(args.golden)
    target = paragraph(missing_bullet, "Homework is to be done individually.")
    target._p.getparent().remove(target._p)
    save(missing_bullet, args.output_root, "missing_instruction_bullet")

    wrong_question = Document(args.golden)
    q2 = next(item for item in wrong_question.paragraphs if clean(item.text).startswith("2. A modified 7-stage"))
    for run in q2.runs:
        if "M1 and M2" in run.text:
            run.text = run.text.replace("M1 and M2", "Ml and M2")
    save(wrong_question, args.output_root, "misread_m1_as_ml")

    mutate_zip(args.golden, args.output_root / "wrong_native_equation_relation" / OUTPUT_NAME, wrong_equation)

    wrong_stage = Document(args.golden)
    pipeline = table(wrong_stage, ["Instruction", *[str(index) for index in range(1, 16)]])
    pipeline.rows[2].cells[4].text = ""
    save(wrong_stage, args.output_root, "missing_decode_stall_cell")

    wrong_arrow = Document(args.golden)
    pipeline = table(wrong_arrow, ["Instruction", *[str(index) for index in range(1, 16)]])
    pipeline.rows[3].cells[8].text = "W"
    save(wrong_arrow, args.output_root, "missing_store_dependency_arrow")

    missing_instruction = Document(args.golden)
    remove_row(table(missing_instruction, ["Instruction", *[str(index) for index in range(1, 16)]]).rows[-1])
    save(missing_instruction, args.output_root, "missing_pipeline_instruction")

    mutate_zip(args.golden, args.output_root / "wrong_segment_fingerprint" / OUTPUT_NAME, corrupt_segment_sha)

    missing_segment = Document(args.golden)
    ledger = table(missing_segment, ["Key", "Role", "Chars", "SHA-12", "Text locator", "Scan location"])
    remove_row(ledger.rows[14])
    save(missing_segment, args.output_root, "missing_question_evidence_row")

    wrong_conflict = Document(args.golden)
    conflicts = table(wrong_conflict, ["Issue", "Source evidence", "Required treatment"])
    conflicts.rows[2].cells[2].text = "Reconstruct the implied seven-stage diagram."
    save(wrong_conflict, args.output_root, "invented_missing_diagram_policy")

    wrong_convention = Document(args.golden)
    conventions = table(wrong_convention, ["Token family", "Verified rendering", "Disambiguation"])
    conventions.rows[2].cells[1].text = "F1 / F2 / Ml / M2"
    save(wrong_convention, args.output_root, "wrong_digit_letter_convention")

    medium = Document(args.golden)
    closure = table(medium, ["Control", "Result", "Evidence / rule"])
    remove_row(closure.rows[-1])
    save(medium, args.output_root, "medium_missing_custom_property_closure")

    mutate_zip(args.golden, args.output_root / "missing_student_id_property" / OUTPUT_NAME, drop_property)
    mutate_zip(args.golden, args.output_root / "broken_question_navigation" / OUTPUT_NAME, break_link)
    mutate_zip(args.golden, args.output_root / "missing_repeating_header" / OUTPUT_NAME, drop_repeating_header)

    invented = Document(args.golden)
    invented.add_paragraph("The missing Question 2 diagram has been reconstructed and approved for submission.")
    invented.add_table(rows=2, cols=3)
    save(invented, args.output_root, "fabricated_question2_diagram")

    unexpected = Document(args.golden)
    save(unexpected, args.output_root, "unexpected_extra_docx")
    Document().save(args.output_root / "unexpected_extra_docx" / "extra.docx")

    flattened = Document()
    flattened.add_heading("CSE 240A Homework 2", level=0)
    pixel = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
    with tempfile.TemporaryDirectory(prefix="ocr-flat-") as temp_dir:
        image = Path(temp_dir) / "scan.png"
        image.write_bytes(pixel)
        flattened.add_picture(str(image), width=Inches(7), height=Inches(8.8))
        save(flattened, args.output_root, "flattened_scan_only")

    corrupt = args.output_root / "corrupt_docx"
    corrupt.mkdir(exist_ok=True)
    (corrupt / OUTPUT_NAME).write_bytes(b"not a DOCX")
    (args.output_root / "empty_output").mkdir(exist_ok=True)


if __name__ == "__main__":
    main()
