#!/usr/bin/env python3
"""Build task-specific calibration cases for the identity-preservation package."""

from __future__ import annotations

import argparse
import base64
import shutil
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from lxml import etree


OUTPUT_NAME = "A12345678.docx"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
NS = {"w": W_NS, "cp": CP_NS}


def clean(value: object) -> str:
    return " ".join(str(value or "").split()).strip()


def table(document: Document, header: list[str]):
    matches = [item for item in document.tables if [clean(cell.text) for cell in item.rows[0].cells] == header]
    if len(matches) != 1:
        raise ValueError(f"Expected exactly one table headed {header}; got {len(matches)}")
    return matches[0]


def source_start(document: Document) -> int:
    matches = [index for index, paragraph in enumerate(document.paragraphs) if clean(paragraph.text) == "CSE 221: Homework 1"]
    if len(matches) != 1:
        raise ValueError(f"Expected one original assignment title; got {matches}")
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


def drop_custom_property(name: str, payload: bytes) -> bytes:
    if name != "docProps/custom.xml":
        return payload
    root = etree.fromstring(payload)
    for prop in root.xpath("//cp:property[@name='PreservedParagraphs']", namespaces=NS):
        prop.getparent().remove(prop)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def break_internal_link(name: str, payload: bytes) -> bytes:
    if name != "word/document.xml":
        return payload
    root = etree.fromstring(payload)
    links = root.xpath("//w:hyperlink[@w:anchor='source_student_id']", namespaces=NS)
    if len(links) != 1:
        raise ValueError(f"Expected one internal identity link; got {len(links)}")
    links[0].set(qn("w:anchor"), "missing_student_id_target")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def corrupt_ledger_fingerprint(name: str, payload: bytes) -> bytes:
    if name != "word/document.xml":
        return payload
    original = b"0815efe9d8ce"
    if payload.count(original) != 1:
        raise ValueError(f"Expected one Hydra ledger fingerprint; got {payload.count(original)}")
    return payload.replace(original, b"000000000000")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--golden", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=True)

    identity_header = ["Field", "Extracted value", "Source / rule"]
    ledger_header = ["#", "Role", "List state", "Runs", "Chars", "SHA-12", "Text locator"]
    closure_header = ["Control", "Result", "Evidence / rule"]

    wrong_identity = Document(args.golden)
    table(wrong_identity, identity_header).rows[1].cells[1].text = "B12345678"
    save(wrong_identity, args.output_root, "wrong_extracted_student_id")

    wrong_source_text = Document(args.golden)
    start = source_start(wrong_source_text)
    wrong_source_text.paragraphs[start + 6].runs[0].text = wrong_source_text.paragraphs[start + 6].runs[0].text.replace("rights amplification", "privilege amplification", 1)
    save(wrong_source_text, args.output_root, "modified_protected_question")

    missing_source_paragraph = Document(args.golden)
    start = source_start(missing_source_paragraph)
    paragraph = missing_source_paragraph.paragraphs[start + 4]
    paragraph._p.getparent().remove(paragraph._p)
    save(missing_source_paragraph, args.output_root, "missing_grading_policy_paragraph")

    merged_identity_runs = Document(args.golden)
    start = source_start(merged_identity_runs)
    identity_paragraph = merged_identity_runs.paragraphs[start + 2]
    for run_node in list(identity_paragraph._p.findall(qn("w:r"))):
        identity_paragraph._p.remove(run_node)
    identity_paragraph.add_run("Student ID: A12345678").bold = True
    save(merged_identity_runs, args.output_root, "merged_identity_run_boundaries")

    flattened_numbering = Document(args.golden)
    start = source_start(flattened_numbering)
    question = flattened_numbering.paragraphs[start + 6]
    if question._p.pPr is None or question._p.pPr.numPr is None:
        raise ValueError("Original question unexpectedly lacks numbering")
    question._p.pPr.remove(question._p.pPr.numPr)
    save(flattened_numbering, args.output_root, "flattened_question_numbering")

    wrong_title_format = Document(args.golden)
    start = source_start(wrong_title_format)
    wrong_title_format.paragraphs[start].runs[0].bold = False
    save(wrong_title_format, args.output_root, "wrong_original_title_formatting")

    mutate_zip(args.golden, args.output_root / "wrong_ledger_sha12" / OUTPUT_NAME, corrupt_ledger_fingerprint)

    missing_ledger_row = Document(args.golden)
    remove_row(table(missing_ledger_row, ledger_header).rows[11])
    save(missing_ledger_row, args.output_root, "missing_ledger_paragraph_row")

    medium = Document(args.golden)
    remove_row(table(medium, closure_header).rows[-1])
    save(medium, args.output_root, "medium_missing_body_fingerprint_closure")

    mutate_zip(args.golden, args.output_root / "missing_custom_property" / OUTPUT_NAME, drop_custom_property)
    mutate_zip(args.golden, args.output_root / "broken_identity_navigation" / OUTPUT_NAME, break_internal_link)

    invented_answer = Document(args.golden)
    appendix = [paragraph for paragraph in invented_answer.paragraphs if clean(paragraph.text) == "Preservation Ledger & Closure"]
    if len(appendix) != 1:
        raise ValueError("Appendix title missing")
    inserted = OxmlElement("w:p")
    run = OxmlElement("w:r")
    value = OxmlElement("w:t")
    value.text = "Student answer: Hydra is secure because the assignment says so."
    run.append(value)
    inserted.append(run)
    appendix[0]._p.addprevious(inserted)
    save(invented_answer, args.output_root, "invented_student_answer")

    unexpected = Document(args.golden)
    save(unexpected, args.output_root, "unexpected_extra_docx")
    Document().save(args.output_root / "unexpected_extra_docx" / "extra.docx")

    alternate = args.output_root / "alternate_id_filename_only"
    alternate.mkdir(exist_ok=True)
    shutil.copyfile(args.golden, alternate / "B12345678.docx")

    flattened = Document()
    flattened.add_heading("Submission Identity Control Sheet", level=0)
    pixel = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
    with tempfile.TemporaryDirectory(prefix="identity-control-flat-") as temp_dir:
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
