#!/usr/bin/env python3
"""Build calibration failures specific to the controlled master / QA-copy task."""

from __future__ import annotations

import argparse
import base64
import shutil
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches
from lxml import etree


MASTER = Path("homework.docx")
REVIEW = Path("homework/homework.docx")
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
NS = {"w": W_NS, "m": M_NS, "cp": CP_NS}


def clean(value: object) -> str:
    return " ".join(str(value or "").split()).strip()


def clone_pair(golden: Path, root: Path, name: str) -> Path:
    target = root / name
    (target / REVIEW.parent).mkdir(parents=True, exist_ok=True)
    shutil.copy2(golden / MASTER, target / MASTER)
    shutil.copy2(golden / REVIEW, target / REVIEW)
    return target


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


def mutate_zip(source: Path, destination: Path, mutation) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as archive:
        entries = {info.filename: (info, archive.read(info.filename)) for info in archive.infolist()}
    mutation(entries)
    fixed = (1980, 1, 1, 0, 0, 0)
    with zipfile.ZipFile(destination, "w") as outgoing:
        for name in sorted(entries):
            info, payload = entries[name]
            replacement = zipfile.ZipInfo(name, fixed)
            replacement.compress_type = zipfile.ZIP_DEFLATED
            replacement.external_attr = info.external_attr
            outgoing.writestr(replacement, payload)


def change_property(name: str, value: str):
    def mutation(entries):
        root = etree.fromstring(entries["docProps/custom.xml"][1])
        nodes = root.xpath(f"//cp:property[@name='{name}']", namespaces=NS)
        if len(nodes) != 1:
            raise ValueError(name)
        child = next(iter(nodes[0]))
        child.text = value
        entries["docProps/custom.xml"] = (entries["docProps/custom.xml"][0], etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"))
    return mutation


def remove_equation(entries) -> None:
    root = etree.fromstring(entries["word/document.xml"][1])
    node = root.xpath("//m:oMathPara", namespaces=NS)[0]
    node.getparent().remove(node)
    entries["word/document.xml"] = (entries["word/document.xml"][0], etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"))


def change_comment_author(entries) -> None:
    root = etree.fromstring(entries["word/comments.xml"][1])
    node = root.xpath("//w:comment", namespaces=NS)[2]
    node.set(qn("w:author"), "Anonymous OCR")
    node.set(qn("w:initials"), "AO")
    entries["word/comments.xml"] = (entries["word/comments.xml"][0], etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"))


def remove_comment(entries) -> None:
    root = etree.fromstring(entries["word/comments.xml"][1])
    node = root.xpath("//w:comment", namespaces=NS)[-1]
    node.getparent().remove(node)
    entries["word/comments.xml"] = (entries["word/comments.xml"][0], etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"))


def lock_control(entries) -> None:
    root = etree.fromstring(entries["word/document.xml"][1])
    node = root.xpath("//w:sdt[w:sdtPr/w:tag[@w:val='QA_Status']]/w:sdtPr", namespaces=NS)[0]
    lock = etree.SubElement(node, f"{{{W_NS}}}lock")
    lock.set(f"{{{W_NS}}}val", "sdtContentLocked")
    entries["word/document.xml"] = (entries["word/document.xml"][0], etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"))


def remove_control(entries) -> None:
    root = etree.fromstring(entries["word/document.xml"][1])
    node = root.xpath("//w:sdt[w:sdtPr/w:tag[@w:val='QA_Reviewer']]", namespaces=NS)[0]
    node.getparent().remove(node)
    entries["word/document.xml"] = (entries["word/document.xml"][0], etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"))


def break_link(entries) -> None:
    root = etree.fromstring(entries["word/document.xml"][1])
    node = root.xpath("//w:hyperlink[@w:anchor='master_handoff']", namespaces=NS)[0]
    node.set(qn("w:anchor"), "missing_handoff")
    entries["word/document.xml"] = (entries["word/document.xml"][0], etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"))


def remove_review_header(entries) -> None:
    name = next(name for name in entries if name.startswith("word/header") and name.endswith(".xml"))
    payload = entries[name][1].replace(b"QA WORKING COPY", b"UNCONTROLLED COPY")
    entries[name] = (entries[name][0], payload)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--golden", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=True)

    (args.output_root / "empty_output").mkdir(exist_ok=True)

    corrupt = clone_pair(args.golden, args.output_root, "corrupt_master")
    (corrupt / MASTER).write_bytes(b"not a DOCX")

    missing_review = clone_pair(args.golden, args.output_root, "missing_review_copy")
    (missing_review / REVIEW).unlink()

    extra = clone_pair(args.golden, args.output_root, "unexpected_extra_docx")
    Document().save(extra / "extra.docx")

    flattened = args.output_root / "flattened_pair"
    (flattened / REVIEW.parent).mkdir(parents=True, exist_ok=True)
    pixel = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
    with tempfile.TemporaryDirectory(prefix="controlled-flat-") as temp_dir:
        image = Path(temp_dir) / "scan.png"
        image.write_bytes(pixel)
        for path in (flattened / MASTER, flattened / REVIEW):
            doc = Document()
            doc.add_heading("CSE 240A Homework 2", level=0)
            doc.add_picture(str(image), width=Inches(7), height=Inches(8.8))
            doc.save(path)

    wrong_student = clone_pair(args.golden, args.output_root, "unsynchronized_review_student_id")
    doc = Document(wrong_student / REVIEW)
    target = paragraph(doc, "Ryan Wang ID:A012345")
    target.text = "Ryan Wang ID:A01234S"
    doc.save(wrong_student / REVIEW)

    fabricated = clone_pair(args.golden, args.output_root, "fabricated_review_diagram")
    doc = Document(fabricated / REVIEW)
    doc.add_paragraph("The missing Question 2 diagram was reconstructed and approved for release.")
    doc.save(fabricated / REVIEW)

    wrong_pipeline = clone_pair(args.golden, args.output_root, "wrong_master_pipeline_stage")
    doc = Document(wrong_pipeline / MASTER)
    pipeline = table(doc, ["Instruction", *[str(index) for index in range(1, 16)]])
    pipeline.rows[1].cells[3].text = "Y"
    doc.save(wrong_pipeline / MASTER)

    missing_equation = clone_pair(args.golden, args.output_root, "missing_review_native_equation")
    mutate_zip(missing_equation / REVIEW, missing_equation / REVIEW, remove_equation)

    wrong_digest = clone_pair(args.golden, args.output_root, "wrong_review_payload_digest")
    mutate_zip(wrong_digest / REVIEW, wrong_digest / REVIEW, change_property("CanonicalPayloadSHA256", "0" * 64))

    wrong_role = clone_pair(args.golden, args.output_root, "wrong_master_role_property")
    mutate_zip(wrong_role / MASTER, wrong_role / MASTER, change_property("DocumentRole", "REVIEW"))

    wrong_path = clone_pair(args.golden, args.output_root, "wrong_review_copy_path_property")
    mutate_zip(wrong_path / REVIEW, wrong_path / REVIEW, change_property("CopyPath", "homework.docx"))

    missing_block = clone_pair(args.golden, args.output_root, "missing_master_payload_block")
    doc = Document(missing_block / MASTER)
    register = table(doc, ["Block", "Role", "Chars", "SHA-12", "Verified locator"])
    register._tbl.remove(register.rows[-1]._tr)
    doc.save(missing_block / MASTER)

    wrong_exception = clone_pair(args.golden, args.output_root, "wrong_master_exception_treatment")
    doc = Document(wrong_exception / MASTER)
    exceptions = table(doc, ["ID", "Evidence", "Master treatment"])
    exceptions.rows[2].cells[2].text = "Recreate the implied second diagram."
    doc.save(wrong_exception / MASTER)

    wrong_disposition = clone_pair(args.golden, args.output_root, "wrong_review_issue_disposition")
    doc = Document(wrong_disposition / REVIEW)
    issues = table(doc, ["ID", "Issue", "Source evidence", "Required disposition"])
    issues.rows[2].cells[3].text = "Invent the missing diagram to complete the page."
    doc.save(wrong_disposition / REVIEW)

    wrong_author = clone_pair(args.golden, args.output_root, "wrong_comment_author")
    mutate_zip(wrong_author / REVIEW, wrong_author / REVIEW, change_comment_author)

    missing_comment_case = clone_pair(args.golden, args.output_root, "missing_review_comment")
    mutate_zip(missing_comment_case / REVIEW, missing_comment_case / REVIEW, remove_comment)

    locked = clone_pair(args.golden, args.output_root, "locked_qa_status_control")
    mutate_zip(locked / REVIEW, locked / REVIEW, lock_control)

    missing_control_case = clone_pair(args.golden, args.output_root, "missing_qa_reviewer_control")
    mutate_zip(missing_control_case / REVIEW, missing_control_case / REVIEW, remove_control)

    broken = clone_pair(args.golden, args.output_root, "broken_master_navigation")
    mutate_zip(broken / MASTER, broken / MASTER, break_link)

    no_header = clone_pair(args.golden, args.output_root, "missing_review_role_header")
    mutate_zip(no_header / REVIEW, no_header / REVIEW, remove_review_header)

    boundary = clone_pair(args.golden, args.output_root, "missing_reviewer_boundary_note")
    doc = Document(boundary / REVIEW)
    target = paragraph(doc, "Return the reviewed working copy alongside the unchanged controlled master.")
    target._p.getparent().remove(target._p)
    doc.save(boundary / REVIEW)

    medium = clone_pair(args.golden, args.output_root, "medium_incomplete_review_queue")
    doc = Document(medium / REVIEW)
    issues = table(doc, ["ID", "Issue", "Source evidence", "Required disposition"])
    issues._tbl.remove(issues.rows[-1]._tr)
    doc.save(medium / REVIEW)


if __name__ == "__main__":
    main()
