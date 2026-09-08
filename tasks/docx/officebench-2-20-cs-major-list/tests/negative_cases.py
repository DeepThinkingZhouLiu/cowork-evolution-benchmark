#!/usr/bin/env python3
"""Create task-specific calibration cases for the CS-major advising docket."""

from __future__ import annotations

import argparse
import base64
import copy
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Inches
from lxml import etree


OUTPUT_NAME = "cs_major_advising_docket.docx"
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def clean(value: str) -> str:
    return " ".join((value or "").split()).strip()


def tables_with_header(document: Document, header: list[str]):
    matches = []
    for table in document.tables:
        if [clean(cell.text) for cell in table.rows[0].cells] == header:
            matches.append(table)
    if not matches:
        raise ValueError(f"Table not found: {header}")
    return matches


def save_case(document: Document, root: Path, name: str) -> Path:
    folder = root / name
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / OUTPUT_NAME
    document.save(path)
    return path


def remove_row(row) -> None:
    row._tr.getparent().remove(row._tr)


def mutate_zip(source: Path, target: Path, mutate) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(target, "w") as outgoing:
        for info in incoming.infolist():
            payload = incoming.read(info.filename)
            replacement = mutate(info.filename, payload)
            if replacement is not None:
                outgoing.writestr(info, replacement)


def strip_internal_links(name: str, payload: bytes) -> bytes:
    if name != "word/document.xml":
        return payload
    root = etree.fromstring(payload)
    for hyperlink in root.xpath("//w:hyperlink[@w:anchor]", namespaces=W_NS):
        parent = hyperlink.getparent()
        position = parent.index(hyperlink)
        for child in list(hyperlink):
            parent.insert(position, child)
            position += 1
        parent.remove(hyperlink)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def strip_comments(name: str, payload: bytes) -> bytes:
    if name == "word/comments.xml":
        root = etree.fromstring(payload)
        for comment in root.xpath("//w:comment", namespaces=W_NS):
            comment.getparent().remove(comment)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    if name == "word/document.xml":
        root = etree.fromstring(payload)
        for node in root.xpath("//w:commentRangeStart | //w:commentRangeEnd", namespaces=W_NS):
            node.getparent().remove(node)
        for reference in root.xpath("//w:commentReference", namespaces=W_NS):
            run = reference.getparent()
            run.getparent().remove(run)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    return payload


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--golden", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=True)

    foreign = Document(args.golden)
    queue = tables_with_header(foreign, ["Order", "Student ID", "Name", "Standing", "Final", "Δ M->F", "Tier", "Source row", "Evidence"])[0]
    row = queue.rows[3].cells
    for index, value in enumerate(("3", "S1002", "Gafo Pegoqeju", "Junior", "66", "+15", "Priority 3", "Sheet1!4")):
        row[index].text = value
    save_case(foreign, args.output_root, "wrong_filter_foreign_student")

    reordered = Document(args.golden)
    queue = tables_with_header(reordered, ["Order", "Student ID", "Name", "Standing", "Final", "Δ M->F", "Tier", "Source row", "Evidence"])[0]
    first, second = queue.rows[1]._tr, queue.rows[2]._tr
    table_element = queue._tbl
    table_element.replace(first, copy.deepcopy(second))
    table_element.replace(second, copy.deepcopy(first))
    save_case(reordered, args.output_root, "wrong_queue_order")

    wrong_tier = Document(args.golden)
    queue = tables_with_header(wrong_tier, ["Order", "Student ID", "Name", "Standing", "Final", "Δ M->F", "Tier", "Source row", "Evidence"])[0]
    for run in queue.rows[2].cells[6].paragraphs[0].runs:
        if clean(run.text) == "Priority 1":
            run.text = "Priority 2"
    save_case(wrong_tier, args.output_root, "wrong_tier_boundary")

    wrong_score = Document(args.golden)
    assessments = tables_with_header(wrong_score, ["Assessment", "Score", "Source cell"])
    assessments[0].rows[5].cells[1].text = "11"
    save_case(wrong_score, args.output_root, "wrong_profile_score")

    wrong_locator = Document(args.golden)
    assessments = tables_with_header(wrong_locator, ["Assessment", "Score", "Source cell"])
    assessments[1].rows[3].cells[2].text = "Sheet1!H7"
    save_case(wrong_locator, args.output_root, "wrong_source_locator")

    wrong_derived = Document(args.golden)
    derived = tables_with_header(wrong_derived, ["Derived check", "Result", "Calculation basis"])
    derived[2].rows[1].cells[1].text = "51.0"
    save_case(wrong_derived, args.output_root, "wrong_student_average")

    wrong_aggregate = Document(args.golden)
    appendix = tables_with_header(wrong_aggregate, ["Assessment", "CS N", "CS avg", "Non-CS N", "Non-CS avg", "CS - non-CS"])[0]
    appendix.rows[5].cells[5].text = "-59.8"
    save_case(wrong_aggregate, args.output_root, "wrong_cohort_difference")

    missing_evidence = Document(args.golden)
    assessments = tables_with_header(missing_evidence, ["Assessment", "Score", "Source cell"])
    remove_row(assessments[1].rows[4])
    save_case(missing_evidence, args.output_root, "missing_hw3_evidence")

    privacy = Document(args.golden)
    privacy.add_paragraph("Non-CS follow-up: S1002, Gafo Pegoqeju")
    save_case(privacy, args.output_root, "non_cs_identity_leak")

    mutate_zip(args.golden, args.output_root / "missing_internal_links" / OUTPUT_NAME, strip_internal_links)
    mutate_zip(args.golden, args.output_root / "missing_native_comments" / OUTPUT_NAME, strip_comments)

    medium = Document(args.golden)
    controls = tables_with_header(medium, ["Control", "Result", "Evidence"])[0]
    remove_row(controls.rows[5])
    save_case(medium, args.output_root, "medium_missing_tier_control_row")

    flattened = Document()
    flattened.add_heading("CS Major Advising Review Docket", level=0)
    one_pixel_png = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
    with tempfile.TemporaryDirectory(prefix="cs-major-flat-") as temp_dir:
        image = Path(temp_dir) / "flattened.png"
        image.write_bytes(one_pixel_png)
        flattened.add_picture(str(image), width=Inches(7), height=Inches(8.5))
        save_case(flattened, args.output_root, "flattened_screenshot_only")

    corrupt = args.output_root / "corrupt_docx"
    corrupt.mkdir(exist_ok=True)
    (corrupt / OUTPUT_NAME).write_bytes(b"not a DOCX")
    (args.output_root / "empty_output").mkdir(exist_ok=True)


if __name__ == "__main__":
    main()
