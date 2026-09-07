#!/usr/bin/env python3
"""Construct calibration failures for the 123-row CS161 verification register."""

from __future__ import annotations

import argparse
import base64
import shutil
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Inches
from lxml import etree


OUTPUT = Path("CS161/students.docx")
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"cp": CP_NS, "w": W_NS}
ROSTER_HEADER = ["Row", "Student ID", "Last", "First", "Standing", "HW1", "HW2", "Midterm", "HW3", "Final"]
QUEUE_HEADER = ["Rank", "Row", "Student ID", "Name", "Hits", "Mean", "Spread", "Final−Mid", "Rule codes"]


def clean(value: object) -> str:
    return " ".join(str(value if value is not None else "").split()).strip()


def clone(golden: Path, root: Path, name: str) -> Path:
    target = root / name
    (target / OUTPUT.parent).mkdir(parents=True, exist_ok=True)
    shutil.copy2(golden / OUTPUT, target / OUTPUT)
    return target


def find_table(document: Document, header: list[str], occurrence: int = 0):
    matches = [table for table in document.tables if [clean(cell.text) for cell in table.rows[0].cells] == header]
    if len(matches) <= occurrence:
        raise ValueError(f"table not found: {header}; occurrence={occurrence}; count={len(matches)}")
    return matches[occurrence]


def swap_rows(table, first: int, second: int) -> None:
    a = [cell.text for cell in table.rows[first].cells]
    b = [cell.text for cell in table.rows[second].cells]
    for cell, value in zip(table.rows[first].cells, b):
        cell.text = value
    for cell, value in zip(table.rows[second].cells, a):
        cell.text = value


def mutate_zip(path: Path, mutation) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = {info.filename: (info, archive.read(info.filename)) for info in archive.infolist()}
    mutation(entries)
    with tempfile.TemporaryDirectory(prefix=".cs161-negative-", dir=path.parent) as temp_dir:
        target = Path(temp_dir) / path.name
        with zipfile.ZipFile(target, "w") as outgoing:
            for name in sorted(entries):
                info, payload = entries[name]
                replacement = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
                replacement.compress_type = zipfile.ZIP_DEFLATED
                replacement.external_attr = info.external_attr
                outgoing.writestr(replacement, payload)
        target.replace(path)


def change_property(entries) -> None:
    root = etree.fromstring(entries["docProps/custom.xml"][1])
    node = root.xpath("//cp:property[@name='VerificationCount']", namespaces=NS)[0]
    next(iter(node)).text = "55"
    entries["docProps/custom.xml"] = (entries["docProps/custom.xml"][0], etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"))


def break_navigation(entries) -> None:
    root = etree.fromstring(entries["word/document.xml"][1])
    node = root.xpath("//w:hyperlink[@w:anchor='queue_2']", namespaces=NS)[0]
    node.set(f"{{{W_NS}}}anchor", "missing_queue_2")
    entries["word/document.xml"] = (entries["word/document.xml"][0], etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--golden", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=True)

    (args.output_root / "empty_output").mkdir(exist_ok=True)
    corrupt = clone(args.golden, args.output_root, "corrupt_docx")
    (corrupt / OUTPUT).write_bytes(b"not a DOCX")

    alternate = args.output_root / "wrong_output_path"
    alternate.mkdir(exist_ok=True)
    shutil.copy2(args.golden / OUTPUT, alternate / "students.docx")

    extra = clone(args.golden, args.output_root, "unexpected_extra_docx")
    Document().save(extra / "extra.docx")

    flattened = args.output_root / "flattened_roster"
    (flattened / OUTPUT.parent).mkdir(parents=True, exist_ok=True)
    pixel = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
    with tempfile.TemporaryDirectory(prefix="cs161-flat-") as temp_dir:
        image = Path(temp_dir) / "page.png"
        image.write_bytes(pixel)
        document = Document()
        document.add_heading("CS161 Roster Verification Register", 0)
        document.add_picture(str(image), width=Inches(7), height=Inches(8.5))
        document.save(flattened / OUTPUT)

    missing = clone(args.golden, args.output_root, "missing_roster_row")
    document = Document(missing / OUTPUT)
    table = find_table(document, ROSTER_HEADER, 2)
    table._tbl.remove(table.rows[-1]._tr)
    document.save(missing / OUTPUT)

    duplicate = clone(args.golden, args.output_root, "duplicate_roster_row")
    document = Document(duplicate / OUTPUT)
    table = find_table(document, ROSTER_HEADER, 4)
    for destination, source in zip(table.rows[-1].cells, table.rows[-2].cells):
        destination.text = source.text
    document.save(duplicate / OUTPUT)

    wrong_id = clone(args.golden, args.output_root, "wrong_student_id")
    document = Document(wrong_id / OUTPUT)
    find_table(document, ROSTER_HEADER, 1).rows[1].cells[1].text = "S9999"
    document.save(wrong_id / OUTPUT)

    swapped = clone(args.golden, args.output_root, "swapped_source_order")
    document = Document(swapped / OUTPUT)
    swap_rows(find_table(document, ROSTER_HEADER, 3), 1, 2)
    document.save(swapped / OUTPUT)

    wrong_standing = clone(args.golden, args.output_root, "wrong_academic_standing")
    document = Document(wrong_standing / OUTPUT)
    find_table(document, ROSTER_HEADER, 0).rows[5].cells[4].text = "Freshman"
    document.save(wrong_standing / OUTPUT)

    wrong_score = clone(args.golden, args.output_root, "wrong_source_score")
    document = Document(wrong_score / OUTPUT)
    find_table(document, ROSTER_HEADER, 0).rows[2].cells[9].text = "20"
    document.save(wrong_score / OUTPUT)

    missing_queue = clone(args.golden, args.output_root, "missing_queue_student")
    document = Document(missing_queue / OUTPUT)
    table = find_table(document, QUEUE_HEADER, 1)
    table._tbl.remove(table.rows[-1]._tr)
    document.save(missing_queue / OUTPUT)

    wrong_order = clone(args.golden, args.output_root, "wrong_queue_order")
    document = Document(wrong_order / OUTPUT)
    swap_rows(find_table(document, QUEUE_HEADER, 0), 1, 2)
    document.save(wrong_order / OUTPUT)

    for name, column, value in (
        ("wrong_queue_rule_code", 8, "PERFECT"),
        ("wrong_queue_hit_count", 4, "1"),
        ("wrong_queue_mean", 5, "99.9"),
        ("wrong_queue_spread", 6, "96"),
        ("wrong_queue_delta", 7, "+73"),
    ):
        target = clone(args.golden, args.output_root, name)
        document = Document(target / OUTPUT)
        find_table(document, QUEUE_HEADER, 0).rows[1].cells[column].text = value
        document.save(target / OUTPUT)

    wrong_overall = clone(args.golden, args.output_root, "wrong_overall_mean")
    document = Document(wrong_overall / OUTPUT)
    find_table(document, ["Assessment", "Mean", "Median", "Minimum", "Maximum", "Zero count", "100 count"]).rows[1].cells[1].text = "46.26"
    document.save(wrong_overall / OUTPUT)

    wrong_standing_summary = clone(args.golden, args.output_root, "wrong_standing_flagged_count")
    document = Document(wrong_standing_summary / OUTPUT)
    find_table(document, ["Standing", "Count", "HW1 mean", "HW2 mean", "Midterm mean", "HW3 mean", "Final mean", "Flagged"]).rows[1].cells[7].text = "15"
    document.save(wrong_standing_summary / OUTPUT)

    wrong_incidence = clone(args.golden, args.output_root, "wrong_rule_incidence")
    document = Document(wrong_incidence / OUTPUT)
    find_table(document, ["Standing", "ZERO", "PERFECT", "SPREAD80", "SWING60", "Any flag"]).rows[-1].cells[3].text = "33"
    document.save(wrong_incidence / OUTPUT)

    wrong_block = clone(args.golden, args.output_root, "wrong_block_digest")
    document = Document(wrong_block / OUTPUT)
    find_table(document, ["Block", "Source rows", "Count", "First ID", "Last ID", "SHA-12"]).rows[2].cells[5].text = "000000000000"
    document.save(wrong_block / OUTPUT)

    wrong_property = clone(args.golden, args.output_root, "wrong_custom_property")
    mutate_zip(wrong_property / OUTPUT, change_property)

    broken_link = clone(args.golden, args.output_root, "broken_internal_link")
    mutate_zip(broken_link / OUTPUT, break_navigation)

    wrong_orientation = clone(args.golden, args.output_root, "wrong_register_orientation")
    document = Document(wrong_orientation / OUTPUT)
    section = document.sections[1]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    document.save(wrong_orientation / OUTPUT)

    fabricated = clone(args.golden, args.output_root, "fabricated_academic_decision")
    document = Document(fabricated / OUTPUT)
    document.add_paragraph("Students in the verification queue failed the course and must retake it.")
    document.save(fabricated / OUTPUT)

    missing_boundary = clone(args.golden, args.output_root, "missing_closure_boundary")
    document = Document(missing_boundary / OUTPUT)
    paragraphs = [paragraph for paragraph in document.paragraphs if clean(paragraph.text).startswith("Verification flags are mechanical source checks only.")]
    paragraphs[-1]._element.getparent().remove(paragraphs[-1]._element)
    document.save(missing_boundary / OUTPUT)

    medium = clone(args.golden, args.output_root, "medium_incomplete_closure")
    document = Document(medium / OUTPUT)
    reconciliation = find_table(document, ["Check", "Control", "Expected", "Observed", "Result"])
    reconciliation._tbl.remove(reconciliation.rows[-1]._tr)
    document.save(medium / OUTPUT)


if __name__ == "__main__":
    main()
