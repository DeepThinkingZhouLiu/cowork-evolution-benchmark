#!/usr/bin/env python3
"""Create task-specific calibration cases for the CS161 roster control book."""

from __future__ import annotations

import argparse
import base64
import copy
import shutil
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Inches
from lxml import etree


W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def norm(value: str) -> str:
    return " ".join((value or "").split()).strip().lower()


def find_table(document, headers: list[str]):
    wanted = [value.lower() for value in headers]
    for table in document.tables:
        actual = [norm(cell.text) for cell in table.rows[0].cells]
        if actual == wanted:
            return table
    raise ValueError(f"Table not found: {headers}")


def save_case(document, root: Path, name: str) -> Path:
    folder = root / name
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "cs161_roster_control_book.docx"
    document.save(path)
    return path


def remove_row(row) -> None:
    row._element.getparent().remove(row._element)


def strip_internal_links(source: Path, target: Path) -> None:
    with zipfile.ZipFile(source) as input_zip, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as output_zip:
        for info in input_zip.infolist():
            payload = input_zip.read(info.filename)
            if info.filename == "word/document.xml":
                root = etree.fromstring(payload)
                for hyperlink in root.xpath("//w:hyperlink[@w:anchor]", namespaces=W_NS):
                    hyperlink.attrib.pop("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor", None)
                payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            output_zip.writestr(info, payload)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--golden", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=True)

    changed = Document(args.golden)
    roster = find_table(changed, ["Source row", "Student ID", "Last name", "First name", "Academic standing", "Major", "HW1", "HW2", "Midterm", "HW3", "Final"])
    roster.rows[7].cells[10].text = "98"
    save_case(changed, args.output_root, "wrong_student_score")

    reordered = Document(args.golden)
    roster = find_table(reordered, ["Source row", "Student ID", "Last name", "First name", "Academic standing", "Major", "HW1", "HW2", "Midterm", "HW3", "Final"])
    first = [cell.text for cell in roster.rows[1].cells]
    second = [cell.text for cell in roster.rows[2].cells]
    for cell, value in zip(roster.rows[1].cells, second):
        cell.text = value
    for cell, value in zip(roster.rows[2].cells, first):
        cell.text = value
    save_case(reordered, args.output_root, "wrong_source_order")

    wrong_summary = Document(args.golden)
    distribution = find_table(wrong_summary, ["Dimension", "Category", "Count", "Share"])
    distribution.rows[1].cells[2].text = "4"
    distribution.rows[1].cells[3].text = "44.4%"
    save_case(wrong_summary, args.output_root, "wrong_distribution")

    missing_assessment = Document(args.golden)
    assessment = find_table(missing_assessment, ["Assessment", "Valid N", "Average", "Minimum", "Maximum"])
    remove_row(assessment.rows[-1])
    save_case(missing_assessment, args.output_root, "missing_final_reconciliation")

    extra = Document(args.golden)
    roster = find_table(extra, ["Source row", "Student ID", "Last name", "First name", "Academic standing", "Major", "HW1", "HW2", "Midterm", "HW3", "Final"])
    roster._tbl.append(copy.deepcopy(roster.rows[-1]._tr))
    save_case(extra, args.output_root, "unsupported_extra_student")

    broken_dir = args.output_root / "broken_roster_links"
    broken_dir.mkdir(exist_ok=True)
    strip_internal_links(args.golden, broken_dir / "cs161_roster_control_book.docx")

    medium = Document(args.golden)
    quality = find_table(medium, ["Control", "Result", "Evidence"])
    remove_row(quality.rows[-1])
    save_case(medium, args.output_root, "medium_missing_one_quality_control")

    flattened = Document()
    flattened.add_heading("CS161 Roster Control Book", level=0)
    one_pixel_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
    )
    with tempfile.TemporaryDirectory(prefix="cs161-flat-") as temp:
        image_path = Path(temp) / "flattened.png"
        image_path.write_bytes(one_pixel_png)
        flattened.add_picture(str(image_path), width=Inches(6.5), height=Inches(8.0))
        save_case(flattened, args.output_root, "flattened_screenshot_only")

    corrupt = args.output_root / "corrupt_docx"
    corrupt.mkdir(exist_ok=True)
    (corrupt / "cs161_roster_control_book.docx").write_bytes(b"not a DOCX")
    (args.output_root / "empty_output").mkdir(exist_ok=True)


if __name__ == "__main__":
    main()
