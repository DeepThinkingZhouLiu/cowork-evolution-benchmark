#!/usr/bin/env python3
"""Build task-specific calibration cases for the perfect-Final verification packet."""

from __future__ import annotations

import argparse
import base64
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Inches
from lxml import etree


OUTPUT_NAME = "perfect_final_verification_packet.docx"
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
M_NS = {"m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}


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
            outgoing.writestr(info, mutate(info.filename, payload))


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


def strip_equations(name: str, payload: bytes) -> bytes:
    if name != "word/document.xml":
        return payload
    root = etree.fromstring(payload)
    for equation in root.xpath("//m:oMath", namespaces=M_NS):
        equation.getparent().remove(equation)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--golden", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    args.output_root.mkdir(parents=True, exist_ok=True)

    ledger_header = ["Source row", "Student ID", "Name", "Final", "Final cell", "Decision", "Reason"]
    ledger_error = Document(args.golden)
    ledger = tables_with_header(ledger_error, ledger_header)[0]
    ledger.rows[7].cells[5].text = "Selected"
    ledger.rows[7].cells[6].text = "Final = 100"
    save_case(ledger_error, args.output_root, "wrong_threshold_inclusion_97")

    missing_ledger = Document(args.golden)
    ledger = tables_with_header(missing_ledger, ledger_header)[0]
    remove_row(ledger.rows[3])
    save_case(missing_ledger, args.output_root, "missing_population_ledger_row")

    wrong_subject = Document(args.golden)
    for paragraph in wrong_subject.paragraphs:
        if clean(paragraph.text) == "Verification Certificate • S1007":
            paragraph.text = "Verification Certificate • S1008"
            break
    save_case(wrong_subject, args.output_root, "wrong_certificate_subject")

    wrong_score = Document(args.golden)
    assessments = tables_with_header(wrong_score, ["Assessment", "Score", "Source cell"])
    assessments[0].rows[2].cells[1].text = "95"
    save_case(wrong_score, args.output_root, "wrong_certificate_score")

    wrong_locator = Document(args.golden)
    assessments = tables_with_header(wrong_locator, ["Assessment", "Score", "Source cell"])
    assessments[1].rows[4].cells[2].text = "Sheet1!H8"
    save_case(wrong_locator, args.output_root, "wrong_certificate_locator")

    wrong_calculation = Document(args.golden)
    calculations = tables_with_header(wrong_calculation, ["Check", "Result", "Basis"])
    calculations[0].rows[1].cells[1].text = "61.00"
    save_case(wrong_calculation, args.output_root, "wrong_prefinal_average")

    mutate_zip(args.golden, args.output_root / "missing_native_equations" / OUTPUT_NAME, strip_equations)

    wrong_controls = Document(args.golden)
    controls = tables_with_header(wrong_controls, ["Control", "Result", "Control", "Result"])[0]
    controls.rows[4].cells[3].text = "54.0"
    save_case(wrong_controls, args.output_root, "wrong_final_mean")

    wrong_distribution = Document(args.golden)
    distribution = tables_with_header(wrong_distribution, ["Final band", "Count", "Share"])[0]
    distribution.rows[5].cells[1].text = "3"
    distribution.rows[5].cells[2].text = "33.3%"
    save_case(wrong_distribution, args.output_root, "wrong_distribution_band")

    wrong_boundary = Document(args.golden)
    boundary = tables_with_header(wrong_boundary, ["Source row", "Student ID", "Name", "Final", "Gap", "Final cell", "Decision"])[0]
    boundary.rows[1].cells[1].text = "S1008"
    boundary.rows[1].cells[3].text = "81"
    boundary.rows[1].cells[4].text = "19"
    boundary.rows[1].cells[5].text = "Sheet1!I10"
    save_case(wrong_boundary, args.output_root, "wrong_nearest_boundary")

    mutate_zip(args.golden, args.output_root / "missing_internal_links" / OUTPUT_NAME, strip_internal_links)

    medium = Document(args.golden)
    closure = tables_with_header(medium, ["Control", "Result", "Evidence"])[0]
    remove_row(closure.rows[-1])
    save_case(medium, args.output_root, "medium_missing_closure_control")

    flattened = Document()
    flattened.add_heading("Perfect Final Score Verification Packet", level=0)
    one_pixel_png = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
    with tempfile.TemporaryDirectory(prefix="final100-flat-") as temp_dir:
        image = Path(temp_dir) / "flattened.png"
        image.write_bytes(one_pixel_png)
        flattened.add_picture(str(image), width=Inches(7), height=Inches(8.5))
        save_case(flattened, args.output_root, "flattened_screenshot_only")

    extra = Document(args.golden)
    save_case(extra, args.output_root, "unexpected_extra_docx")
    Document().save(args.output_root / "unexpected_extra_docx" / "extra.docx")

    corrupt = args.output_root / "corrupt_docx"
    corrupt.mkdir(exist_ok=True)
    (corrupt / OUTPUT_NAME).write_bytes(b"not a DOCX")
    (args.output_root / "empty_output").mkdir(exist_ok=True)


if __name__ == "__main__":
    main()
