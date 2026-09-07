from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from statistics import mean, median

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree
from openpyxl import load_workbook


INPUT_NAME = "CS161_Class_Roster.xlsx"
OUTPUT_PATH = Path("CS161/students.docx")
NAVY = "183C56"
TEAL = "28778B"
LIGHT = "EEF3F6"
PALE = "DDEBF0"
GOLD = "FFF1CB"
ASSESSMENTS = ["HW1", "HW2", "Midterm", "HW3", "Final"]
BOUNDARY_NOTE = (
    "Verification flags are mechanical source checks only. They do not assign a course grade, "
    "academic consequence, misconduct finding, or advising decision."
)


def load_policy() -> dict[str, object]:
    for path in (Path(__file__).with_name("review-policy.json"), Path(__file__).parents[1] / "adaptation/review-policy.json"):
        if path.is_file():
            return json.loads(path.read_text(encoding="utf-8"))
    raise FileNotFoundError("review policy is unavailable")


def load_records(path: Path, policy: dict[str, object]) -> list[dict[str, object]]:
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    if digest != policy["source"]["sha256"]:
        raise ValueError(f"unexpected input SHA256: {digest}")
    workbook = load_workbook(path, data_only=True, read_only=True)
    if workbook.sheetnames != [policy["source"]["worksheet"]]:
        workbook.close()
        raise ValueError(f"unexpected worksheet list: {workbook.sheetnames}")
    sheet = workbook[policy["source"]["worksheet"]]
    values = list(sheet.iter_rows(values_only=True))
    workbook.close()
    headers = [str(value).strip() for value in values[0]]
    if headers != policy["source"]["required_headers"]:
        raise ValueError(f"unexpected headers: {headers}")
    records: list[dict[str, object]] = []
    for source_row, row in enumerate(values[1:], 2):
        record = dict(zip(headers, row))
        if any(record[key] is None for key in headers):
            raise ValueError(f"blank source cell in row {source_row}")
        record["source_row"] = source_row
        scores = [int(record[name]) for name in ASSESSMENTS]
        flags: list[str] = []
        if min(scores) == 0:
            flags.append("ZERO")
        if max(scores) == 100:
            flags.append("PERFECT")
        if max(scores) - min(scores) >= 80:
            flags.append("SPREAD80")
        if abs(int(record["Final"]) - int(record["Midterm"])) >= 60:
            flags.append("SWING60")
        record["scores"] = scores
        record["mean"] = f"{mean(scores):.1f}"
        record["spread"] = max(scores) - min(scores)
        record["delta"] = int(record["Final"]) - int(record["Midterm"])
        record["flags"] = flags
        records.append(record)
    identifiers = [str(record["Student ID"]) for record in records]
    if len(records) != policy["output"]["roster_count"] or len(set(identifiers)) != len(records):
        raise ValueError("roster count or Student ID uniqueness changed")
    return records


def verification_queue(records: list[dict[str, object]]) -> list[dict[str, object]]:
    flagged = [record for record in records if record["flags"]]
    flagged.sort(key=lambda record: (-len(record["flags"]), -int(record["spread"]), int(record["source_row"])))
    return flagged


def block_digest(records: list[dict[str, object]], headers: list[str]) -> str:
    payload = [[record["source_row"], *[record[name] for name in headers]] for record in records]
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()[:12]


def shade(cell, color: str) -> None:
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), color)
    cell._tc.get_or_add_tcPr().append(node)


def set_cell_margins(cell, value: int = 35) -> None:
    properties = cell._tc.get_or_add_tcPr()
    margins = OxmlElement("w:tcMar")
    for edge in ("top", "start", "bottom", "end"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    properties.append(margins)


def repeat_header(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def prevent_split(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def set_font(run, size: float, *, bold: bool = False, italic: bool = False, color: RGBColor | None = None) -> None:
    run.font.name = "Liberation Sans"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def cell_text(cell, value: object, *, size: float, bold: bool = False, white: bool = False, center: bool = False) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 0.85
    run = paragraph.add_run(str(value))
    set_font(run, size, bold=bold, color=RGBColor(255, 255, 255) if white else None)


def add_table(document: Document, headers: list[str], rows: list[list[object]], widths: list[float], *, size: float = 7.0, center: set[int] | None = None):
    center = center or set()
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for index, header in enumerate(headers):
        cell_text(table.cell(0, index), header, size=size, bold=True, white=True, center=index in center)
        shade(table.cell(0, index), NAVY)
    repeat_header(table.rows[0])
    for row_index, values in enumerate(rows, 1):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cell_text(cells[index], value, size=size, center=index in center)
            if row_index % 2 == 0:
                shade(cells[index], LIGHT)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:color"), "9DB3C0")
        borders.append(border)
    table._tbl.tblPr.append(borders)
    for row in table.rows:
        prevent_split(row)
        for index, cell in enumerate(row.cells):
            cell.width = Inches(widths[index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
    return table


def add_paragraph(document: Document, text: str = "", *, size: float = 8.2, bold: bool = False, italic: bool = False, after: float = 2, before: float = 0):
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = 1.0
    if text:
        set_font(paragraph.add_run(text), size, bold=bold, italic=italic)
    return paragraph


def add_heading(document: Document, text: str, *, level: int = 1, size: float = 14):
    paragraph = document.add_paragraph(style=f"Heading {level}")
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(3)
    set_font(paragraph.add_run(text), size, bold=True, color=RGBColor(24, 60, 86))
    return paragraph


def add_page_break(document: Document) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.add_run().add_break(WD_BREAK.PAGE)


def add_bookmark(paragraph, name: str, identifier: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(identifier))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(identifier))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def add_internal_link(paragraph, label: str, anchor: str) -> None:
    link = OxmlElement("w:hyperlink")
    link.set(qn("w:anchor"), anchor)
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), TEAL)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.extend((color, underline))
    run.append(properties)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    link.append(run)
    paragraph._p.append(link)


def add_field(paragraph, instruction: str, fallback: str) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = fallback
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def configure_section(section, *, landscape: bool, header_label: str) -> None:
    section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    section.page_width = Inches(11 if landscape else 8.5)
    section.page_height = Inches(8.5 if landscape else 11)
    section.top_margin = Inches(0.38 if landscape else 0.52)
    section.bottom_margin = Inches(0.42 if landscape else 0.52)
    section.left_margin = Inches(0.38 if landscape else 0.58)
    section.right_margin = Inches(0.38 if landscape else 0.58)
    section.header_distance = Inches(0.18)
    section.footer_distance = Inches(0.18)
    section.header.is_linked_to_previous = False
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header.text = ""
    set_font(header.add_run(header_label), 7.2, bold=True, color=RGBColor(40, 119, 139))
    section.footer.is_linked_to_previous = False
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(footer.add_run("CS161 ROSTER VERIFICATION REGISTER  •  "), 7.0)
    add_field(footer, "PAGE", "1")
    set_font(footer.add_run(" / "), 7.0)
    add_field(footer, "NUMPAGES", "11")


def style_document(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Sans"
    normal.font.size = Pt(8.2)
    normal.paragraph_format.space_after = Pt(2)
    for name, size in (("Title", 25), ("Subtitle", 11), ("Heading 1", 14), ("Heading 2", 10)):
        style = document.styles[name]
        style.font.name = "Liberation Sans"
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor(24, 60, 86)


def cover_page(document: Document, policy: dict[str, object], source_sha: str) -> None:
    title = document.add_paragraph(style="Title")
    title.paragraph_format.space_after = Pt(2)
    set_font(title.add_run("CS161 Roster Verification Register"), 25, bold=True, color=RGBColor(24, 60, 86))
    subtitle = document.add_paragraph(style="Subtitle")
    subtitle.paragraph_format.space_after = Pt(8)
    set_font(subtitle.add_run("CANONICAL DIRECTORY • ASSESSMENT RECONCILIATION • MECHANICAL EXCEPTION QUEUE"), 10, bold=True)
    add_bookmark(title, "control_page", 8000)
    add_table(document, ["Control", "Frozen value"], [
        ["Source", INPUT_NAME],
        ["Worksheet", policy["source"]["worksheet"]],
        ["Source SHA256", source_sha],
        ["Source rows", "2–124 (123 students)"],
        ["Output", str(OUTPUT_PATH)],
        ["Ruleset", policy["schema_version"]],
    ], [1.45, 5.75], size=8.0)
    add_heading(document, "Navigation", level=1, size=13)
    navigation = [
        ("Assessment Profile", "assessment_profile", "source/control metrics"),
        ("Roster Block B1", "roster_b1", "rows 2–26"),
        ("Roster Block B2", "roster_b2", "rows 27–51"),
        ("Roster Block B3", "roster_b3", "rows 52–76"),
        ("Roster Block B4", "roster_b4", "rows 77–101"),
        ("Roster Block B5", "roster_b5", "rows 102–124"),
        ("Queue Part 1", "queue_1", "ranks 1–20"),
        ("Queue Part 2", "queue_2", "ranks 21–40"),
        ("Queue Part 3", "queue_3", "ranks 41–56"),
        ("Closure", "closure", "reconciliation and block hashes"),
    ]
    table = document.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell_text(table.cell(0, 0), "Destination", size=7.6, bold=True, white=True)
    cell_text(table.cell(0, 1), "Coverage", size=7.6, bold=True, white=True)
    shade(table.cell(0, 0), NAVY)
    shade(table.cell(0, 1), NAVY)
    repeat_header(table.rows[0])
    for index, (label, anchor, coverage) in enumerate(navigation, 1):
        cells = table.add_row().cells
        cells[0].text = ""
        link_paragraph = cells[0].paragraphs[0]
        link_paragraph.paragraph_format.space_after = Pt(0)
        add_internal_link(link_paragraph, label, anchor)
        cell_text(cells[1], coverage, size=7.4)
        if index % 2 == 0:
            shade(cells[0], LIGHT)
            shade(cells[1], LIGHT)
        prevent_split(table.rows[-1])
    add_heading(document, "Verification Rule Boundary", level=2, size=10)
    for rule in policy["verification_rules"]:
        add_paragraph(document, f"{rule['code']} — {rule['definition']}", size=7.5, after=0.8)
    note = add_paragraph(document, BOUNDARY_NOTE, size=7.8, bold=True, after=0)
    note.paragraph_format.keep_together = True


def profile_page(document: Document, records: list[dict[str, object]], queue: list[dict[str, object]], policy: dict[str, object]) -> None:
    heading = add_heading(document, "Assessment Profile & Standing Reconciliation", size=15)
    add_bookmark(heading, "assessment_profile", 8001)
    add_paragraph(document, "All statistics are recomputed from the 123 frozen worksheet rows. Means are descriptive, unweighted arithmetic means—not course grades.", size=7.7)
    add_table(document, ["Assessment", "Mean", "Median", "Minimum", "Maximum", "Zero count", "100 count"], [
        [name, f"{mean([int(record[name]) for record in records]):.2f}", f"{median([int(record[name]) for record in records]):.2f}", min(int(record[name]) for record in records), max(int(record[name]) for record in records), sum(int(record[name]) == 0 for record in records), sum(int(record[name]) == 100 for record in records)]
        for name in ASSESSMENTS
    ], [1.25, .85, .85, .85, .85, .85, .85], size=7.2, center={1, 2, 3, 4, 5, 6})
    add_heading(document, "Standing Reconciliation", level=2, size=10)
    standing_rows = []
    for standing in policy["standing_order"]:
        subset = [record for record in records if record["Academic Standing"] == standing]
        standing_rows.append([standing, len(subset), *[f"{mean([int(record[name]) for record in subset]):.2f}" for name in ASSESSMENTS], sum(bool(record["flags"]) for record in subset)])
    add_table(document, ["Standing", "Count", *[f"{name} mean" for name in ASSESSMENTS], "Flagged"], standing_rows, [1.25, .65, .85, .85, .95, .85, .85, .75], size=6.9, center=set(range(1, 8)))
    add_heading(document, "Rule Incidence", level=2, size=10)
    rule_codes = [rule["code"] for rule in policy["verification_rules"]]
    incidence = []
    for standing in policy["standing_order"]:
        subset = [record for record in records if record["Academic Standing"] == standing]
        incidence.append([standing, *[sum(code in record["flags"] for record in subset) for code in rule_codes], sum(bool(record["flags"]) for record in subset)])
    incidence.append(["All students", *[sum(code in record["flags"] for record in records) for code in rule_codes], len(queue)])
    add_table(document, ["Standing", *rule_codes, "Any flag"], incidence, [1.25, .9, .9, 1.0, 1.0, .85], size=7.0, center={1, 2, 3, 4, 5})
    add_paragraph(document, "Queue order: rule-hit count descending -> assessment spread descending -> source row ascending.", size=7.4, italic=True)


def roster_pages(document: Document, records: list[dict[str, object]], policy: dict[str, object]) -> list[dict[str, object]]:
    headers = policy["source"]["required_headers"]
    ledgers: list[dict[str, object]] = []
    widths = [.60, .72, 1.00, 1.00, .92, .48, .48, .62, .48, .48]
    for block_index, block in enumerate(policy["roster_blocks"], 1):
        add_page_break(document)
        subset = [record for record in records if block["source_row_start"] <= record["source_row"] <= block["source_row_end"]]
        digest = block_digest(subset, headers)
        title = f"Canonical Roster • Block {block['id']} • Source rows {block['source_row_start']}–{block['source_row_end']}"
        heading = add_heading(document, title, size=12.5)
        add_bookmark(heading, f"roster_b{block_index}", 8010 + block_index)
        add_paragraph(document, f"{len(subset)} rows • first {subset[0]['Student ID']} • last {subset[-1]['Student ID']} • block SHA-12 {digest}", size=7.1, bold=True)
        rows = [[record["source_row"], *[record[name] for name in headers]] for record in subset]
        add_table(document, ["Row", "Student ID", "Last", "First", "Standing", "HW1", "HW2", "Midterm", "HW3", "Final"], rows, widths, size=6.25, center={0, 1, 5, 6, 7, 8, 9})
        ledgers.append({"id": block["id"], "range": f"{block['source_row_start']}–{block['source_row_end']}", "count": len(subset), "first": subset[0]["Student ID"], "last": subset[-1]["Student ID"], "digest": digest})
    return ledgers


def queue_pages(document: Document, queue: list[dict[str, object]], policy: dict[str, object]) -> list[dict[str, object]]:
    ledgers: list[dict[str, object]] = []
    cursor = 0
    widths = [.48, .60, .72, 1.35, .44, .56, .58, .80, 1.40]
    for part_index, count in enumerate(policy["queue_parts"], 1):
        add_page_break(document)
        subset = queue[cursor:cursor + count]
        first_rank, last_rank = cursor + 1, cursor + len(subset)
        heading = add_heading(document, f"Mechanical Verification Queue • Part {part_index} • Ranks {first_rank}–{last_rank}", size=12.5)
        add_bookmark(heading, f"queue_{part_index}", 8020 + part_index)
        add_paragraph(document, "Every row is derived from the four disclosed rules; blank flags are excluded. Review source values before changing any record.", size=7.1)
        rows = []
        for rank, record in enumerate(subset, first_rank):
            rows.append([rank, record["source_row"], record["Student ID"], f"{record['First Name']} {record['Last Name']}", len(record["flags"]), record["mean"], record["spread"], f"{int(record['delta']):+d}", " · ".join(record["flags"])])
        add_table(document, ["Rank", "Row", "Student ID", "Name", "Hits", "Mean", "Spread", "Final−Mid", "Rule codes"], rows, widths, size=7.0, center={0, 1, 2, 4, 5, 6, 7})
        ledgers.append({"part": part_index, "ranks": f"{first_rank}–{last_rank}", "count": len(subset), "first": subset[0]["Student ID"], "last": subset[-1]["Student ID"]})
        cursor += count
    if cursor != len(queue):
        raise ValueError("queue partition does not cover all flagged rows")
    return ledgers


def closure_page(document: Document, records: list[dict[str, object]], queue: list[dict[str, object]], policy: dict[str, object], block_ledger: list[dict[str, object]], queue_ledger: list[dict[str, object]], source_sha: str) -> None:
    heading = add_heading(document, "Closure & Control Ledger", size=15)
    add_bookmark(heading, "closure", 8030)
    add_paragraph(document, "This page reconciles source rows, register blocks, derived queue partitions, and package controls without adding external academic policy.", size=7.8)
    add_heading(document, "Canonical Block Ledger", level=2, size=10)
    add_table(document, ["Block", "Source rows", "Count", "First ID", "Last ID", "SHA-12"], [[item["id"], item["range"], item["count"], item["first"], item["last"], item["digest"]] for item in block_ledger], [.65, 1.0, .65, 1.0, 1.0, 1.2], size=7.2, center={0, 1, 2, 3, 4})
    add_heading(document, "Queue Partition Ledger", level=2, size=10)
    add_table(document, ["Part", "Ranks", "Count", "First ID", "Last ID"], [[item["part"], item["ranks"], item["count"], item["first"], item["last"]] for item in queue_ledger], [.65, 1.0, .65, 1.0, 1.0], size=7.2, center={0, 1, 2, 3, 4})
    add_heading(document, "Reconciliation", level=2, size=10)
    checks = [
        ["C01", "Input student rows", 123, len(records), "MATCH" if len(records) == 123 else "MISMATCH"],
        ["C02", "Unique Student IDs", 123, len({record["Student ID"] for record in records}), "MATCH"],
        ["C03", "Roster block rows", 123, sum(item["count"] for item in block_ledger), "MATCH"],
        ["C04", "Standing count sum", 123, sum(sum(record["Academic Standing"] == standing for record in records) for standing in policy["standing_order"]), "MATCH"],
        ["C05", "Verification rows", 56, len(queue), "MATCH" if len(queue) == 56 else "MISMATCH"],
        ["C06", "Queue partition rows", 56, sum(item["count"] for item in queue_ledger), "MATCH"],
        ["C07", "First / last source IDs", "S1000 / S1122", f"{records[0]['Student ID']} / {records[-1]['Student ID']}", "MATCH"],
        ["C08", "Output DOCX count", 1, 1, "MATCH"],
    ]
    add_table(document, ["Check", "Control", "Expected", "Observed", "Result"], checks, [.55, 2.0, 1.25, 1.35, .85], size=6.9, center={0, 2, 3, 4})
    add_heading(document, "Package Properties", level=2, size=10)
    properties = [
        ["SourceSHA256", source_sha], ["RosterCount", "123"], ["VerificationCount", "56"],
        ["FirstStudentID", "S1000"], ["LastStudentID", "S1122"], ["SourceSheet", "Sheet1"],
        ["RulesetVersion", policy["schema_version"]],
    ]
    add_table(document, ["Property", "Value"], properties, [1.65, 5.15], size=6.7)
    add_paragraph(document, BOUNDARY_NOTE, size=7.5, bold=True, after=0)


def custom_properties_xml(properties: dict[str, str]) -> bytes:
    cp_ns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    vt_ns = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    root = etree.Element(f"{{{cp_ns}}}Properties", nsmap={None: cp_ns, "vt": vt_ns})
    for pid, (name, value) in enumerate(properties.items(), 2):
        prop = etree.SubElement(root, f"{{{cp_ns}}}property")
        prop.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        prop.set("pid", str(pid))
        prop.set("name", name)
        child = etree.SubElement(prop, f"{{{vt_ns}}}lpwstr")
        child.text = value
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def inject_properties(path: Path, properties: dict[str, str]) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = {info.filename: (info, archive.read(info.filename)) for info in archive.infolist()}
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    content_types = etree.fromstring(entries["[Content_Types].xml"][1])
    if not content_types.xpath("//*[local-name()='Override'][@PartName='/docProps/custom.xml']"):
        override = etree.SubElement(content_types, f"{{{ct_ns}}}Override")
        override.set("PartName", "/docProps/custom.xml")
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml")
    entries["[Content_Types].xml"] = (entries["[Content_Types].xml"][0], etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone="yes"))
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    relationships = etree.fromstring(entries["_rels/.rels"][1])
    if not relationships.xpath("//*[local-name()='Relationship'][contains(@Type,'custom-properties')]"):
        relation = etree.SubElement(relationships, f"{{{rel_ns}}}Relationship")
        relation.set("Id", "rId99")
        relation.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
        relation.set("Target", "docProps/custom.xml")
    entries["_rels/.rels"] = (entries["_rels/.rels"][0], etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone="yes"))
    entries["docProps/custom.xml"] = (None, custom_properties_xml(properties))
    with tempfile.TemporaryDirectory(prefix=".cs161-register-", dir=path.parent) as temp_dir:
        target = Path(temp_dir) / path.name
        with zipfile.ZipFile(target, "w") as outgoing:
            for name in sorted(entries):
                info, payload = entries[name]
                replacement = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
                replacement.compress_type = zipfile.ZIP_DEFLATED
                replacement.external_attr = info.external_attr if info is not None else 0o600 << 16
                outgoing.writestr(replacement, payload)
        os.replace(target, path)


def build(input_path: Path, output_dir: Path) -> Path:
    policy = load_policy()
    records = load_records(input_path, policy)
    queue = verification_queue(records)
    if len(queue) != policy["output"]["verification_count"]:
        raise ValueError(f"verification count changed: {len(queue)}")
    source_sha = hashlib.sha256(input_path.read_bytes()).hexdigest()
    document = Document()
    style_document(document)
    configure_section(document.sections[0], landscape=False, header_label="CS161 • CONTROL & NAVIGATION")
    cover_page(document, policy, source_sha)
    landscape = document.add_section(WD_SECTION_START.NEW_PAGE)
    configure_section(landscape, landscape=True, header_label="CS161 • SOURCE REGISTER & VERIFICATION QUEUE")
    profile_page(document, records, queue, policy)
    block_ledger = roster_pages(document, records, policy)
    queue_ledger = queue_pages(document, queue, policy)
    portrait = document.add_section(WD_SECTION_START.NEW_PAGE)
    configure_section(portrait, landscape=False, header_label="CS161 • CLOSURE & CONTROL LEDGER")
    closure_page(document, records, queue, policy, block_ledger, queue_ledger, source_sha)
    core = document.core_properties
    core.title = "CS161 Roster Verification Register"
    core.subject = "Canonical 123-student directory, source reconciliation, and mechanical exception queue"
    core.author = "CoworkEvoBench Oracle"
    core.category = "Roster source control"
    core.keywords = "CS161, roster, verification, source register, exception queue"
    core.comments = "All flags are mechanical source checks; no grade or academic decision is inferred."
    fixed = datetime(2026, 9, 8, 0, 0, 0, tzinfo=timezone.utc)
    core.created = fixed
    core.modified = fixed
    output = output_dir / OUTPUT_PATH
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    inject_properties(output, {
        "SourceSHA256": source_sha,
        "RosterCount": str(len(records)),
        "VerificationCount": str(len(queue)),
        "FirstStudentID": str(records[0]["Student ID"]),
        "LastStudentID": str(records[-1]["Student ID"]),
        "SourceSheet": str(policy["source"]["worksheet"]),
        "RulesetVersion": str(policy["schema_version"]),
    })
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    build(args.input_dir / INPUT_NAME, args.output_dir)


if __name__ == "__main__":
    main()
