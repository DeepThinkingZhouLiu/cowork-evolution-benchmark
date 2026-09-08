from __future__ import annotations

import argparse
import shutil
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree
from openpyxl import load_workbook


OUTPUT_NAME = "highest_score_recognition_review.docx"


@dataclass(frozen=True)
class Record:
    source_row: int
    name: str
    midterm1: int
    midterm2: int

    @property
    def delta(self) -> int:
        return self.midterm2 - self.midterm1

    @property
    def direction(self) -> str:
        return "Down" if self.delta < 0 else ("Up" if self.delta > 0 else "No change")


def read_shortlist(path: Path) -> list[tuple[int, str, int]]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    assert [cell.value for cell in sheet[1]] == ["Name", "midterm1"]
    rows = [(row_number, str(name), int(score)) for row_number, (name, score) in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2)]
    workbook.close()
    assert len(rows) == 4
    return rows


def read_register(path: Path) -> list[Record]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    assert [cell.value for cell in sheet[1]] == ["Name", "midterm1", "midterm2"]
    rows = [Record(row_number, str(name), int(first), int(second)) for row_number, (name, first, second) in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2)]
    workbook.close()
    assert len(rows) == 20
    return rows


def shade(cell, color: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    element = properties.find(qn("w:shd")) or OxmlElement("w:shd")
    element.set(qn("w:fill"), color)
    if element.getparent() is None:
        properties.append(element)


def table_borders(table, color: str = "B7C9D6") -> None:
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)
        borders.append(node)
    table._tbl.tblPr.append(borders)


def repeat_header(row) -> None:
    node = OxmlElement("w:tblHeader")
    node.set(qn("w:val"), "true")
    row._tr.get_or_add_trPr().append(node)


def no_split(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def bookmark(paragraph, name: str, identifier: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(identifier))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(identifier))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def link(paragraph, label: str, anchor: str) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.extend([color, underline])
    run.append(properties)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def field(paragraph, instruction: str, display: str) -> None:
    node = OxmlElement("w:fldSimple")
    node.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = display
    run.append(text)
    node.append(run)
    paragraph._p.append(node)


def add_heading(document: Document, text: str, level: int, bookmark_name: str | None = None, bookmark_id: int = 1):
    paragraph = document.add_heading(text, level=level)
    if bookmark_name:
        bookmark(paragraph, bookmark_name, bookmark_id)
    return paragraph


def set_cell(cell, value: str, bold: bool = False, size: float = 8.5) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(value)
    run.bold = bold
    run.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def header_row(table, labels: list[str]) -> None:
    row = table.rows[0]
    repeat_header(row)
    no_split(row)
    for cell, label in zip(row.cells, labels):
        set_cell(cell, label, bold=True, size=8)
        shade(cell, "1F4E78")
        for run in cell.paragraphs[0].runs:
            run.font.color.rgb = RGBColor(255, 255, 255)


def configure_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Sans"
    normal.font.size = Pt(9)
    normal.paragraph_format.space_after = Pt(3)
    document.styles["Title"].font.name = "Liberation Sans"
    document.styles["Title"].font.size = Pt(23)
    document.styles["Title"].font.bold = True
    document.styles["Title"].font.color.rgb = RGBColor(31, 78, 120)
    for name, size, color in [("Heading 1", 15, "1F4E78"), ("Heading 2", 11.5, "2F5597")]:
        style = document.styles[name]
        style.font.name = "Liberation Sans"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.keep_with_next = True


def configure_page(section, landscape: bool = False) -> None:
    section.top_margin = Cm(1.65)
    section.bottom_margin = Cm(1.6)
    section.left_margin = Cm(1.65)
    section.right_margin = Cm(1.65)
    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width


def configure_running_parts(document: Document) -> None:
    for section in document.sections:
        section.header.is_linked_to_previous = False
        paragraph = section.header.paragraphs[0]
        paragraph.text = "MIDTERM RECOGNITION CONTROL REVIEW  |  HOLD"
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.runs[0].font.size = Pt(8)
        paragraph.runs[0].font.bold = True
        paragraph.runs[0].font.color.rgb = RGBColor(192, 0, 0)
        section.footer.is_linked_to_previous = False
        page = section.footer.paragraphs[0]
        page.alignment = WD_ALIGN_PARAGRAPH.CENTER
        page.add_run("Controlled working paper  •  ")
        field(page, " PAGE ", "1")
        page.add_run(" / ")
        field(page, " NUMPAGES ", "4")
        for run in page.runs:
            run.font.size = Pt(8)


def inject_comments(path: Path, comments: list[tuple[str, str]]) -> None:
    wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns = {"w": wns}
    temporary = path.with_suffix(".comments.docx")
    with zipfile.ZipFile(path) as source:
        parts = {info.filename: source.read(info.filename) for info in source.infolist()}
        infos = {info.filename: info for info in source.infolist()}
    root = etree.fromstring(parts["word/document.xml"])
    comments_root = etree.Element(qn("w:comments"), nsmap={"w": wns})
    for comment_id, (anchor, note) in enumerate(comments):
        paragraphs = root.xpath("//w:p[contains(string(.), $anchor)]", namespaces=ns, anchor=anchor)
        assert len(paragraphs) == 1, (anchor, len(paragraphs))
        paragraph = paragraphs[0]
        runs = paragraph.xpath("./w:r[w:t]", namespaces=ns)
        assert runs
        start = OxmlElement("w:commentRangeStart")
        start.set(qn("w:id"), str(comment_id))
        end = OxmlElement("w:commentRangeEnd")
        end.set(qn("w:id"), str(comment_id))
        reference_run = OxmlElement("w:r")
        reference = OxmlElement("w:commentReference")
        reference.set(qn("w:id"), str(comment_id))
        reference_run.append(reference)
        first_index = paragraph.index(runs[0])
        last_index = paragraph.index(runs[-1])
        paragraph.insert(first_index, start)
        paragraph.insert(last_index + 2, end)
        paragraph.insert(last_index + 3, reference_run)
        comment = etree.SubElement(comments_root, qn("w:comment"))
        comment.set(qn("w:id"), str(comment_id))
        comment.set(qn("w:author"), "Academic Records Review")
        cp = etree.SubElement(comment, qn("w:p"))
        cr = etree.SubElement(cp, qn("w:r"))
        etree.SubElement(cr, qn("w:t")).text = note
    parts["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    parts["word/comments.xml"] = etree.tostring(comments_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    content_types = etree.fromstring(parts["[Content_Types].xml"])
    override = etree.SubElement(content_types, f"{{{content_ns}}}Override")
    override.set("PartName", "/word/comments.xml")
    override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    parts["[Content_Types].xml"] = etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone="yes")
    relationships = etree.fromstring(parts["word/_rels/document.xml.rels"])
    used = {relationship.get("Id") for relationship in relationships}
    number = 1
    while f"rId{number}" in used:
        number += 1
    relationship = etree.SubElement(relationships, f"{{{rel_ns}}}Relationship")
    relationship.set("Id", f"rId{number}")
    relationship.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
    relationship.set("Target", "comments.xml")
    parts["word/_rels/document.xml.rels"] = etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone="yes")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for name, payload in parts.items():
            info = zipfile.ZipInfo(name, (2026, 9, 8, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = infos[name].external_attr if name in infos else 0o600 << 16
            target.writestr(info, payload)
    shutil.move(temporary, path)


def build(input_dir: Path, output_path: Path) -> None:
    shortlist = read_shortlist(input_dir / "nomination_shortlist.xlsx")
    records = read_register(input_dir / "detailed_midterm_register.xlsx")
    shortlist_pairs = {(name, score) for _, name, score in shortlist}
    maximum = max(record.midterm1 for record in records)
    co_highest = [record for record in records if record.midterm1 == maximum]
    minimum = min(record.midterm1 for record in records)
    distinct_names = len({record.name for record in records})
    counts = Counter(record.name for record in records)
    duplicates = {name: count for name, count in counts.items() if count > 1}
    omitted = [record for record in co_highest if (record.name, record.midterm1) not in shortlist_pairs]
    assert [(record.name, record.midterm1) for record in co_highest] == [("Tina", 100), ("Rachel", 100)]
    assert [record.name for record in omitted] == ["Tina"]

    document = Document()
    configure_styles(document)
    configure_page(document.sections[0])
    properties = document.core_properties
    properties.title = "Midterm Recognition Release-Readiness Review"
    properties.author = "Academic Records Review"
    properties.subject = "Evidence reconciliation and controlled recognition drafts"
    properties.keywords = "midterm; recognition; reconciliation; hold; audit"

    title = document.add_paragraph(style="Title")
    title.add_run("Midterm Recognition\nRelease-Readiness Review")
    subtitle = document.add_paragraph("Evidence reconciliation working paper  |  Prepared from the two frozen Excel inputs")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in subtitle.runs:
        run.italic = True
        run.font.color.rgb = RGBColor(89, 89, 89)

    status = document.add_table(rows=1, cols=2)
    status.autofit = False
    status.columns[0].width = Cm(4)
    status.columns[1].width = Cm(12.5)
    set_cell(status.cell(0, 0), "RELEASE STATUS", bold=True, size=9)
    set_cell(status.cell(0, 1), "HOLD — co-highest record omitted from nomination shortlist", bold=True, size=10)
    shade(status.cell(0, 0), "7F0000")
    shade(status.cell(0, 1), "F4CCCC")
    status.cell(0, 0).paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    table_borders(status, "7F0000")

    add_heading(document, "Navigation", 1, "navigation", 1)
    nav = document.add_paragraph()
    link(nav, "Executive finding", "executive_finding")
    nav.add_run("  •  ")
    link(nav, "Shortlist reconciliation", "shortlist_reconciliation")
    nav.add_run("  •  ")
    link(nav, "Controlled drafts", "controlled_drafts")
    nav.add_run("  •  ")
    link(nav, "20-record evidence ledger", "record_ledger")
    toc = document.add_paragraph()
    toc.add_run("Word table of contents field: ").bold = True
    field(toc, ' TOC \\o "1-2" \\h \\z \\u ', "Right-click and update field")

    add_heading(document, "Executive finding", 1, "executive_finding", 2)
    document.add_paragraph(
        "The nomination shortlist names Rachel at 100 as its highest entry. The detailed register contains two separate Midterm 1 records at 100: Tina (source row 16) and Rachel (source row 18). Tina is absent from the shortlist. Release is therefore on hold; neither record may be described as the sole highest score."
    )
    facts = document.add_table(rows=3, cols=4)
    header_row(facts, ["Control metric", "Result", "Control metric", "Result"])
    metric_rows = [
        ("Detailed records", str(len(records)), "Distinct displayed names", str(distinct_names)),
        ("Midterm 1 max / min / spread", f"{maximum} / {minimum} / {maximum - minimum}", "Midterm 2 max / min", f"{max(r.midterm2 for r in records)} / {min(r.midterm2 for r in records)}"),
    ]
    for table_row, values in zip(facts.rows[1:], metric_rows):
        no_split(table_row)
        for cell, value in zip(table_row.cells, values):
            set_cell(cell, value, bold=value in {str(len(records)), str(distinct_names)})
    table_borders(facts)
    document.add_paragraph(
        "Identity boundary: every spreadsheet row is treated as a separate assessment record. Repeated display names cannot be merged or assumed to identify one student because neither input provides a student ID."
    )

    add_heading(document, "Nomination shortlist reconciliation", 1, "shortlist_reconciliation", 3)
    reconciliation = document.add_table(rows=1, cols=6)
    header_row(reconciliation, ["Shortlist row", "Name", "Listed M1", "Detailed exact match", "Detailed source row", "Control result"])
    for source_row, name, score in shortlist:
        matches = [record for record in records if record.name == name and record.midterm1 == score]
        row = reconciliation.add_row()
        values = [str(source_row), name, str(score), "Yes" if matches else "No", ", ".join(str(record.source_row) for record in matches) or "—", "Verified exact pair" if matches else "Unresolved"]
        for cell, value in zip(row.cells, values):
            set_cell(cell, value)
        no_split(row)
    table_borders(reconciliation)
    gap = document.add_paragraph()
    gap.add_run("Blocking gap: ").bold = True
    gap.add_run("Tina | Midterm 1 100 | detailed register row 16 | no exact entry in nomination_shortlist.xlsx.")

    add_heading(document, "Controlled recognition drafts", 1, "controlled_drafts", 4)
    document.add_paragraph("These are editable working drafts only. Do not release until the nomination population and identity controls are resolved.")
    for index, record in enumerate(co_highest, start=1):
        heading = document.add_heading(f"Draft {index}: {record.name}", level=2)
        bookmark(heading, f"draft_{record.name.lower()}", 10 + index)
        anchor = document.add_paragraph()
        anchor.add_run(f"{record.name} — DRAFT — NOT FOR RELEASE").bold = True
        paragraph = document.add_paragraph(
            f"Dear {record.name}, your Midterm 1 assessment record is 100, tied for the highest Midterm 1 score in the detailed register. This draft does not assert a sole winner and is held pending nomination-list reconciliation."
        )
        paragraph.paragraph_format.left_indent = Cm(0.6)
        trail = document.add_paragraph()
        link(trail, f"Open detailed evidence row {record.source_row}", f"record_{record.source_row}")

    landscape = document.add_section(WD_SECTION.NEW_PAGE)
    configure_page(landscape, landscape=True)
    add_heading(document, "Detailed assessment-record ledger", 1, "record_ledger", 20)
    document.add_paragraph(
        "Source: detailed_midterm_register.xlsx, Sheet1. Delta = Midterm 2 − Midterm 1. Direction is Down for a negative delta, Up for a positive delta, and No change for zero. Shortlist match requires an exact Name + Midterm 1 pair."
    )
    ledger = document.add_table(rows=1, cols=8)
    header_row(ledger, ["Source row", "Displayed name", "M1", "M2", "Delta", "Direction", "Exact shortlist pair", "Evidence / control note"])
    for record in records:
        row = ledger.add_row()
        exact = (record.name, record.midterm1) in shortlist_pairs
        note = "Co-highest; OMITTED from shortlist" if record in omitted else ("Co-highest; listed" if record in co_highest else ("Exact pair listed" if exact else "Not listed"))
        values = [str(record.source_row), record.name, str(record.midterm1), str(record.midterm2), f"{record.delta:+d}", record.direction, "Yes" if exact else "No", note]
        for cell, value in zip(row.cells, values):
            set_cell(cell, value, bold=record in co_highest, size=7.5)
        shade(row.cells[7], "F4CCCC" if record in omitted else ("FFF2CC" if record in co_highest else "FFFFFF"))
        bookmark(row.cells[0].paragraphs[0], f"record_{record.source_row}", 100 + record.source_row)
        no_split(row)
    table_borders(ledger)

    add_heading(document, "Duplicate-name and evidence-boundary controls", 1, "identity_controls", 30)
    duplicate_table = document.add_table(rows=1, cols=3)
    header_row(duplicate_table, ["Displayed name", "Record count", "Required treatment"])
    for name, count in duplicates.items():
        row = duplicate_table.add_row()
        for cell, value in zip(row.cells, [name, str(count), "Keep as separate assessment records; no student ID is available"]):
            set_cell(cell, value)
        no_split(row)
    table_borders(duplicate_table)
    add_heading(document, "Release controls and source notes", 1, "source_notes", 40)
    bullets = [
        "Do not publish either recognition statement while the shortlist omits a co-highest record.",
        "Do not infer student identity, course, class rank beyond Midterm 1, grade, GPA, instructor, term, or award eligibility from these inputs.",
        "Resolve Tina's missing nomination and obtain a stable student identifier before any recipient-level release.",
        "Frozen sources used: nomination_shortlist.xlsx and detailed_midterm_register.xlsx, both Sheet1; all findings above can be recomputed from those files.",
    ]
    for text in bullets:
        document.add_paragraph(text, style="List Bullet")

    configure_running_parts(document)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    inject_comments(
        output_path,
        [
            ("Tina — DRAFT — NOT FOR RELEASE", "Evidence: detailed_midterm_register.xlsx, Sheet1 row 16, Midterm 1=100. No Tina+100 row exists in nomination_shortlist.xlsx."),
            ("Rachel — DRAFT — NOT FOR RELEASE", "Evidence: detailed_midterm_register.xlsx, Sheet1 row 18, Midterm 1=100; nomination_shortlist.xlsx, Sheet1 row 2, Rachel+100."),
        ],
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    arguments = parser.parse_args()
    build(arguments.input_dir, arguments.output_dir / OUTPUT_NAME)


if __name__ == "__main__":
    main()
