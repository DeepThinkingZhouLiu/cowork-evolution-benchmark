from __future__ import annotations

import argparse
import hashlib
import os
import re
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree


INPUT_NAME = "homework_withname.docx"
INPUT_SHA256 = "7df2fd4bf45616e4e78f6adf4dfb5c10ac18886bc777bd43064740be12b8b03b"
NORMALIZED_BODY_SHA256 = "9215262c621e091ae431a8a2b35e56b5d75a92f8183693ce5f5c0a13ccaabf05"
ROLES = [
    "assignment title",
    "academic term",
    "student identity",
    "answering instruction",
    "grading policy",
    "submission instruction",
    "question 1",
    "question 1 option",
    "question 1 option",
    "question 1 option",
    "support instruction",
    "example answer",
    "quotation example",
]
NAVY = "173B57"
BLUE = "286F8E"
LIGHT = "F2F6F8"
GREEN = "E3F1E8"
GOLD = "FFF1CC"


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value if value is not None else "")).strip()


def paragraph_numbering(paragraph) -> str:
    properties = paragraph._p.pPr
    numbering = properties.numPr if properties is not None else None
    if numbering is None or numbering.numId is None:
        return "none"
    level = numbering.ilvl.val if numbering.ilvl is not None else 0
    return f"numId {numbering.numId.val} • level {level}"


def extract_truth(document: Document) -> dict[str, object]:
    texts = [clean(paragraph.text) for paragraph in document.paragraphs]
    if len(texts) != 13:
        raise ValueError(f"Expected 13 source paragraphs, found {len(texts)}")
    joined = "\n".join(texts)
    identifiers = re.findall(r"\b[A-Z]\d{8}\b", joined)
    if len(identifiers) != 1:
        raise ValueError(f"Expected exactly one student ID, found {identifiers}")
    title_match = re.fullmatch(r"(CSE \d+): (Homework \d+)", texts[0])
    term_match = re.fullmatch(r"(Winter|Spring|Summer|Fall) (\d{4})", texts[1])
    if not title_match or not term_match:
        raise ValueError("Could not parse assignment title or term")
    return {
        "student_id": identifiers[0],
        "course": title_match.group(1),
        "assignment": title_match.group(2),
        "term": term_match.group(0),
        "texts": texts,
        "body_sha256": hashlib.sha256(joined.encode("utf-8")).hexdigest(),
        "runs": sum(len(paragraph.runs) for paragraph in document.paragraphs),
        "characters": sum(len(text) for text in texts),
    }


def shade(cell, color: str) -> None:
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), color)
    cell._tc.get_or_add_tcPr().append(node)


def margins(cell, value: int = 65) -> None:
    properties = cell._tc.get_or_add_tcPr()
    node = properties.first_child_found_in("w:tcMar")
    if node is None:
        node = OxmlElement("w:tcMar")
        properties.append(node)
    for edge in ("top", "start", "bottom", "end"):
        margin = OxmlElement(f"w:{edge}")
        margin.set(qn("w:w"), str(value))
        margin.set(qn("w:type"), "dxa")
        node.append(margin)


def repeat_header(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def prevent_split(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def cell_text(cell, value: object, *, bold: bool = False, white: bool = False, size: float = 7.5) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(str(value))
    run.font.name = "Liberation Sans"
    run.font.size = Pt(size)
    run.bold = bold
    if white:
        run.font.color.rgb = RGBColor(255, 255, 255)


def make_table(document: Document, headers: list[str], rows: list[list[object]], widths: list[float], size: float = 7.5):
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for index, header in enumerate(headers):
        cell_text(table.cell(0, index), header, bold=True, white=True, size=size)
        shade(table.cell(0, index), NAVY)
    repeat_header(table.rows[0])
    for row_index, values in enumerate(rows, start=1):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cell_text(cells[index], value, size=size)
            if row_index % 2 == 0:
                shade(cells[index], LIGHT)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "5")
        border.set(qn("w:color"), "B8C8D1")
        borders.append(border)
    table._tbl.tblPr.append(borders)
    for row in table.rows:
        prevent_split(row)
        for index, cell in enumerate(row.cells):
            cell.width = Inches(widths[index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            margins(cell)
    return table


def title(document: Document, text: str, *, size: float = 22) -> None:
    paragraph = document.add_paragraph()
    paragraph.style = document.styles["Title"]
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(text)
    run.font.name = "Liberation Sans"
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = RGBColor(23, 59, 87)


def heading(document: Document, text: str, *, level: int = 1) -> None:
    paragraph = document.add_paragraph()
    paragraph.style = document.styles[f"Heading {level}"]
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.space_after = Pt(3)
    run = paragraph.add_run(text)
    run.font.name = "Liberation Sans"
    run.font.size = Pt(14 if level == 1 else 10.5)
    run.font.bold = True
    run.font.color.rgb = RGBColor(23, 59, 87) if level == 1 else RGBColor(40, 111, 142)


def body_paragraph(document: Document, text: str, *, bold: bool = False, color: RGBColor | None = None, size: float = 8.5):
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(3)
    run = paragraph.add_run(text)
    run.font.name = "Liberation Sans"
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color
    return paragraph


def add_link(paragraph, text: str, anchor: str) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), BLUE)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.extend((color, underline))
    run.append(properties)
    value = OxmlElement("w:t")
    value.text = text
    run.append(value)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_bookmark(paragraph, name: str, identifier: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(identifier))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(identifier))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def page_break(document: Document) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.add_run().add_break(WD_BREAK.PAGE)


def build_cover(truth: dict[str, object]) -> Document:
    cover = Document()
    section = cover.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    title(cover, "Submission Identity\nControl Sheet")
    body_paragraph(cover, "DERIVED FILENAME • SOURCE PRESERVATION • WORD METADATA", bold=True, color=RGBColor(40, 111, 142), size=9)
    callout = make_table(cover, ["Source", "Derived output", "Identity rule"], [[INPUT_NAME, f"{truth['student_id']}.docx", "Exactly one A–Z letter + 8 digits"]], [2.35, 2.15, 2.55], 8)
    for cell in callout.rows[1].cells:
        shade(cell, GREEN)

    heading(cover, "Extracted Identity Record", level=1)
    identity = [
        ["Student ID", truth["student_id"], "Body paragraph 3 • run 2"],
        ["Course", truth["course"], "Body paragraph 1"],
        ["Assignment", truth["assignment"], "Body paragraph 1"],
        ["Academic term", truth["term"], "Body paragraph 2"],
        ["ID occurrence count", "1", "Whole original body"],
        ["ID pattern", "PASS", r"^[A-Z]\d{8}$"],
    ]
    make_table(cover, ["Field", "Extracted value", "Source / rule"], identity, [1.55, 2.25, 3.25], 7.8)

    heading(cover, "Filename & Package Controls", level=1)
    filename = [
        ["Derived filename", f"{truth['student_id']}.docx", "PASS: stem equals extracted ID"],
        ["Extension", ".docx", "PASS: editable OOXML"],
        ["Expected deliverables", "1", "No sidecar or alternate-ID DOCX"],
        ["Source SHA256", INPUT_SHA256, "Frozen OfficeBench input"],
    ]
    make_table(cover, ["Control", "Value", "Decision / evidence"], filename, [1.45, 2.85, 2.75], 7.1)

    heading(cover, "Preservation Baseline", level=1)
    preservation = [
        ["Original body", "13 paragraphs • 14 runs • 1,817 normalized characters"],
        ["Native numbering", "1 level-0 question • 3 level-1 options"],
        ["Original objects", "0 tables • 0 images • 0 hyperlinks • 1 section"],
        ["Normalized body SHA256", NORMALIZED_BODY_SHA256],
    ]
    make_table(cover, ["Baseline", "Frozen result"], preservation, [2.05, 5.0], 7.2)
    link_paragraph = body_paragraph(cover, "Navigation: ", bold=True, size=8.5)
    add_link(link_paragraph, "Jump to the original Student ID", "source_student_id")
    page_break(cover)
    return cover


def prepend_cover(document: Document, cover: Document) -> None:
    body = document._element.body
    insertion = 0
    for element in cover._element.body:
        if element.tag == qn("w:sectPr"):
            continue
        body.insert(insertion, deepcopy(element))
        insertion += 1


def add_appendix(document: Document, truth: dict[str, object], source_paragraphs) -> None:
    page_break(document)
    title(document, "Preservation Ledger & Closure", size=20)
    body_paragraph(document, "The ledger fingerprints the untouched assignment block after the identity-driven cover and before this appendix.", size=8)
    rows = []
    for index, (paragraph, role) in enumerate(zip(source_paragraphs, ROLES), start=1):
        text = clean(paragraph.text)
        rows.append([
            index,
            role,
            paragraph_numbering(paragraph),
            len(paragraph.runs),
            len(text),
            hashlib.sha256(text.encode("utf-8")).hexdigest()[:12],
            text[:34] + ("…" if len(text) > 34 else ""),
        ])
    heading(document, "Paragraph Preservation Ledger", level=1)
    make_table(document, ["#", "Role", "List state", "Runs", "Chars", "SHA-12", "Text locator"], rows, [0.32, 1.22, 1.0, 0.42, 0.43, 0.9, 2.1], 6.1)
    heading(document, "Independent Closure", level=1)
    closure = [
        ["Original paragraphs", "13", "Input body"],
        ["Preserved assignment paragraphs", "13", "Contiguous page-2 block"],
        ["Paragraph text matches", "13 / 13", "Exact normalized text"],
        ["Run text sequence", "14 / 14", "Exact source run boundaries"],
        ["Native numbering", "4", "1 level 0 + 3 level 1"],
        ["Student ID occurrences", "1", "Original assignment block"],
        ["Source native objects", "0 / 0 / 0", "Tables / images / hyperlinks"],
        ["Identity bookmark / link", "1 / 1", "source_student_id"],
        ["Custom properties", "6", "StudentID, Course, Assignment, Term, SourceSHA256, PreservedParagraphs"],
        ["Body fingerprint", NORMALIZED_BODY_SHA256, "13 normalized paragraphs joined by LF"],
    ]
    table = make_table(document, ["Control", "Result", "Evidence / rule"], closure, [1.65, 2.6, 2.8], 6.8)
    for cell in table.rows[3].cells:
        shade(cell, GREEN)
    heading(document, "Submission Boundary", level=2)
    body_paragraph(document, "The original assignment wording and numbering are preserved. This package does not add student answers, grades, instructor decisions, submission timestamps, or identity facts beyond the source document.", size=7.7)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Identity Control  •  Page ")
    run.font.name = "Liberation Sans"
    run.font.size = Pt(8)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    field_run = OxmlElement("w:r")
    field_text = OxmlElement("w:t")
    field_text.text = "1"
    field_run.append(field_text)
    field.append(field_run)
    paragraph._p.append(field)


def configure_header_footer(document: Document, student_id: str) -> None:
    section = document.sections[0]
    section.header_distance = Inches(0.35)
    section.footer_distance = Inches(0.35)
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = header.add_run(f"CSE 221  •  SUBMISSION CONTROL  •  {student_id}")
    run.font.name = "Liberation Sans"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor(40, 111, 142)
    add_page_number(section.footer.paragraphs[0])


def custom_properties_xml(truth: dict[str, object]) -> bytes:
    custom_ns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    vt_ns = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    root = etree.Element(f"{{{custom_ns}}}Properties", nsmap={None: custom_ns, "vt": vt_ns})
    values = [
        ("StudentID", str(truth["student_id"]), "lpwstr"),
        ("Course", str(truth["course"]), "lpwstr"),
        ("Assignment", str(truth["assignment"]), "lpwstr"),
        ("Term", str(truth["term"]), "lpwstr"),
        ("SourceSHA256", INPUT_SHA256, "lpwstr"),
        ("PreservedParagraphs", "13", "i4"),
    ]
    for pid, (name, value, kind) in enumerate(values, start=2):
        prop = etree.SubElement(root, f"{{{custom_ns}}}property")
        prop.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        prop.set("pid", str(pid))
        prop.set("name", name)
        child = etree.SubElement(prop, f"{{{vt_ns}}}{kind}")
        child.text = value
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def inject_custom_properties(path: Path, truth: dict[str, object]) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = {info.filename: (info, archive.read(info.filename)) for info in archive.infolist()}
    content_types = etree.fromstring(entries["[Content_Types].xml"][1])
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    if not content_types.xpath("/*[local-name()='Types']/*[local-name()='Override'][@PartName='/docProps/custom.xml']"):
        override = etree.SubElement(content_types, f"{{{ct_ns}}}Override")
        override.set("PartName", "/docProps/custom.xml")
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml")
    entries["[Content_Types].xml"] = (entries["[Content_Types].xml"][0], etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone="yes"))
    relationships = etree.fromstring(entries["_rels/.rels"][1])
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    if not relationships.xpath("/*[local-name()='Relationships']/*[local-name()='Relationship'][contains(@Type, '/custom-properties')]"):
        relation = etree.SubElement(relationships, f"{{{rel_ns}}}Relationship")
        relation.set("Id", "rId99")
        relation.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
        relation.set("Target", "docProps/custom.xml")
    entries["_rels/.rels"] = (entries["_rels/.rels"][0], etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone="yes"))
    entries["docProps/custom.xml"] = (None, custom_properties_xml(truth))
    with tempfile.TemporaryDirectory(prefix="identity-custom-") as temp_dir:
        target = Path(temp_dir) / path.name
        fixed = (1980, 1, 1, 0, 0, 0)
        with zipfile.ZipFile(target, "w") as outgoing:
            for name in sorted(entries):
                info, payload = entries[name]
                replacement = zipfile.ZipInfo(name, fixed)
                replacement.compress_type = zipfile.ZIP_DEFLATED
                replacement.external_attr = info.external_attr if info is not None else 0o600 << 16
                outgoing.writestr(replacement, payload)
        os.replace(target, path)


def build(input_path: Path, output_dir: Path) -> Path:
    if hashlib.sha256(input_path.read_bytes()).hexdigest() != INPUT_SHA256:
        raise ValueError("Source does not match the frozen OfficeBench document")
    document = Document(input_path)
    source_paragraphs = list(document.paragraphs)
    truth = extract_truth(document)
    if truth["body_sha256"] != NORMALIZED_BODY_SHA256:
        raise ValueError("Original assignment text fingerprint mismatch")
    add_bookmark(source_paragraphs[2], "source_student_id", 7301)
    prepend_cover(document, build_cover(truth))
    add_appendix(document, truth, source_paragraphs)
    configure_header_footer(document, str(truth["student_id"]))
    core = document.core_properties
    core.title = f"{truth['course']} {truth['assignment']} Submission Control"
    core.subject = f"Identity-preserved submission package for {truth['student_id']}"
    core.author = "CoworkEvoBench Oracle"
    core.category = "Submission identity control"
    core.keywords = "student ID, filename control, source preservation, custom properties"
    core.comments = "The 13-paragraph assignment block is preserved from the frozen source."
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / f"{truth['student_id']}.docx"
    document.save(output)
    inject_custom_properties(output, truth)
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=Path(os.environ.get("INPUT_DIR", "/app/data/input_files")) / INPUT_NAME)
    parser.add_argument("--output-dir", type=Path, default=Path(os.environ.get("OUTPUT_DIR", "/app/output")))
    args = parser.parse_args()
    build(args.input, args.output_dir)


if __name__ == "__main__":
    main()
