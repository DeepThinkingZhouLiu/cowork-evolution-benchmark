from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree


INPUT_NAME = "homework.jpg"
OUTPUT_NAME = "homework.docx"
INPUT_SHA256 = "1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf"
NAVY = "173B57"
BLUE = "286F8E"
LIGHT = "F2F6F8"
GREEN = "E3F1E8"
GOLD = "FFF1CC"


def load_truth() -> dict[str, object]:
    candidates = [Path(__file__).parents[1] / "adaptation/transcription-truth.json", Path(__file__).with_name("transcription-truth.json")]
    for path in candidates:
        if path.is_file():
            return json.loads(path.read_text(encoding="utf-8"))
    raise FileNotFoundError("transcription truth is unavailable")


def shade(cell, color: str) -> None:
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), color)
    cell._tc.get_or_add_tcPr().append(node)


def margins(cell, value: int = 55) -> None:
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


def cell_text(cell, value: object, *, bold: bool = False, white: bool = False, size: float = 7.0, align=WD_ALIGN_PARAGRAPH.LEFT) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(str(value))
    run.font.name = "Liberation Sans"
    run.font.size = Pt(size)
    run.bold = bold
    if white:
        run.font.color.rgb = RGBColor(255, 255, 255)


def make_table(document: Document, headers: list[str], rows: list[list[object]], widths: list[float], size: float = 7.0, center_after_first: bool = False):
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for index, header in enumerate(headers):
        alignment = WD_ALIGN_PARAGRAPH.CENTER if center_after_first and index else WD_ALIGN_PARAGRAPH.LEFT
        cell_text(table.cell(0, index), header, bold=True, white=True, size=size, align=alignment)
        shade(table.cell(0, index), NAVY)
    repeat_header(table.rows[0])
    for row_index, values in enumerate(rows, start=1):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            alignment = WD_ALIGN_PARAGRAPH.CENTER if center_after_first and index else WD_ALIGN_PARAGRAPH.LEFT
            cell_text(cells[index], value, size=size, align=alignment)
            if row_index % 2 == 0:
                shade(cells[index], LIGHT)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "5")
        border.set(qn("w:color"), "AABCC7")
        borders.append(border)
    table._tbl.tblPr.append(borders)
    for row in table.rows:
        prevent_split(row)
        for index, cell in enumerate(row.cells):
            cell.width = Inches(widths[index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            margins(cell)
    return table


def set_font(run, *, size: float, bold: bool = False, italic: bool = False, color: RGBColor | None = None, name: str = "Liberation Serif") -> None:
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def paragraph(document: Document, text: str = "", *, size: float = 8.5, bold: bool = False, align=WD_ALIGN_PARAGRAPH.LEFT, before: float = 0, after: float = 2, left: float = 0):
    item = document.add_paragraph()
    item.alignment = align
    item.paragraph_format.space_before = Pt(before)
    item.paragraph_format.space_after = Pt(after)
    item.paragraph_format.left_indent = Inches(left)
    if text:
        set_font(item.add_run(text), size=size, bold=bold)
    return item


def heading(document: Document, text: str, *, size: float = 16, level: int = 1, after: float = 3):
    item = document.add_paragraph(style=f"Heading {level}")
    item.paragraph_format.space_before = Pt(4)
    item.paragraph_format.space_after = Pt(after)
    set_font(item.add_run(text), size=size, bold=True, color=RGBColor(23, 59, 87), name="Liberation Sans")
    return item


def add_bookmark(item, name: str, identifier: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(identifier))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(identifier))
    item._p.insert(0, start)
    item._p.append(end)


def add_internal_link(item, label: str, anchor: str) -> None:
    link = OxmlElement("w:hyperlink")
    link.set(qn("w:anchor"), anchor)
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), BLUE)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.extend((color, underline))
    run.append(properties)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    link.append(run)
    item._p.append(link)


def add_equation(document: Document, expression: str) -> None:
    item = document.add_paragraph()
    item.alignment = WD_ALIGN_PARAGRAPH.CENTER
    item.paragraph_format.space_before = Pt(1)
    item.paragraph_format.space_after = Pt(2)
    math = OxmlElement("m:oMathPara")
    equation = OxmlElement("m:oMath")
    run = OxmlElement("m:r")
    text = OxmlElement("m:t")
    text.text = expression
    run.append(text)
    equation.append(run)
    math.append(equation)
    item._p.append(math)


def add_page_field(item, prefix: str) -> None:
    item.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(item.add_run(prefix), size=7.5, name="Liberation Sans")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    field.append(run)
    item._p.append(field)


def configure(document: Document) -> None:
    section = document.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.45)
    section.left_margin = Inches(0.48)
    section.right_margin = Inches(0.48)
    section.header_distance = Inches(0.22)
    section.footer_distance = Inches(0.22)
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    set_font(header.add_run("CSE 240A  •  EDITABLE OCR RECONSTRUCTION  •  A012345"), size=7.2, bold=True, color=RGBColor(40, 111, 142), name="Liberation Sans")
    add_page_field(section.footer.paragraphs[0], "homework.jpg  •  Page ")
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Serif"
    normal.font.size = Pt(8.5)


def page_break(document: Document) -> None:
    item = document.add_paragraph()
    item.paragraph_format.space_after = Pt(0)
    item.add_run().add_break(WD_BREAK.PAGE)


def question(document: Document, number: str, prompt: str, bookmark: str) -> None:
    item = paragraph(document, size=7.8, after=1.5, left=0.12)
    set_font(item.add_run(f"{number}. "), size=7.8, bold=True)
    set_font(item.add_run(prompt), size=7.8)
    add_bookmark(item, bookmark, 6200 + int(number))


def build_reconstruction(document: Document, truth: dict[str, object]) -> None:
    header = truth["header"]
    title = paragraph(document, header["title"], size=18, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
    title.style = document.styles["Title"]
    title.runs[0].bold = False
    paragraph(document, header["subtitle"], size=11.5, align=WD_ALIGN_PARAGRAPH.CENTER, after=8)
    paragraph(document, header["identity"], size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    paragraph(document, header["due"], size=10, align=WD_ALIGN_PARAGRAPH.CENTER, after=7)
    paragraph(document, "Instructions", size=9.5, bold=True, after=2)
    for text in truth["instructions"]:
        item = document.add_paragraph(style="List Bullet")
        item.paragraph_format.space_after = Pt(0.6)
        item.paragraph_format.left_indent = Inches(0.2)
        item.paragraph_format.first_line_indent = Inches(-0.1)
        set_font(item.add_run(text), size=7.7)
    section = paragraph(document, truth["section_heading"], size=12, bold=True, before=2, after=2)
    add_bookmark(section, "pipelining_section", 6199)

    first = truth["questions"][0]
    question(document, first["number"], first["prompt"], "question_1")
    paragraph(document, first["note"], size=7.4, after=2, left=0.23)
    cycles = truth["pipeline"]["cycles"]
    rows = []
    arrow_cells = {
        ("lw R2, 0 (R5)", "5"): "W\n↓",
        ("add R3, R2, R2", "8"): "W\n↓",
        ("add R2, R2, R5", "9"): "M\n↓",
    }
    for source_row in truth["pipeline"]["rows"]:
        instruction = source_row["instruction"]
        values = [instruction]
        for cycle in cycles:
            values.append(arrow_cells.get((instruction, cycle), source_row["stages"].get(cycle, "")))
        rows.append(values)
    pipeline = make_table(document, ["Instruction", *cycles], rows, [1.72, *([0.35] * 15)], 6.0, center_after_first=True)
    for cell in pipeline.rows[1].cells:
        shade(cell, GREEN)

    second = truth["questions"][1]
    question(document, second["number"], second["prompt"], "question_2")
    paragraph(document, second["note"], size=7.4, after=3, left=0.23)

    third = truth["questions"][2]
    question(document, third["number"], third["prompt"], "question_3")
    add_equation(document, third["equation"])
    paragraph(document, third["conclusion"], size=8.2, after=0, left=0.23)


def segment_rows(truth: dict[str, object]) -> list[list[str]]:
    segments = [
        ("H1", "title", truth["header"]["title"], "top center"),
        ("H2", "subtitle", truth["header"]["subtitle"], "below title"),
        ("H3", "identity", truth["header"]["identity"], "header identity line"),
        ("H4", "due", truth["header"]["due"], "header due line"),
        ("S0", "label", "Instructions", "left heading"),
    ]
    segments.extend((f"I{index}", "instruction bullet", text, f"bullet {index}") for index, text in enumerate(truth["instructions"], 1))
    segments.append(("S1", "section heading", truth["section_heading"], "pipelining heading"))
    for item in truth["questions"]:
        segments.append((f"Q{item['number']}", "question prompt", f"{item['number']}. {item['prompt']}", f"visible prompt {item['number']}"))
        if "note" in item:
            segments.append((f"N{item['number']}", "question note", item["note"], f"note below prompt {item['number']}"))
        if "equation" in item:
            segments.append(("E3", "worked equation", item["equation"], "centered under Q3"))
            segments.append(("C3", "conclusion", item["conclusion"], "below equation"))
    rows = []
    for key, role, text, locator in segments:
        rows.append([key, role, str(len(text)), hashlib.sha256(text.encode("utf-8")).hexdigest()[:12], text[:42] + ("…" if len(text) > 42 else ""), locator])
    return rows


def build_evidence_page(document: Document, truth: dict[str, object]) -> None:
    page_break(document)
    heading(document, "Transcription Evidence Ledger", size=19)
    paragraph(document, "Every visible source segment is represented as editable Word text. Fingerprints use the verified transcription below; the raster scan is evidence, not an embedded output page.", size=8, after=3)
    source = truth["source"]
    anchor_rows = [
        ["Source file", source["filename"], "Frozen runtime input"],
        ["SHA256", source["sha256"], "Exact package/environment bytes"],
        ["Raster geometry", f"{source['width_px']} × {source['height_px']} px", "JPEG; one visible page"],
        ["Reconstruction", "Page 1", "Editable text + native pipeline table"],
        ["Embedded source image", "0", "No screenshot or flattened-page fallback"],
    ]
    anchor = make_table(document, ["Anchor", "Verified value", "Use"], anchor_rows, [1.2, 3.55, 2.7], 7.0)
    for cell in anchor.rows[2].cells:
        shade(cell, GREEN)
    heading(document, "Segment Fingerprints", size=13, level=2)
    make_table(document, ["Key", "Role", "Chars", "SHA-12", "Text locator", "Scan location"], segment_rows(truth), [0.38, 1.05, 0.42, 0.86, 2.85, 1.55], 5.9)
    paragraph(document, "Boundary: punctuation, capitalization, numeric tokens, instruction contradictions, and visible worked content are retained. No missing answer, diagram, or grading fact is inferred.", size=7.3, bold=True, after=0)


def occupied_cycles(row: dict[str, object]) -> str:
    return " · ".join(f"{cycle}:{stage}" for cycle, stage in row["stages"].items())


def build_audit_page(document: Document, truth: dict[str, object]) -> None:
    page_break(document)
    heading(document, "Pipeline Matrix Audit & Exception Closure", size=18)
    nav = paragraph(document, "Navigate to source reconstruction: ", size=8, bold=True, after=3)
    for index in (1, 2, 3):
        add_internal_link(nav, f"Q{index}", f"question_{index}")
        if index != 3:
            set_font(nav.add_run("  •  "), size=8, name="Liberation Sans")

    audit_rows = []
    arrows = truth["pipeline"]["visible_down_arrows"]
    arrow_by_instruction = {row["instruction"]: "—" for row in truth["pipeline"]["rows"]}
    arrow_by_instruction["lw R2, 0 (R5)"] = arrows[0]
    arrow_by_instruction["add R3, R2, R2"] = arrows[1]
    arrow_by_instruction["add R2, R2, R5"] = arrows[2]
    for row in truth["pipeline"]["rows"]:
        stalls = "D@3–4" if row["instruction"] == "sub R10, R2, 10" else ("F@3–4" if row["instruction"] == "add R3, R2, R2" else "—")
        audit_rows.append([row["instruction"], occupied_cycles(row), stalls, arrow_by_instruction[row["instruction"]]])
    heading(document, "Six-Instruction Pipeline Closure", size=12.5, level=2)
    make_table(document, ["Instruction", "Occupied cycle:stage cells", "Visible stall", "Down-arrow relation"], audit_rows, [1.25, 2.3, 0.85, 3.05], 6.1)

    heading(document, "Source Conflicts — Preserve, Do Not Repair", size=12.5, level=2)
    conflict_rows = [[item["issue"], item["source_evidence"], item["treatment"]] for item in truth["source_conflicts"]]
    conflict_table = make_table(document, ["Issue", "Source evidence", "Required treatment"], conflict_rows, [1.3, 3.15, 3.0], 6.5)
    for cell in conflict_table.rows[1].cells + conflict_table.rows[2].cells:
        shade(cell, GOLD)

    heading(document, "Verified Transcription Conventions", size=12.5, level=2)
    make_table(document, ["Token family", "Verified rendering", "Disambiguation"], truth["transcription_conventions"], [1.3, 2.0, 4.15], 6.6)

    heading(document, "Package Closure", size=12.5, level=2)
    closure = [
        ["Visible source pages", "1", "homework.jpg"],
        ["Editable reconstruction pages", "1", "Output page 1"],
        ["Question prompts", "3", "Preserved despite source count conflict"],
        ["Instruction bullets", "5", "Exact visible order"],
        ["Pipeline rows / cycles", "6 / 15", "Native 7 × 16 Word table"],
        ["Occupied stage cells", "34", "Includes repeated F/D stalls"],
        ["Visible dependency arrows", "3", "Explicit relations retained"],
        ["Native equation objects", "1", "Question 3 worked calculation"],
        ["Bookmarks / internal links", "4 / 3", "Section + Q1/Q2/Q3; links to questions"],
        ["Custom properties", "6", "SourceSHA256, Course, Homework, StudentName, StudentID, DueDate"],
    ]
    table = make_table(document, ["Control", "Result", "Evidence / rule"], closure, [2.0, 1.0, 4.45], 6.4)
    for cell in table.rows[6].cells:
        shade(cell, GREEN)


def custom_properties_xml(truth: dict[str, object]) -> bytes:
    custom_ns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    vt_ns = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    root = etree.Element(f"{{{custom_ns}}}Properties", nsmap={None: custom_ns, "vt": vt_ns})
    values = [
        ("SourceSHA256", INPUT_SHA256),
        ("Course", "CSE 240A"),
        ("Homework", "Homework 2"),
        ("StudentName", "Ryan Wang"),
        ("StudentID", "A012345"),
        ("DueDate", "June 1, 2021"),
    ]
    for pid, (name, value) in enumerate(values, start=2):
        prop = etree.SubElement(root, f"{{{custom_ns}}}property")
        prop.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        prop.set("pid", str(pid))
        prop.set("name", name)
        child = etree.SubElement(prop, f"{{{vt_ns}}}lpwstr")
        child.text = value
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def inject_custom_properties(path: Path, truth: dict[str, object]) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = {info.filename: (info, archive.read(info.filename)) for info in archive.infolist()}
    content_types = etree.fromstring(entries["[Content_Types].xml"][1])
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    override = etree.SubElement(content_types, f"{{{ct_ns}}}Override")
    override.set("PartName", "/docProps/custom.xml")
    override.set("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml")
    entries["[Content_Types].xml"] = (entries["[Content_Types].xml"][0], etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone="yes"))
    relationships = etree.fromstring(entries["_rels/.rels"][1])
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    relation = etree.SubElement(relationships, f"{{{rel_ns}}}Relationship")
    relation.set("Id", "rId99")
    relation.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
    relation.set("Target", "docProps/custom.xml")
    entries["_rels/.rels"] = (entries["_rels/.rels"][0], etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone="yes"))
    entries["docProps/custom.xml"] = (None, custom_properties_xml(truth))
    with tempfile.TemporaryDirectory(prefix="ocr-custom-") as temp_dir:
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
        raise ValueError("Input does not match the frozen OfficeBench scan")
    truth = load_truth()
    if truth["source"]["sha256"] != INPUT_SHA256:
        raise ValueError("Transcription ledger and source snapshot disagree")
    document = Document()
    configure(document)
    build_reconstruction(document, truth)
    build_evidence_page(document, truth)
    build_audit_page(document, truth)
    core = document.core_properties
    core.title = "CSE 240A Homework 2 Editable OCR Reconstruction"
    core.subject = "Verified transcription, native pipeline matrix, and source-conflict audit"
    core.author = "CoworkEvoBench Oracle"
    core.category = "OCR reconstruction control"
    core.keywords = "CSE 240A, Homework 2, OCR, pipeline matrix, source conflict"
    core.comments = "Source-visible text is transcribed without inventing the missing Question 2 diagram."
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_NAME
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
