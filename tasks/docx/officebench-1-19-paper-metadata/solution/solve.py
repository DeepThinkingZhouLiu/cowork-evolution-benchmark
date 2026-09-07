from __future__ import annotations

import argparse
import subprocess
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree


OUTPUT_NAME = "os_copilot_research_intake.docx"
TITLE = "OS-Copilot: Towards Generalist Computer Agents with Self-Improvement"
AUTHORS = [
    ("1", "Zhiyong Wu", "Shanghai AI Laboratory", "Equal contribution; listed contact"),
    ("2", "Chengcheng Han", "East China Normal University", "Equal contribution"),
    ("3", "Zichen Ding", "East China Normal University", "—"),
    ("4", "Zhenmin Weng", "East China Normal University", "—"),
    ("5", "Zhoumianze Liu", "Shanghai AI Laboratory", "—"),
    ("6", "Shunyu Yao", "Princeton University", "—"),
    ("7", "Tao Yu", "The University of Hong Kong", "—"),
    ("8", "Lingpeng Kong", "The University of Hong Kong", "—"),
]

EVIDENCE = [
    ("E01", "Identity", TITLE, "Page 1, title block", "Verified"),
    ("E02", "Authors", "Eight authors in displayed order", "Page 1, author line", "Verified"),
    ("E03", "Affiliations", "2 Shanghai AI Lab; 3 ECNU; 1 Princeton; 2 HKU", "Page 1, symbol legend", "Verified"),
    ("E04", "Contribution note", "Zhiyong Wu and Chengcheng Han are marked equal contribution", "Page 1, * footnote", "Verified"),
    ("E05", "Contact", "wuzhiyong@pjlab.org.cn; https://os-copilot.github.io/", "Page 1, title block", "Verified"),
    ("E06", "Research gap", "Narrow software/site agents constrain general computer-task applicability", "Page 1, abstract", "Verified"),
    ("E07", "Framework", "OS-Copilot interfaces with comprehensive operating-system elements", "Page 1, abstract", "Verified"),
    ("E08", "Agent", "FRIDAY is the self-improving embodied agent created with OS-Copilot", "Page 1, abstract", "Verified"),
    ("E09", "Coverage", "Web, code terminals, files, multimedia, and third-party applications", "Page 1, abstract", "Verified"),
    ("E10", "Reported result", "On GAIA, FRIDAY outperforms previous methods by 35%", "Page 1, abstract", "Verified; baseline unnamed here"),
    ("E11", "Learning claim", "Excel and PowerPoint control/self-improvement with minimal supervision", "Page 1, abstract", "Verified"),
    ("E12", "Figure 1", "Focused environment; Excel calculation/chart; OS-Copilot website", "Page 1, figure caption", "Verified"),
]

COMMENTS = [
    (
        "Only Zhiyong Wu and Chengcheng Han carry the equal-contribution marker.",
        "The asterisk appears beside these two names only; do not extend the note to the remaining authors.",
    ),
    (
        "35% uplift is reported against previous methods; the one-page source does not name the baseline.",
        "Keep the claim qualified. The frozen page supports the percentage but not a baseline identity or uncertainty interval.",
    ),
    (
        "The frozen PDF contains one page, so methods, detailed evaluation tables, limitations, and references are outside this intake evidence.",
        "This is an evidence boundary, not a conclusion that the full paper lacks those sections.",
    ),
]


def source_text(path: Path) -> str:
    return subprocess.check_output(["pdftotext", "-layout", str(path), "-"], text=True)


def verify_source(path: Path) -> None:
    text = " ".join(source_text(path).split())
    required = [
        "OS-Copilot", "Zhiyong Wu", "Chengcheng Han", "Shanghai AI Laboratory",
        "Princeton University", "The University of Hong Kong", "wuzhiyong@pjlab.org.cn",
        "2402.07456v2", "15 Feb 2024", "outperforms previous methods by 35%",
        "Excel", "Powerpoint", "focused working environment", "creating a website for OS-Copilot",
    ]
    assert all(value in text for value in required)
    info = subprocess.check_output(["pdfinfo", str(path)], text=True)
    assert "Pages:           1" in info and "Tagged:          no" in info


def shade(cell, color: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    element = properties.find(qn("w:shd")) or OxmlElement("w:shd")
    element.set(qn("w:fill"), color)
    if element.getparent() is None:
        properties.append(element)


def repeat_header(row) -> None:
    element = OxmlElement("w:tblHeader")
    element.set(qn("w:val"), "true")
    row._tr.get_or_add_trPr().append(element)


def no_split(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def put(cell, value: str, bold: bool = False, size: float = 8.5) -> None:
    cell.text = ""
    run = cell.paragraphs[0].add_run(value)
    run.bold = bold
    run.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def bookmark(paragraph, name: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
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


def field(paragraph, instruction: str, display: str = "1") -> None:
    element = OxmlElement("w:fldSimple")
    element.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = display
    run.append(text)
    element.append(run)
    paragraph._p.append(element)


def inject_comments(path: Path) -> None:
    namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    content_type_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    relation_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    temporary = path.with_suffix(".comments.docx")
    with zipfile.ZipFile(path, "r") as source:
        parts = {info.filename: source.read(info.filename) for info in source.infolist()}
        infos = {info.filename: info for info in source.infolist()}
    document = etree.fromstring(parts["word/document.xml"])
    comments = etree.Element(qn("w:comments"), nsmap={"w": namespaces["w"]})
    for comment_id, (anchor, note) in enumerate(COMMENTS):
        nodes = document.xpath("//w:t[text()=$value]", namespaces=namespaces, value=anchor)
        assert len(nodes) == 1, (anchor, len(nodes))
        run = nodes[0].getparent()
        parent = run.getparent()
        position = parent.index(run)
        start = OxmlElement("w:commentRangeStart")
        start.set(qn("w:id"), str(comment_id))
        end = OxmlElement("w:commentRangeEnd")
        end.set(qn("w:id"), str(comment_id))
        reference_run = OxmlElement("w:r")
        reference = OxmlElement("w:commentReference")
        reference.set(qn("w:id"), str(comment_id))
        reference_run.append(reference)
        parent.insert(position, start)
        parent.insert(position + 2, end)
        parent.insert(position + 3, reference_run)
        comment = etree.SubElement(comments, qn("w:comment"))
        comment.set(qn("w:id"), str(comment_id))
        comment.set(qn("w:author"), "Research Intake QA")
        paragraph = etree.SubElement(comment, qn("w:p"))
        comment_run = etree.SubElement(paragraph, qn("w:r"))
        comment_text = etree.SubElement(comment_run, qn("w:t"))
        comment_text.text = note
    parts["word/document.xml"] = etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone="yes")
    parts["word/comments.xml"] = etree.tostring(comments, xml_declaration=True, encoding="UTF-8", standalone="yes")
    content_types = etree.fromstring(parts["[Content_Types].xml"])
    if not content_types.xpath("//*[local-name()='Override'][@PartName='/word/comments.xml']"):
        override = etree.SubElement(content_types, f"{{{content_type_ns}}}Override")
        override.set("PartName", "/word/comments.xml")
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    parts["[Content_Types].xml"] = etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone="yes")
    relationships = etree.fromstring(parts["word/_rels/document.xml.rels"])
    if not relationships.xpath("//*[local-name()='Relationship'][contains(@Type, '/comments')]"):
        used = {item.get("Id") for item in relationships}
        number = 1
        while f"rId{number}" in used:
            number += 1
        relationship = etree.SubElement(relationships, f"{{{relation_ns}}}Relationship")
        relationship.set("Id", f"rId{number}")
        relationship.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
        relationship.set("Target", "comments.xml")
    parts["word/_rels/document.xml.rels"] = etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone="yes")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for name, payload in parts.items():
            if name in infos:
                target.writestr(infos[name], payload)
            else:
                target.writestr(name, payload)
    temporary.replace(path)


def build(input_dir: Path, output_dir: Path) -> None:
    pdf = input_dir / "paper.pdf"
    verify_source(pdf)
    document = Document()
    document.styles["Normal"].font.name = "Aptos"
    document.styles["Normal"].font.size = Pt(9)
    document.styles["Title"].font.color.rgb = RGBColor(31, 78, 121)
    section = document.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.62)
    section.left_margin = section.right_margin = Inches(0.68)
    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("OS-Copilot Research Intake & Evidence Review")
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run("Evidence boundary: frozen one-page PDF | Intake date: 2026-09-08").bold = True
    document.add_heading("Navigation", 1)
    field(document.add_paragraph(), 'TOC \\o "1-2" \\h \\z \\u')

    document.add_heading("1. Bibliographic identity", 1)
    identity = document.add_table(rows=8, cols=2)
    identity.style = "Table Grid"
    identity_rows = [
        ("Displayed title", TITLE), ("arXiv identity", "arXiv:2402.07456v2 [cs.AI]"),
        ("Displayed date", "15 Feb 2024"), ("Listed contact", "wuzhiyong@pjlab.org.cn"),
        ("Project URL", "https://os-copilot.github.io/"), ("Frozen file extent", "1 page; Letter; PDF 1.5"),
        ("Embedded Title/Author", "Both blank in PDF properties"), ("Accessibility flag", "Tagged: no"),
    ]
    for row, values in zip(identity.rows, identity_rows):
        put(row.cells[0], values[0], True, 9); shade(row.cells[0], "D9EAF7"); put(row.cells[1], values[1], False, 9); no_split(row)

    document.add_heading("2. Ordered authorship and affiliation map", 1)
    authors = document.add_table(rows=1, cols=4); authors.style = "Table Grid"
    for cell, label in zip(authors.rows[0].cells, ["Order", "Author", "Affiliation", "Source-bounded note"]):
        put(cell, label, True); shade(cell, "1F4E79")
    repeat_header(authors.rows[0])
    for values in AUTHORS:
        row = authors.add_row()
        for cell, value in zip(row.cells, values): put(cell, value)
        no_split(row)
    document.add_paragraph(COMMENTS[0][0])
    document.add_paragraph("The page lists one contact address; it does not label a corresponding author.")

    landscape = document.add_section(WD_ORIENT.LANDSCAPE)
    landscape.orientation = WD_ORIENT.LANDSCAPE
    landscape.page_width, landscape.page_height = Inches(11.69), Inches(8.27)
    landscape.top_margin = landscape.bottom_margin = Inches(0.52)
    landscape.left_margin = landscape.right_margin = Inches(0.46)
    document.add_heading("3. Evidence control matrix", 1)
    matrix = document.add_table(rows=1, cols=5); matrix.style = "Table Grid"
    for cell, label in zip(matrix.rows[0].cells, ["ID", "Control", "Evidence-bound finding", "Source locator", "Status"]):
        put(cell, label, True, 7.5); shade(cell, "1F4E79")
    repeat_header(matrix.rows[0])
    for values in EVIDENCE:
        row = matrix.add_row()
        for cell, value in zip(row.cells, values): put(cell, value, False, 7.5)
        row.cells[0].text = ""; link(row.cells[0].paragraphs[0], values[0], f"evidence_{values[0]}")
        shade(row.cells[4], "E2F0D9"); no_split(row)
    document.add_paragraph(COMMENTS[1][0])

    portrait = document.add_section(WD_ORIENT.PORTRAIT)
    portrait.orientation = WD_ORIENT.PORTRAIT
    portrait.page_width, portrait.page_height = Inches(8.5), Inches(11)
    portrait.top_margin = portrait.bottom_margin = Inches(0.62)
    portrait.left_margin = portrait.right_margin = Inches(0.68)
    document.add_heading("4. Structured abstract review", 1)
    abstract = document.add_table(rows=5, cols=2); abstract.style = "Table Grid"
    abstract_rows = [
        ("Problem", "Autonomous computer interaction is longstanding and potentially valuable."),
        ("Gap", "Most agents target one software product or website, limiting general computer-task use."),
        ("Proposed system", "OS-Copilot supplies the generalist framework; FRIDAY is its self-improving embodied agent."),
        ("Evidence on this page", "FRIDAY is reported to beat previous GAIA methods by 35% and to learn Excel and PowerPoint control with minimal supervision."),
        ("Claimed significance", "The framework and findings are positioned as infrastructure and insight for more capable general-purpose computer agents."),
    ]
    for row, values in zip(abstract.rows, abstract_rows):
        put(row.cells[0], values[0], True, 9); shade(row.cells[0], "D9EAF7"); put(row.cells[1], values[1], False, 9); no_split(row)

    document.add_heading("5. Figure 1 reconstruction checklist", 1)
    figure = document.add_table(rows=1, cols=3); figure.style = "Table Grid"
    for cell, label in zip(figure.rows[0].cells, ["Scenario", "Requested outcome", "Caption reading rule"]):
        put(cell, label, True); shade(cell, "1F4E79")
    repeat_header(figure.rows[0])
    figure_rows = [
        ("1", "Prepare a focused working environment", "Top = initial OS state; bottom = final state"),
        ("2", "Calculate and draw a chart in Excel", "Bottom text gives FRIDAY's subtasks"),
        ("3", "Create a website for OS-Copilot", "Boxes/ovals highlight changes"),
    ]
    for values in figure_rows:
        row = figure.add_row()
        for cell, value in zip(row.cells, values): put(cell, value)
        no_split(row)

    document.add_heading("6. Intake limitations and follow-up", 1)
    risks = document.add_table(rows=1, cols=3); risks.style = "Table Grid"
    for cell, label in zip(risks.rows[0].cells, ["Issue", "Evidence consequence", "Follow-up"]):
        put(cell, label, True); shade(cell, "1F4E79")
    repeat_header(risks.rows[0])
    risk_rows = [
        ("One-page source", "Methods, evaluation detail, limitations, and references are not in scope", "Request the complete paper before technical review"),
        ("Blank embedded metadata", "PDF Title and Author properties cannot be used as authority", "Use displayed title block and retain this discrepancy"),
        ("Untagged PDF", "Reading order/accessibility is not machine-declared", "Run accessibility remediation on any publication derivative"),
        ("Qualified metric", "35% is visible; baseline identity and uncertainty are absent", "Do not name a baseline or add statistical detail"),
    ]
    for values in risk_rows:
        row = risks.add_row()
        for cell, value in zip(row.cells, values): put(cell, value)
        no_split(row)
    document.add_paragraph(COMMENTS[2][0])
    document.add_paragraph("Do not add a DOI, venue, publisher, corresponding-author designation, benchmark baseline, confidence interval, or claims from outside the frozen page.")

    document.add_heading("7. Evidence register", 1)
    for index, values in enumerate(EVIDENCE, 1):
        paragraph = document.add_paragraph(style="List Number")
        paragraph.add_run(f"{values[0]} | {values[3]} | ").bold = True
        paragraph.add_run(values[2])
        bookmark(paragraph, f"evidence_{values[0]}", index)

    for section in document.sections:
        section.header.is_linked_to_previous = False; section.footer.is_linked_to_previous = False
        header = section.header.paragraphs[0]; header.alignment = WD_ALIGN_PARAGRAPH.RIGHT; header.text = "OS-Copilot Research Intake | Source: paper.pdf"
        footer = section.footer.paragraphs[0]; footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run("Page "); field(footer, "PAGE"); footer.add_run(" of "); field(footer, "NUMPAGES")
    document.core_properties.title = "OS-Copilot Research Intake & Evidence Review"
    document.core_properties.subject = "Evidence-bounded bibliographic and abstract intake from paper.pdf"
    document.core_properties.author = "Research Intake QA"
    document.core_properties.keywords = "OS-Copilot, FRIDAY, paper intake, evidence review"
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_NAME
    document.save(output)
    inject_comments(output)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", required=True)
    parser.add_argument("--output-dir", required=True)
    arguments = parser.parse_args()
    build(Path(arguments.input_dir), Path(arguments.output_dir))


if __name__ == "__main__":
    main()
