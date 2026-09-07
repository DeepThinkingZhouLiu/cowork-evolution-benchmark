#!/usr/bin/env python3
"""制作本题专用的课程政策整合红线工作稿。"""

from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.shared import Inches, Pt
from lxml import etree


OUTPUT_NAME = "sample_syllabus_policy_review.docx"
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
CHANGE_DATE = datetime(2026, 9, 7, 14, 0, tzinfo=timezone.utc).isoformat()

COMMENT_SPECS = [
    (
        "falsification or sabotage of research data",
        "The general integrity clause identifies research-data misconduct. The inserted research guidance adds conduct controls for work involving information from people; it does not replace this violation rule.",
        "general_research_misconduct",
    ),
    (
        "verified using Turnitin technology",
        "This sentence addresses originality verification for written work. It should not be read as satisfying the separate participant-selection, risk-benefit, or anonymity controls in the research guidance.",
        None,
    ),
    (
        "apply to your interactions with the staff of the schools as well as the community organizations and its clients.",
        "This clause extends the University policy and Code to field and community interactions. The inserted guidance supplies the additional human-subject research principles for data collected in those settings.",
        "off_campus_application",
    ),
]

TRACKED_BOOKMARKS = {
    1: "research_data_scope",
    4: "research_study_purpose",
    5: "research_participant_selection",
    6: "research_risk_benefit",
    7: "research_anonymity",
}

SCOPE_DECISION = (
    "The four policy components should be read cumulatively. The general Academic Integrity section identifies misconduct, "
    "including fabrication, falsification or sabotage of research data, and explains that violations may lead to sanctions. "
    "The Turnitin paragraph is narrower: it describes originality verification for written work and the student's pledge and "
    "consent to that verification. The inserted research guidance addresses a different operational risk. It treats information "
    "collected from students, teachers or classmates—whether oral, written, observed or taken from records or artifacts—as "
    "research data, then requires a clear purpose, unbiased participant selection, a risk-benefit assessment and anonymity. "
    "Finally, the off-campus section carries the University's Academic Integrity Policy and Code of Student Responsibility into "
    "interactions with partner-school staff, community organizations and clients. None of these clauses cancels another: "
    "Turnitin does not substitute for human-subject safeguards, and off-campus application does not narrow the research-data "
    "definition. This remains a policy-review template, not a representation that course-specific placeholders are complete."
)

MATRIX_ROWS = [
    (
        "General research-data misconduct",
        "The general policy lists fabrication, falsification or sabotage of research data among integrity violations.",
        "Keep the general misconduct and sanction framework alongside the research-specific controls.",
        "general_research_misconduct",
    ),
    (
        "Research-data scope",
        "Information collected from students, teachers or classmates—orally, in writing, by observation, or from records or artifacts—is research data.",
        "Handle all such information professionally and conduct the work ethically.",
        "research_data_scope",
    ),
    (
        "Study purpose",
        "The study must have clear questions and a clear link between the data sought and those questions.",
        "Define the questions and explain how the requested data will answer them.",
        "research_study_purpose",
    ),
    (
        "Unbiased participant selection",
        "Inclusion and exclusion must be appropriate and unbiased, not based on extraneous characteristics unrelated to the questions.",
        "Apply selection criteria that are connected to the research questions.",
        "research_participant_selection",
    ),
    (
        "Risk/benefit balance",
        "Risks, discomforts or hazards and potential benefits, including participant benefits, must be clear and balanced.",
        "State both sides and evaluate the balance before proceeding with the study activity.",
        "research_risk_benefit",
    ),
    (
        "Anonymity",
        "The approved guidance requires anonymity for all participants at all times.",
        "Protect participant anonymity throughout the research activity.",
        "research_anonymity",
    ),
    (
        "Field/community application",
        "The University policy and Code apply to interactions with partner-school staff, community organizations and their clients.",
        "Apply the integrity framework during field, clinical, internship and community-facing work.",
        "off_campus_application",
    ),
]


def clone_run_with_text(run_element, text: str):
    cloned = deepcopy(run_element)
    for child in list(cloned):
        if child.tag != qn("w:rPr"):
            cloned.remove(child)
    node = OxmlElement("w:t")
    if text[:1].isspace() or text[-1:].isspace():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
    cloned.append(node)
    return cloned


def add_bookmark_around_paragraph(paragraph_element, name: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    ppr = paragraph_element.find(qn("w:pPr"))
    insert_at = 1 if ppr is not None else 0
    paragraph_element.insert(insert_at, start)
    paragraph_element.append(end)


def annotate_existing_phrase(
    document,
    phrase: str,
    comment_id: int,
    bookmark_name: str | None,
    bookmark_id: int | None,
) -> None:
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if phrase not in run.text:
                continue
            before, _, after = run.text.partition(phrase)
            parent = run._r.getparent()
            index = parent.index(run._r)
            elements = []
            if before:
                elements.append(clone_run_with_text(run._r, before))
            if bookmark_name is not None and bookmark_id is not None:
                bookmark_start = OxmlElement("w:bookmarkStart")
                bookmark_start.set(qn("w:id"), str(bookmark_id))
                bookmark_start.set(qn("w:name"), bookmark_name)
                elements.append(bookmark_start)
            comment_start = OxmlElement("w:commentRangeStart")
            comment_start.set(qn("w:id"), str(comment_id))
            elements.append(comment_start)
            elements.append(clone_run_with_text(run._r, phrase))
            comment_end = OxmlElement("w:commentRangeEnd")
            comment_end.set(qn("w:id"), str(comment_id))
            elements.append(comment_end)
            reference_run = OxmlElement("w:r")
            reference_props = OxmlElement("w:rPr")
            reference_style = OxmlElement("w:rStyle")
            reference_style.set(qn("w:val"), "CommentReference")
            reference_props.append(reference_style)
            reference_run.append(reference_props)
            reference = OxmlElement("w:commentReference")
            reference.set(qn("w:id"), str(comment_id))
            reference_run.append(reference)
            elements.append(reference_run)
            if bookmark_name is not None and bookmark_id is not None:
                bookmark_end = OxmlElement("w:bookmarkEnd")
                bookmark_end.set(qn("w:id"), str(bookmark_id))
                elements.append(bookmark_end)
            if after:
                elements.append(clone_run_with_text(run._r, after))
            parent.remove(run._r)
            for offset, element in enumerate(elements):
                parent.insert(index + offset, element)
            return
    raise ValueError(f"保护模板中未找到批注锚点：{phrase}")


def add_comments_part(document) -> None:
    root = OxmlElement("w:comments")
    for comment_id, (_, text, _) in enumerate(COMMENT_SPECS):
        comment = OxmlElement("w:comment")
        comment.set(qn("w:id"), str(comment_id))
        comment.set(qn("w:author"), "Course Policy Editor")
        comment.set(qn("w:initials"), "CPE")
        comment.set(qn("w:date"), CHANGE_DATE)
        paragraph = OxmlElement("w:p")
        run = OxmlElement("w:r")
        text_node = OxmlElement("w:t")
        text_node.text = text
        run.append(text_node)
        paragraph.append(run)
        comment.append(paragraph)
        root.append(comment)
    part = Part(
        PackURI("/word/comments.xml"),
        COMMENTS_CONTENT_TYPE,
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"),
        document.part.package,
    )
    document.part.relate_to(part, COMMENTS_REL_TYPE)


def enable_track_changes(document) -> None:
    settings = document.settings.element
    existing = settings.find(qn("w:trackRevisions"))
    if existing is None:
        settings.insert(0, OxmlElement("w:trackRevisions"))


def clone_as_tracked_insertion(source_paragraph, change_id: int, bookmark_name: str | None, bookmark_id: int) -> object:
    cloned = deepcopy(source_paragraph._p)
    inserted = OxmlElement("w:ins")
    inserted.set(qn("w:id"), str(change_id))
    inserted.set(qn("w:author"), "Course Policy Editor")
    inserted.set(qn("w:date"), CHANGE_DATE)
    for child in list(cloned):
        if child.tag == qn("w:pPr"):
            continue
        cloned.remove(child)
        inserted.append(child)
    cloned.append(inserted)
    if bookmark_name:
        add_bookmark_around_paragraph(cloned, bookmark_name, bookmark_id)
    return cloned


def add_internal_link(paragraph, bookmark_name: str, label: str = "Go to source") -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), bookmark_name)
    hyperlink.set(qn("w:history"), "1")
    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    props.extend([color, underline])
    text = OxmlElement("w:t")
    text.text = label
    run.extend([props, text])
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def format_matrix(table) -> None:
    table.autofit = False
    widths = [1.25, 2.35, 2.35, 1.05]
    properties = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:color"), "7F8C8D")
        borders.append(border)
    properties.append(borders)
    for row_index, row in enumerate(table.rows):
        row_props = row._tr.get_or_add_trPr()
        if row_index == 0:
            repeat = OxmlElement("w:tblHeader")
            repeat.set(qn("w:val"), "true")
            row_props.append(repeat)
        else:
            row_props.append(OxmlElement("w:cantSplit"))
        for column_index, cell in enumerate(row.cells):
            cell.width = Inches(widths[column_index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(1)
                for run in paragraph.runs:
                    run.font.size = Pt(8.5)
            if row_index == 0:
                shade = OxmlElement("w:shd")
                shade.set(qn("w:fill"), "D9EAF7")
                cell._tc.get_or_add_tcPr().append(shade)
                for run in cell.paragraphs[0].runs:
                    run.bold = True


def build(input_dir: Path, output_dir: Path) -> Path:
    master_path = input_dir / "sample_syllabus.docx"
    policy_path = input_dir / "academic_integrity.docx"
    document = Document(master_path)
    supplement = Document(policy_path)
    enable_track_changes(document)

    turnitin_anchor = next(
        paragraph._p
        for paragraph in document.paragraphs
        if "verified using Turnitin technology" in paragraph.text
    )
    cursor = turnitin_anchor
    nonblank_source_indexes = [0, 1, 2, 4, 5, 6, 7]
    for change_id, source_index in enumerate(nonblank_source_indexes, start=100):
        tracked = clone_as_tracked_insertion(
            supplement.paragraphs[source_index],
            change_id,
            TRACKED_BOOKMARKS.get(source_index),
            200 + source_index,
        )
        cursor.addnext(tracked)
        cursor = tracked

    next_bookmark_id = 300
    for comment_id, (phrase, _, bookmark_name) in enumerate(COMMENT_SPECS):
        annotate_existing_phrase(
            document,
            phrase,
            comment_id,
            bookmark_name,
            next_bookmark_id if bookmark_name else None,
        )
        if bookmark_name:
            next_bookmark_id += 1
    add_comments_part(document)

    review_section = document.add_section(WD_SECTION.NEW_PAGE)
    review_section.header.is_linked_to_previous = False
    review_section.header.paragraphs[0].text = "College of Education — Policy Integration Review"

    document.add_heading("Policy Integration Review", level=1)
    document.add_heading("Scope Decision", level=2)
    document.add_paragraph(SCOPE_DECISION)
    document.add_heading("Research Activity Control Matrix", level=2)
    matrix = document.add_table(rows=1, cols=4)
    for cell, label in zip(matrix.rows[0].cells, ["Control", "Evidence-based scope", "Operational consequence", "Source link"]):
        cell.text = label
    for control, scope, consequence, bookmark in MATRIX_ROWS:
        cells = matrix.add_row().cells
        cells[0].text = control
        cells[1].text = scope
        cells[2].text = consequence
        cells[3].text = ""
        add_internal_link(cells[3].paragraphs[0], bookmark)
    format_matrix(matrix)

    document.add_heading("Editorial Status", level=2)
    document.add_paragraph(
        "The research-policy paragraphs are tracked insertions and the three scope notes are anchored comments. "
        "The protected master remains a faculty template with course-specific prompts; this review does not mark it publication-ready."
    )

    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / OUTPUT_NAME
    document.save(target)
    return target


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    build(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
