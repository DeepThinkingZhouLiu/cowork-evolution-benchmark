#!/usr/bin/env python3
"""为 The House that Creaked 制作当题专用的选集编辑工作稿。"""

from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.shared import Inches, Pt, RGBColor
from lxml import etree


OUTPUT_NAME = "house_creak_editorial_review.docx"
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"


COMMENT_SPECS = {
    "electrical_sparking": (
        "power point which didn't spark whenever an appliance requiring over 1000 watts was plugged into it.",
        "This electrical warning functions as concrete foreshadowing: the house is already unsafe before the visible collapse.",
    ),
    "bed_ceiling_collapse": (
        "a double bed crash through the dining room ceiling.",
        "The structural failure becomes explicit at this point and converts Ros's earlier alarm into an immediate rescue cue.",
    ),
    "condemned_disclosure": (
        "That house was condemned years ago!",
        "This late disclosure confirms that the central risk was real and preventable, reframing the family's earlier warning signs.",
    ),
}

BOOKMARK_SPECS = {
    "move_victorian_gothic": "Victorian Gothic",
    "creaking_bedrooms": "The sound was worst in the master bedroom and children's rooms next to it",
    **{name: phrase for name, (phrase, _) in COMMENT_SPECS.items()},
    "gas_explosion": "there was a huge 'boom!' from the kitchen",
}

SYNOPSIS = (
    "Ros, a watchful family dog, reluctantly moves with Mother, Father, Peter and Amanda from a familiar suburb to a remote "
    "Victorian Gothic house. While Father treats the move as a status decision, Ros responds to the building with persistent "
    "unease. The family's isolation and the house's physical warnings accumulate: floors creak most heavily near the bedrooms, "
    "and a power point sparks under a load above 1000 watts. During a makeshift lesson in the dining room, Ros barks urgently and "
    "draws Mother and the children into the hall. A double bed then crashes through the dining-room ceiling. As they flee outside, "
    "the mansion collapses; a fractured gas main and the cake left in the oven trigger an explosion and fire. After the family "
    "walks three miles for help, a neighbour reveals that the house had been condemned years earlier. The Women's Institute "
    "provides practical aid, and the family temporarily returns to the children's grandmother while Mother rejects Father's "
    "unchecked decisions."
)

CHRONOLOGY = [
    (
        "1",
        "The house at the end of the overgrown drive was Victorian Gothic",
        "Remote, forbidding property accepted without a documented provenance check.",
        "The family enters an unsafe and isolated setting.",
        "move_victorian_gothic",
    ),
    (
        "2",
        "The sound was worst in the master bedroom and children's rooms next to it",
        "Repeated building movement is concentrated near occupied rooms.",
        "Ros and the family register persistent structural warning signs.",
        "creaking_bedrooms",
    ),
    (
        "3",
        "power point which didn't spark whenever an appliance requiring over 1000 watts was plugged into it",
        "Electrical installation sparks under an ordinary high load.",
        "A concrete utility hazard foreshadows the later catastrophe.",
        "electrical_sparking",
    ),
    (
        "4",
        "a double bed crash through the dining room ceiling",
        "The upper floor and ceiling suffer sudden structural failure.",
        "Ros's warning gets Mother, Peter and Amanda out of the room in time.",
        "bed_ceiling_collapse",
    ),
    (
        "5",
        "there was a huge 'boom!' from the kitchen",
        "The collapse fractures the gas main while a cake remains in the gas oven.",
        "The remaining walls are blown apart and the rubble catches fire.",
        "gas_explosion",
    ),
    (
        "6",
        "That house was condemned years ago!",
        "The neighbour discloses a known pre-existing condition after the loss.",
        "The warning signs are retrospectively confirmed as preventable risk.",
        "condemned_disclosure",
    ),
]

CHARACTERS = [
    (
        "Ros",
        "Hides, listens to the heavy creaking and barks urgently before the ceiling gives way.",
        "Draws Mother, Peter and Amanda out of the dining room before impact.",
        "Her embodied attention supplies the warning the humans repeatedly discount.",
    ),
    (
        "Mother",
        "Privately checks the property online, notices practical hazards, then orders everyone outside during the collapse.",
        "Protects the children and later rejects further unchecked decisions by Father.",
        "Her arc moves from constrained acquiescence to decisive agency.",
    ),
    (
        "Peter and Amanda",
        "Reluctantly adapt to isolation, join the makeshift lesson and follow Mother outside.",
        "They escape and later return to school while living with their grandmother.",
        "Their dependence makes the adults' risk decisions materially consequential.",
    ),
    (
        "Father",
        "Chooses the house to meet his aspirations and dismisses the family's social and practical concerns.",
        "The uninsured loss compounds the consequences of his unilateral choice.",
        "He represents authority detached from local evidence and household risk.",
    ),
    (
        "Neighbour and Women's Institute",
        "The neighbour identifies the prior condemnation; the Women's Institute supplies clothes, meals and dog food.",
        "They connect the family to emergency help and immediate recovery support.",
        "Late local knowledge contrasts with the absent due diligence before purchase.",
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


def annotate_phrase(document, phrase: str, bookmark_name: str, bookmark_id: int, comment_id: int | None = None) -> None:
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
            bookmark_start = OxmlElement("w:bookmarkStart")
            bookmark_start.set(qn("w:id"), str(bookmark_id))
            bookmark_start.set(qn("w:name"), bookmark_name)
            elements.append(bookmark_start)
            if comment_id is not None:
                comment_start = OxmlElement("w:commentRangeStart")
                comment_start.set(qn("w:id"), str(comment_id))
                elements.append(comment_start)
            elements.append(clone_run_with_text(run._r, phrase))
            if comment_id is not None:
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
            bookmark_end = OxmlElement("w:bookmarkEnd")
            bookmark_end.set(qn("w:id"), str(bookmark_id))
            elements.append(bookmark_end)
            if after:
                elements.append(clone_run_with_text(run._r, after))
            parent.remove(run._r)
            for offset, element in enumerate(elements):
                parent.insert(index + offset, element)
            return
    raise ValueError(f"原文中未找到可锚定的单一 run 短语：{phrase}")


def add_comments_part(document) -> None:
    root = OxmlElement("w:comments")
    for comment_id, (bookmark_name, (_, text)) in enumerate(COMMENT_SPECS.items()):
        comment = OxmlElement("w:comment")
        comment.set(qn("w:id"), str(comment_id))
        comment.set(qn("w:author"), "Anthology Editor")
        comment.set(qn("w:initials"), "AE")
        comment.set(qn("w:date"), datetime(2026, 9, 7, tzinfo=timezone.utc).isoformat())
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


def add_toc(paragraph) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), 'TOC \\o "1-2" \\h \\z \\u')
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "Update field to refresh appendix contents"
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def add_internal_link(paragraph, bookmark_name: str, label: str = "Go to evidence") -> None:
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


def format_table(table, widths: list[float]) -> None:
    table_props = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:color"), "9EADBA")
        borders.append(border)
    table_props.append(borders)
    table.autofit = False
    for row_index, row in enumerate(table.rows):
        tr_props = row._tr.get_or_add_trPr()
        if row_index == 0:
            header = OxmlElement("w:tblHeader")
            header.set(qn("w:val"), "true")
            tr_props.append(header)
        else:
            tr_props.append(OxmlElement("w:cantSplit"))
        for column_index, cell in enumerate(row.cells):
            cell.width = Inches(widths[column_index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(2)
                for run in paragraph.runs:
                    run.font.name = "Liberation Sans"
                    run.font.size = Pt(8.5)
                    if row_index == 0:
                        run.bold = True
            if row_index == 0:
                shading = OxmlElement("w:shd")
                shading.set(qn("w:fill"), "D9EAF7")
                cell._tc.get_or_add_tcPr().append(shading)


def ensure_heading_styles(document) -> None:
    for name, size, level, color in (
        ("Heading 1", 18, 0, RGBColor(31, 78, 121)),
        ("Heading 2", 13, 1, RGBColor(47, 84, 116)),
    ):
        if name in [style.name for style in document.styles]:
            continue
        style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = "Liberation Sans"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = color
        outline = OxmlElement("w:outlineLvl")
        outline.set(qn("w:val"), str(level))
        style.element.get_or_add_pPr().append(outline)


def build(input_dir: Path, output_dir: Path) -> Path:
    source_path = input_dir / "House_creak.docx"
    document = Document(source_path)
    ensure_heading_styles(document)
    original_paragraphs = [paragraph.text for paragraph in document.paragraphs]
    if len(original_paragraphs) != 47:
        raise ValueError(f"预期 47 个原文段落，实际为 {len(original_paragraphs)}")

    for bookmark_id, (bookmark_name, phrase) in enumerate(BOOKMARK_SPECS.items(), start=100):
        comment_id = None
        if bookmark_name in COMMENT_SPECS:
            comment_id = list(COMMENT_SPECS).index(bookmark_name)
        annotate_phrase(document, phrase, bookmark_name, bookmark_id, comment_id)
    add_comments_part(document)

    appendix_section = document.add_section(WD_SECTION.NEW_PAGE)
    appendix_section.header.is_linked_to_previous = False
    header = appendix_section.header.paragraphs[0]
    header.text = "Anthology Editorial Review — The House that Creaked"
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in header.runs:
        run.font.name = "Liberation Serif"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(89, 89, 89)

    title = document.add_heading("Editorial Analysis Appendix", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    toc = document.add_paragraph()
    add_toc(toc)

    document.add_heading("Synopsis", level=2)
    synopsis = document.add_paragraph(SYNOPSIS)
    synopsis.paragraph_format.space_after = Pt(10)

    document.add_heading("Incident Chronology", level=2)
    chronology = document.add_table(rows=1, cols=5)
    headers = ["Seq", "Story evidence", "Risk signal", "Consequence", "Source link"]
    for cell, label in zip(chronology.rows[0].cells, headers):
        cell.text = label
    for sequence, evidence, risk, consequence, bookmark in CHRONOLOGY:
        cells = chronology.add_row().cells
        cells[0].text = sequence
        cells[1].text = evidence
        cells[2].text = risk
        cells[3].text = consequence
        cells[4].text = ""
        add_internal_link(cells[4].paragraphs[0], bookmark)
    format_table(chronology, [0.35, 1.65, 1.55, 1.65, 0.8])

    document.add_page_break()
    document.add_heading("Character Decision Matrix", level=2)
    characters = document.add_table(rows=1, cols=4)
    headers = ["Character", "Evidence-based action or belief", "Effect on events", "Editorial interpretation"]
    for cell, label in zip(characters.rows[0].cells, headers):
        cell.text = label
    for values in CHARACTERS:
        cells = characters.add_row().cells
        for cell, value in zip(cells, values):
            cell.text = value
    format_table(characters, [1.05, 1.85, 1.75, 1.55])

    document.add_heading("Editorial Note", level=2)
    note = document.add_paragraph(
        "The apparatus deliberately distinguishes what the characters know at each stage from what the reader learns later. "
        "The three anchored comments mark the escalation from warning, to failure, to confirmation; the source links allow a "
        "copy editor to test every chronology row against the protected story text."
    )
    note.paragraph_format.keep_with_next = True

    if [paragraph.text for paragraph in document.paragraphs[:47]] != original_paragraphs:
        raise AssertionError("批注与书签添加意外改动了原文")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / OUTPUT_NAME
    document.save(output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    build(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
