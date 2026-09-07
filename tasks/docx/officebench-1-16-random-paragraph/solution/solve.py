#!/usr/bin/env python3
"""制作本题专用的咖啡馆文案审校工作稿。"""

from __future__ import annotations

import argparse
import re
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.shared import Pt
from lxml import etree


OUTPUT_NAME = "cafe_profile_editorial_review.docx"
CHANGE_DATE = datetime(2026, 9, 7, 15, 0, tzinfo=timezone.utc).isoformat()
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
WORD_PATTERN = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ0-9]+(?:[’'-][A-Za-zÀ-ÖØ-öø-ÿ0-9]+)*")


EVIDENCE_NOTE = (
    "The five approved revisions remain bounded by the controlled source. S1 keeps the city setting and the café as a "
    "tranquil pause. S2 retains the vintage furniture, soft lighting, cozy interior, and both locals and tourists. S3 keeps "
    "only the named coffee and pastry aromas. S4 preserves the mismatched tables, patron conversation or reflection, and "
    "drinks. S5 retains the outside-inside contrast, slowed-time image, and respite. The register shows the approved copy "
    "reduces the source from 103 words to 68, a net change of -35 words. This tightening removes ornament and compresses "
    "clauses without adding proper nouns, dates, numbers, prices, menu items, addresses, ratings, hours, or measurable "
    "service promises."
)


def split_sentences(text: str) -> list[str]:
    return [item.strip() for item in re.split(r"(?<=[.!?])\s+", text.strip()) if item.strip()]


def words(text: str) -> list[str]:
    return WORD_PATTERN.findall(text)


def ticket_rows(ticket: Document) -> list[dict[str, str]]:
    table = ticket.tables[0]
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    return [dict(zip(headers, [cell.text.strip() for cell in row.cells])) for row in table.rows[1:]]


def enable_track_changes(document: Document) -> None:
    settings = document.settings.element
    if settings.find(qn("w:trackRevisions")) is None:
        settings.insert(0, OxmlElement("w:trackRevisions"))


def text_run(text: str, deleted: bool = False):
    run = OxmlElement("w:r")
    node = OxmlElement("w:delText" if deleted else "w:t")
    if text[:1].isspace() or text[-1:].isspace():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
    run.append(node)
    return run


def tracked_replacement(paragraph, source: str, row: dict[str, str], index: int) -> None:
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)

    label = text_run(f"{row['ID']} | ")
    props = OxmlElement("w:rPr")
    props.append(OxmlElement("w:b"))
    label.insert(0, props)
    paragraph._p.append(label)

    deleted = OxmlElement("w:del")
    deleted.set(qn("w:id"), str(100 + index))
    deleted.set(qn("w:author"), "Copy Desk")
    deleted.set(qn("w:date"), CHANGE_DATE)
    deleted.append(text_run(source, deleted=True))
    paragraph._p.append(deleted)

    inserted = OxmlElement("w:ins")
    inserted.set(qn("w:id"), str(200 + index))
    inserted.set(qn("w:author"), "Copy Desk")
    inserted.set(qn("w:date"), CHANGE_DATE)

    bookmark_start = OxmlElement("w:bookmarkStart")
    bookmark_start.set(qn("w:id"), str(300 + index))
    bookmark_start.set(qn("w:name"), f"review_{row['ID']}")
    inserted.append(bookmark_start)

    approved = row["Approved replacement"]
    anchor = row["Comment anchor"]
    if anchor:
        before, found, after = approved.partition(anchor)
        if not found:
            raise ValueError(f"批注锚点不在批准文案中：{anchor}")
        if before:
            inserted.append(text_run(before))
        start = OxmlElement("w:commentRangeStart")
        start.set(qn("w:id"), str(index))
        inserted.append(start)
        inserted.append(text_run(anchor))
        end = OxmlElement("w:commentRangeEnd")
        end.set(qn("w:id"), str(index))
        inserted.append(end)
        reference_run = OxmlElement("w:r")
        reference_props = OxmlElement("w:rPr")
        reference_style = OxmlElement("w:rStyle")
        reference_style.set(qn("w:val"), "CommentReference")
        reference_props.append(reference_style)
        reference_run.append(reference_props)
        reference = OxmlElement("w:commentReference")
        reference.set(qn("w:id"), str(index))
        reference_run.append(reference)
        inserted.append(reference_run)
        if after:
            inserted.append(text_run(after))
    else:
        inserted.append(text_run(approved))

    bookmark_end = OxmlElement("w:bookmarkEnd")
    bookmark_end.set(qn("w:id"), str(300 + index))
    inserted.append(bookmark_end)
    paragraph._p.append(inserted)


def add_comments_part(document: Document, rows: list[dict[str, str]]) -> None:
    root = OxmlElement("w:comments")
    for index, row in enumerate(rows):
        if not row["Required comment"]:
            continue
        comment = OxmlElement("w:comment")
        comment.set(qn("w:id"), str(index))
        comment.set(qn("w:author"), "Copy Desk")
        comment.set(qn("w:initials"), "CD")
        comment.set(qn("w:date"), CHANGE_DATE)
        paragraph = OxmlElement("w:p")
        paragraph.append(text_run(row["Required comment"]))
        comment.append(paragraph)
        root.append(comment)
    part = Part(
        PackURI("/word/comments.xml"),
        COMMENTS_CONTENT_TYPE,
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"),
        document.part.package,
    )
    document.part.relate_to(part, COMMENTS_REL_TYPE)


def add_internal_link(paragraph, target: str, label: str) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), target)
    hyperlink.set(qn("w:history"), "1")
    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    props.extend([color, underline])
    run.append(props)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def fill_register(document: Document, source_sentences: list[str], rows: list[dict[str, str]]) -> tuple[int, int]:
    register = document.tables[2]
    for table_row, source, ticket in zip(register.rows[1:], source_sentences, rows):
        source_words = words(source)
        approved_words = words(ticket["Approved replacement"])
        cue = " ".join(source_words[:6]) + "…"
        values = [ticket["ID"], cue, str(len(source_words)), str(len(approved_words)), f"{len(approved_words) - len(source_words):+d}"]
        for cell, value in zip(table_row.cells[:5], values):
            cell.text = value
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.size = Pt(8.5)
        final_cell = table_row.cells[5]
        final_cell.text = ticket["Edit class"] + " • "
        paragraph = final_cell.paragraphs[0]
        paragraph.paragraph_format.space_after = Pt(0)
        for run in paragraph.runs:
            run.font.size = Pt(8.5)
        add_internal_link(paragraph, f"review_{ticket['ID']}", f"Review {ticket['ID']}")
    return sum(len(words(value)) for value in source_sentences), sum(len(words(row["Approved replacement"])) for row in rows)


def fill_summary(document: Document, source_total: int, approved_total: int) -> None:
    values = [str(source_total), str(approved_total), f"{approved_total - source_total:+d}", "None"]
    for row, value in zip(document.tables[3].rows, values):
        row.cells[1].text = value


def build(input_dir: Path, output_dir: Path) -> Path:
    source = (input_dir / "random_paragraph.txt").read_text(encoding="utf-8").strip()
    source_sentences = split_sentences(source)
    if len(source_sentences) != 5:
        raise ValueError(f"预期 5 句源文，实际 {len(source_sentences)}")

    document = Document(input_dir / "cafe_profile_editorial_template.docx")
    rows = ticket_rows(Document(input_dir / "cafe_profile_approved_copy_ticket.docx"))
    if [row["ID"] for row in rows] != [f"S{index}" for index in range(1, 6)]:
        raise ValueError("批准票据缺少 S1–S5 连续记录")
    if document.tables[1].cell(0, 0).text.strip() != source:
        raise ValueError("模板受控源与权威 TXT 不一致")

    enable_track_changes(document)
    working = [paragraph for paragraph in document.paragraphs if paragraph.style.name == "Working Sentence"]
    if len(working) != 5:
        raise ValueError(f"模板应有 5 个工作文案段，实际 {len(working)}")
    for index, (paragraph, source_sentence, row) in enumerate(zip(working, source_sentences, rows)):
        tracked_replacement(paragraph, source_sentence, row, index)
    add_comments_part(document, rows)

    source_total, approved_total = fill_register(document, source_sentences, rows)
    fill_summary(document, source_total, approved_total)
    document.tables[0].cell(2, 1).text = "Approved copy applied"
    note = next(paragraph for paragraph in document.paragraphs if paragraph.text.startswith("Add a 90–130 word evidence note"))
    note.text = EVIDENCE_NOTE
    note.style = document.styles["Evidence Note"]
    document.core_properties.title = "Café Profile Editorial Review — Approved Copy"
    document.core_properties.subject = "Tracked, evidence-bound copy-edit review"

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
