#!/usr/bin/env python3
"""officeval_059 参考解：仅在目标单元格追加真实相对超链接，并原样随附 PPT。"""

from __future__ import annotations

import argparse
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
HYPERLINK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
PPT_NAME = "青梧谷灵MR共创计划.pptx"
DOC_OUTPUT_NAME = "青梧谷灵MR共创计划_项目申报书_已链接PPT.docx"
DISPLAY_TEXT = "项目介绍材料PPT：青梧谷灵MR共创计划.pptx"


def qname(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def next_relationship_id(root) -> str:
    values = []
    for relation in root.findall(qname(REL_NS, "Relationship")):
        match = re.fullmatch(r"rId(\d+)", relation.get("Id", ""))
        if match:
            values.append(int(match.group(1)))
    return f"rId{max(values, default=0) + 1}"


def add_link(document_xml: bytes, relationships_xml: bytes) -> tuple[bytes, bytes]:
    parser = etree.XMLParser(remove_blank_text=False)
    document = etree.fromstring(document_xml, parser)
    relationships = etree.fromstring(relationships_xml, parser)
    ns = {"w": W_NS, "r": R_NS}

    target_cell = None
    for row in document.xpath(".//w:tr", namespaces=ns):
        cells = row.xpath("./w:tc", namespaces=ns)
        if len(cells) < 2:
            continue
        label = "".join(cells[0].xpath(".//w:t/text()", namespaces=ns)).replace(" ", "")
        if "项目介绍材料" in label:
            target_cell = cells[1]
            break
    if target_cell is None:
        raise ValueError("未找到‘项目介绍材料’右侧单元格")
    if target_cell.xpath(".//w:hyperlink", namespaces=ns):
        raise ValueError("目标单元格已经存在超链接，拒绝重复写入")

    relationship_id = next_relationship_id(relationships)
    relation = etree.SubElement(relationships, qname(REL_NS, "Relationship"))
    relation.set("Id", relationship_id)
    relation.set("Type", HYPERLINK_TYPE)
    relation.set("Target", PPT_NAME)
    relation.set("TargetMode", "External")

    paragraphs = target_cell.xpath("./w:p", namespaces=ns)
    paragraph = paragraphs[0] if paragraphs else etree.SubElement(target_cell, qname(W_NS, "p"))
    paragraph_properties = paragraph.find(qname(W_NS, "pPr"))
    if paragraph_properties is None:
        paragraph_properties = etree.Element(qname(W_NS, "pPr"))
        paragraph.insert(0, paragraph_properties)
    justification = paragraph_properties.find(qname(W_NS, "jc"))
    if justification is None:
        justification = etree.SubElement(paragraph_properties, qname(W_NS, "jc"))
    justification.set(qname(W_NS, "val"), "center")

    hyperlink = etree.SubElement(paragraph, qname(W_NS, "hyperlink"))
    hyperlink.set(qname(R_NS, "id"), relationship_id)
    hyperlink.set(qname(W_NS, "history"), "1")
    run = etree.SubElement(hyperlink, qname(W_NS, "r"))
    run_properties = etree.SubElement(run, qname(W_NS, "rPr"))
    fonts = etree.SubElement(run_properties, qname(W_NS, "rFonts"))
    for key in ("ascii", "hAnsi", "eastAsia"):
        fonts.set(qname(W_NS, key), "Noto Sans CJK SC")
    color = etree.SubElement(run_properties, qname(W_NS, "color"))
    color.set(qname(W_NS, "val"), "0563C1")
    underline = etree.SubElement(run_properties, qname(W_NS, "u"))
    underline.set(qname(W_NS, "val"), "single")
    size = etree.SubElement(run_properties, qname(W_NS, "sz"))
    size.set(qname(W_NS, "val"), "18")
    size_cs = etree.SubElement(run_properties, qname(W_NS, "szCs"))
    size_cs.set(qname(W_NS, "val"), "18")
    text = etree.SubElement(run, qname(W_NS, "t"))
    text.text = DISPLAY_TEXT

    return (
        etree.tostring(document, xml_declaration=True, encoding="UTF-8", standalone=True),
        etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone=True),
    )


def rewrite_docx(source: Path, destination: Path) -> None:
    with zipfile.ZipFile(source, "r") as archive:
        document_xml, relationships_xml = add_link(
            archive.read("word/document.xml"),
            archive.read("word/_rels/document.xml.rels"),
        )
        destination.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(prefix="officeval059-", suffix=".docx", dir=destination.parent, delete=False) as handle:
            temporary = Path(handle.name)
        try:
            with zipfile.ZipFile(temporary, "w") as output:
                for member in archive.infolist():
                    payload = archive.read(member.filename)
                    if member.filename == "word/document.xml":
                        payload = document_xml
                    elif member.filename == "word/_rels/document.xml.rels":
                        payload = relationships_xml
                    output.writestr(member, payload)
            temporary.replace(destination)
        finally:
            if temporary.exists():
                temporary.unlink()


def build(input_dir: Path, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    rewrite_docx(input_dir / "青梧谷灵MR共创计划_项目申报书.docx", output_dir / DOC_OUTPUT_NAME)
    shutil.copy2(input_dir / PPT_NAME, output_dir / PPT_NAME)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    build(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
