#!/usr/bin/env python3
"""officeval_078 参考解：只改第 8 页 XML/关系表，加入形状级 Excel 相对超链接。"""

from __future__ import annotations

import argparse
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
HYPERLINK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
XLSX_NAME = "CRM_pipeline_restyled_reordered.xlsx"
PPT_OUTPUT_NAME = "区域协同汇报_第8页Excel超链接版.pptx"


def qname(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def next_relationship_id(root) -> str:
    numbers = []
    for relation in root.findall(qname(REL_NS, "Relationship")):
        match = re.fullmatch(r"rId(\d+)", relation.get("Id", ""))
        if match:
            numbers.append(int(match.group(1)))
    return f"rId{max(numbers, default=0) + 1}"


def add_button(slide_xml: bytes, relationships_xml: bytes) -> tuple[bytes, bytes]:
    parser = etree.XMLParser(remove_blank_text=False)
    slide = etree.fromstring(slide_xml, parser)
    relationships = etree.fromstring(relationships_xml, parser)
    ns = {"p": P_NS, "a": A_NS, "r": R_NS}
    tree = slide.find(".//p:cSld/p:spTree", namespaces=ns)
    if tree is None:
        raise ValueError("第 8 页缺少形状树")
    if slide.xpath(".//a:hlinkClick", namespaces=ns):
        raise ValueError("第 8 页已存在超链接，拒绝重复写入")

    existing_ids = [int(value) for value in slide.xpath(".//p:cNvPr/@id", namespaces=ns) if value.isdigit()]
    shape_id = max(existing_ids, default=1) + 1
    relationship_id = next_relationship_id(relationships)
    relation = etree.SubElement(relationships, qname(REL_NS, "Relationship"))
    relation.set("Id", relationship_id)
    relation.set("Type", HYPERLINK_TYPE)
    relation.set("Target", XLSX_NAME)
    relation.set("TargetMode", "External")

    shape = etree.Element(qname(P_NS, "sp"))
    non_visual = etree.SubElement(shape, qname(P_NS, "nvSpPr"))
    properties = etree.SubElement(non_visual, qname(P_NS, "cNvPr"))
    properties.set("id", str(shape_id))
    properties.set("name", "打开 Excel 工作簿：CRM_pipeline_restyled_reordered.xlsx")
    hyperlink = etree.SubElement(properties, qname(A_NS, "hlinkClick"))
    hyperlink.set(qname(R_NS, "id"), relationship_id)
    etree.SubElement(non_visual, qname(P_NS, "cNvSpPr"))
    etree.SubElement(non_visual, qname(P_NS, "nvPr"))

    shape_properties = etree.SubElement(shape, qname(P_NS, "spPr"))
    transform = etree.SubElement(shape_properties, qname(A_NS, "xfrm"))
    offset = etree.SubElement(transform, qname(A_NS, "off"))
    offset.set("x", str(int(27.0 * 360000)))
    offset.set("y", str(int(7.4 * 360000)))
    extent = etree.SubElement(transform, qname(A_NS, "ext"))
    extent.set("cx", str(int(4.6 * 360000)))
    extent.set("cy", str(int(2.4 * 360000)))
    geometry = etree.SubElement(shape_properties, qname(A_NS, "prstGeom"))
    geometry.set("prst", "roundRect")
    etree.SubElement(geometry, qname(A_NS, "avLst"))
    fill = etree.SubElement(shape_properties, qname(A_NS, "solidFill"))
    etree.SubElement(fill, qname(A_NS, "srgbClr")).set("val", "107C41")
    line = etree.SubElement(shape_properties, qname(A_NS, "ln"))
    line.set("w", "19050")
    line_fill = etree.SubElement(line, qname(A_NS, "solidFill"))
    etree.SubElement(line_fill, qname(A_NS, "srgbClr")).set("val", "33C481")

    text_body = etree.SubElement(shape, qname(P_NS, "txBody"))
    body_properties = etree.SubElement(text_body, qname(A_NS, "bodyPr"))
    body_properties.set("anchor", "ctr")
    body_properties.set("lIns", "91440")
    body_properties.set("rIns", "91440")
    etree.SubElement(text_body, qname(A_NS, "lstStyle"))
    for text_value, size, bold in [("打开 Excel", "1500", "1"), ("CRM 工作簿", "1100", "0")]:
        paragraph = etree.SubElement(text_body, qname(A_NS, "p"))
        paragraph_properties = etree.SubElement(paragraph, qname(A_NS, "pPr"))
        paragraph_properties.set("algn", "ctr")
        run = etree.SubElement(paragraph, qname(A_NS, "r"))
        run_properties = etree.SubElement(run, qname(A_NS, "rPr"))
        run_properties.set("lang", "zh-CN")
        run_properties.set("sz", size)
        run_properties.set("b", bold)
        run_fill = etree.SubElement(run_properties, qname(A_NS, "solidFill"))
        etree.SubElement(run_fill, qname(A_NS, "srgbClr")).set("val", "FFFFFF")
        latin = etree.SubElement(run_properties, qname(A_NS, "latin"))
        latin.set("typeface", "Aptos")
        east_asian = etree.SubElement(run_properties, qname(A_NS, "ea"))
        east_asian.set("typeface", "Noto Sans CJK SC")
        etree.SubElement(run, qname(A_NS, "t")).text = text_value
        end_properties = etree.SubElement(paragraph, qname(A_NS, "endParaRPr"))
        end_properties.set("lang", "zh-CN")
        end_properties.set("sz", size)

    tree.append(shape)
    return (
        etree.tostring(slide, xml_declaration=True, encoding="UTF-8", standalone=True),
        etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone=True),
    )


def rewrite_pptx(source: Path, destination: Path) -> None:
    with zipfile.ZipFile(source, "r") as archive:
        slide_xml, relationships_xml = add_button(
            archive.read("ppt/slides/slide8.xml"),
            archive.read("ppt/slides/_rels/slide8.xml.rels"),
        )
        destination.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(prefix="officeval078-", suffix=".pptx", dir=destination.parent, delete=False) as handle:
            temporary = Path(handle.name)
        try:
            with zipfile.ZipFile(temporary, "w") as output:
                for member in archive.infolist():
                    payload = archive.read(member.filename)
                    if member.filename == "ppt/slides/slide8.xml":
                        payload = slide_xml
                    elif member.filename == "ppt/slides/_rels/slide8.xml.rels":
                        payload = relationships_xml
                    output.writestr(member, payload)
            temporary.replace(destination)
        finally:
            if temporary.exists():
                temporary.unlink()


def build(input_dir: Path, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    rewrite_pptx(input_dir / "区域协同汇报.pptx", output_dir / PPT_OUTPUT_NAME)
    shutil.copy2(input_dir / XLSX_NAME, output_dir / XLSX_NAME)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    build(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
