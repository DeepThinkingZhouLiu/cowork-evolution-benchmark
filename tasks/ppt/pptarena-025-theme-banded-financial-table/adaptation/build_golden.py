#!/usr/bin/env python3
"""PPTArena Case 25 任务专属 Golden：只修改原生表格的填充和字色。"""

from __future__ import annotations

import argparse
import hashlib
import zipfile
from pathlib import Path

from lxml import etree


INPUT_SHA256 = "d0b3c9d5b9988a375252315d3300b3237f25684af23ed77a4330aa8021d10c30"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def remove_fill(parent) -> None:
    fill_tags = {
        qn(A_NS, "noFill"),
        qn(A_NS, "solidFill"),
        qn(A_NS, "gradFill"),
        qn(A_NS, "blipFill"),
        qn(A_NS, "pattFill"),
        qn(A_NS, "grpFill"),
    }
    for child in list(parent):
        if child.tag in fill_tags:
            parent.remove(child)


def add_scheme_fill(parent, scheme: str, transforms: tuple[tuple[str, str], ...] = ()) -> None:
    remove_fill(parent)
    fill = etree.Element(qn(A_NS, "solidFill"))
    color = etree.SubElement(fill, qn(A_NS, "schemeClr"), val=scheme)
    for name, value in transforms:
        etree.SubElement(color, qn(A_NS, name), val=value)
    parent.insert(0, fill)


def style_table(slide_xml: bytes) -> bytes:
    root = etree.fromstring(slide_xml)
    tables = root.xpath("./p:cSld/p:spTree/p:graphicFrame/a:graphic/a:graphicData/a:tbl", namespaces=NS)
    if len(tables) != 1:
        raise ValueError(f"应有且仅有 1 个表格，实际为 {len(tables)}")
    rows = tables[0].xpath("./a:tr", namespaces=NS)
    if len(rows) != 5 or any(len(row.xpath("./a:tc", namespaces=NS)) != 4 for row in rows):
        raise ValueError("冻结表格不再是 5×4")

    expected_text = [
        ["Quarter", "Revenue ($)", "Expenses ($)", "Net profit ($)"],
        ["Q1", "$200,000", "$150,000", "$50,000"],
        ["Q2", "$300,000", "$200,000", "$100,000"],
        ["Q3", "$400,000", "$250,000", "$150,000"],
        ["Q4", "$300,000", "$200,000", "$100,000"],
    ]

    for row_index, row in enumerate(rows):
        if row_index == 0:
            fill = ("accent6", (("lumMod", "50000"),))
        elif row_index % 2 == 1:
            fill = ("tx2", (("lumMod", "75000"), ("lumOff", "25000")))
        else:
            fill = ("tx2", (("lumMod", "50000"), ("lumOff", "50000")))
        for column_index, cell in enumerate(row.xpath("./a:tc", namespaces=NS)):
            text = "".join(cell.xpath(".//a:t/text()", namespaces=NS))
            if text != expected_text[row_index][column_index]:
                raise ValueError(f"单元格 {row_index},{column_index} 文字不符：{text!r}")
            properties = cell.find("a:tcPr", NS)
            if properties is None:
                properties = etree.SubElement(cell, qn(A_NS, "tcPr"))
            add_scheme_fill(properties, fill[0], fill[1])

            run_properties = cell.xpath("./a:txBody/a:p/a:r/a:rPr", namespaces=NS)
            end_properties = cell.xpath("./a:txBody/a:p/a:endParaRPr", namespaces=NS)
            if not run_properties:
                raise ValueError(f"单元格 {row_index},{column_index} 无文字运行属性")
            for properties in run_properties + end_properties:
                add_scheme_fill(properties, "bg1")

    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def build(source: Path, output: Path) -> None:
    if not source.is_file() or sha256(source) != INPUT_SHA256:
        raise ValueError("冻结输入缺失或 SHA256 不匹配")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as archive_in, zipfile.ZipFile(output, "w") as archive_out:
        for info in archive_in.infolist():
            data = archive_in.read(info.filename)
            if info.filename == "ppt/slides/slide1.xml":
                data = style_table(data)
            archive_out.writestr(info, data)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    build(args.source, args.output)
    print(f"{args.output}\t{args.output.stat().st_size}\t{sha256(args.output)}")


if __name__ == "__main__":
    main()
