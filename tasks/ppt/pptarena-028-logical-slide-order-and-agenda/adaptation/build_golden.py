#!/usr/bin/env python3
"""PPTArena Case 28 任务专属 Golden：重排原页、议程与页码，不改写其他内容。"""

from __future__ import annotations

import argparse
import copy
import hashlib
import zipfile
from pathlib import Path

from lxml import etree


INPUT_SHA256 = "bcbdf13b305029f1fca7c33d0b6b0020f84acf963bf066471d4be00971c2175b"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A16_NS = "http://schemas.microsoft.com/office/drawing/2014/main"
NS = {"p": P_NS, "a": A_NS, "r": R_NS, "a16": A16_NS}

DESIRED_RELATIONSHIP_ORDER = ["rId3", "rId2", "rId5", "rId7", "rId4", "rId6"]
DESIRED_AGENDA = [
    "Visual appeal",
    "Engaging the audience",
    "Effective delivery techniques",
    "Conclusion",
    "Q&A",
]
PAGE_NUMBER_BY_PART = {
    "ppt/slides/slide2.xml": "1",
    "ppt/slides/slide1.xml": "2",
    "ppt/slides/slide4.xml": "3",
    "ppt/slides/slide6.xml": "4",
    "ppt/slides/slide3.xml": "5",
    "ppt/slides/slide5.xml": "6",
}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def serialize(root) -> bytes:
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def rewrite_presentation(data: bytes) -> bytes:
    root = etree.fromstring(data)
    slide_list = root.find("p:sldIdLst", NS)
    if slide_list is None or len(slide_list) != 6:
        raise ValueError("未找到预期的 6 个幻灯片关系")
    by_relationship = {node.get(f"{{{R_NS}}}id"): node for node in slide_list}
    if set(by_relationship) != set(DESIRED_RELATIONSHIP_ORDER):
        raise ValueError("幻灯片关系集与冻结输入不符")
    for node in list(slide_list):
        slide_list.remove(node)
    for relationship_id in DESIRED_RELATIONSHIP_ORDER:
        slide_list.append(by_relationship[relationship_id])
    return serialize(root)


def reorder_agenda(root) -> None:
    nodes = root.xpath(
        "./p:cSld/p:spTree/p:sp[p:nvSpPr/p:cNvPr/@id='4']/p:txBody",
        namespaces=NS,
    )
    if len(nodes) != 1:
        raise ValueError("未唯一定位议程文本框")
    body = nodes[0]
    paragraphs = body.findall("a:p", NS)
    by_text = {}
    for paragraph in paragraphs:
        text = "".join(paragraph.xpath(".//a:t/text()", namespaces=NS)).strip()
        if text:
            by_text[text] = paragraph
    if set(by_text) != set(DESIRED_AGENDA):
        raise ValueError(f"议程条目与冻结输入不符：{sorted(by_text)}")
    for paragraph in paragraphs:
        body.remove(paragraph)
    for text in DESIRED_AGENDA:
        body.append(by_text[text])


def slide_number_shapes(root):
    return root.xpath(
        "./p:cSld/p:spTree/p:sp[p:nvSpPr/p:nvPr/p:ph/@type='sldNum']",
        namespaces=NS,
    )


def set_slide_number(shape, number: str) -> None:
    fields = shape.xpath("./p:txBody/a:p/a:fld[@type='slidenum']/a:t", namespaces=NS)
    if len(fields) != 1:
        raise ValueError("页码占位符中的 slidenum 字段不唯一")
    fields[0].text = number


def add_missing_slide_number(root, source_number_shape) -> None:
    if slide_number_shapes(root):
        raise ValueError("第 4 个原始页面不应已有页码对象")
    shape = copy.deepcopy(source_number_shape)
    shape_id = shape.find("./p:nvSpPr/p:cNvPr", NS)
    shape_id.set("id", "9")
    shape_id.set("name", "Slide Number Placeholder 8")
    creation = shape.find("./p:nvSpPr/p:cNvPr/a:extLst/a:ext/a16:creationId", NS)
    if creation is not None:
        creation.set("id", "{6F2BFD4D-8B56-4A95-80AB-280000000003}")
    field = shape.find("./p:txBody/a:p/a:fld", NS)
    if field is None:
        raise ValueError("页码源对象缺少 slidenum 字段")
    field.set("id", "{6F2BFD4D-8B56-4A95-80AB-280000000013}")
    set_slide_number(shape, "3")
    tree = root.find("./p:cSld/p:spTree", NS)
    extension = tree.find("p:extLst", NS)
    if extension is None:
        tree.append(shape)
    else:
        tree.insert(tree.index(extension), shape)


def rewrite_slides(contents: dict[str, bytes]) -> dict[str, bytes]:
    roots = {
        name: etree.fromstring(contents[name])
        for name in PAGE_NUMBER_BY_PART
    }
    reorder_agenda(roots["ppt/slides/slide2.xml"])

    number_shape = slide_number_shapes(roots["ppt/slides/slide3.xml"])
    if len(number_shape) != 1:
        raise ValueError("结论页的页码对象不唯一")
    add_missing_slide_number(roots["ppt/slides/slide4.xml"], number_shape[0])

    for name, number in PAGE_NUMBER_BY_PART.items():
        shapes = slide_number_shapes(roots[name])
        if len(shapes) != 1:
            raise ValueError(f"{name} 的页码对象不唯一")
        set_slide_number(shapes[0], number)
        contents[name] = serialize(roots[name])
    return contents


def build(source: Path, output: Path) -> None:
    if not source.is_file() or sha256(source) != INPUT_SHA256:
        raise ValueError("冻结输入缺失或 SHA256 不匹配")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as archive_in:
        contents = {info.filename: archive_in.read(info.filename) for info in archive_in.infolist()}
        contents["ppt/presentation.xml"] = rewrite_presentation(contents["ppt/presentation.xml"])
        contents = rewrite_slides(contents)
        with zipfile.ZipFile(output, "w") as archive_out:
            for info in archive_in.infolist():
                archive_out.writestr(info, contents[info.filename])


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    build(args.source, args.output)
    print(f"{args.output}\t{args.output.stat().st_size}\t{sha256(args.output)}")


if __name__ == "__main__":
    main()
