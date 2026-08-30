#!/usr/bin/env python3
"""PPTArena Case 59：从冻结输入只写入两个 URL 的外部超链接。"""

from __future__ import annotations

import argparse
import copy
import hashlib
import zipfile
from pathlib import Path

from lxml import etree


INPUT_SHA256 = "33132900f1422268aec0a20ba9e6ba0c2be1e6b9a4baeb9f0887c371506792f5"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
HYPERLINK_REL_TYPE = f"{R_NS}/hyperlink"
FULL_SENTENCE = "Visit https://www.berkeley.edu/ and https://illinois.edu/ for more details"
PARTS = [
    ("Visit ", None),
    ("https://www.berkeley.edu/", "rId6"),
    (" and ", None),
    ("https://illinois.edu/", "rId7"),
    (" for more details", None),
]
TARGETS = {
    "rId6": "https://www.berkeley.edu/",
    "rId7": "https://illinois.edu/",
}


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def serialize(root) -> bytes:
    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )


def rewrite_sentence(slide_xml: bytes) -> bytes:
    root = etree.fromstring(slide_xml)
    candidates = []
    for shape in root.xpath(".//p:sp", namespaces={"p": P_NS}):
        visible = "".join(shape.xpath(".//a:t/text()", namespaces={"a": A_NS}))
        if visible == FULL_SENTENCE:
            candidates.append(shape)
    if len(candidates) != 1:
        raise ValueError(f"目标文本框数量应为 1，实际为 {len(candidates)}")
    paragraph = candidates[0].xpath("./p:txBody/a:p", namespaces={"p": P_NS, "a": A_NS})
    if len(paragraph) != 1:
        raise ValueError("目标文本框必须只有一个段落")
    runs = paragraph[0].xpath("./a:r", namespaces={"a": A_NS})
    if len(runs) != 1 or "".join(runs[0].xpath("./a:t/text()", namespaces={"a": A_NS})) != FULL_SENTENCE:
        raise ValueError("冻结输入的目标句运行结构与已审计版本不一致")
    original_run = runs[0]
    original_properties = original_run.find(qn(A_NS, "rPr"))
    insertion_index = paragraph[0].index(original_run)
    paragraph[0].remove(original_run)
    for offset, (text, relationship_id) in enumerate(PARTS):
        run = etree.Element(qn(A_NS, "r"))
        properties = (
            copy.deepcopy(original_properties)
            if original_properties is not None
            else etree.Element(qn(A_NS, "rPr"))
        )
        if relationship_id is not None:
            etree.SubElement(
                properties,
                qn(A_NS, "hlinkClick"),
                {qn(R_NS, "id"): relationship_id},
            )
        run.append(properties)
        text_node = etree.SubElement(run, qn(A_NS, "t"))
        if text.startswith(" ") or text.endswith(" "):
            text_node.set(qn("http://www.w3.org/XML/1998/namespace", "space"), "preserve")
        text_node.text = text
        paragraph[0].insert(insertion_index + offset, run)
    return serialize(root)


def rewrite_relationships(rels_xml: bytes) -> bytes:
    root = etree.fromstring(rels_xml)
    existing = {node.get("Id"): node for node in root}
    for relationship_id, target in TARGETS.items():
        if relationship_id in existing:
            raise ValueError(f"预留关系 ID 已存在：{relationship_id}")
        etree.SubElement(
            root,
            qn(PKG_REL_NS, "Relationship"),
            Id=relationship_id,
            Type=HYPERLINK_REL_TYPE,
            Target=target,
            TargetMode="External",
        )
    return serialize(root)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if sha256(args.input) != INPUT_SHA256:
        raise ValueError("冻结输入 SHA256 不匹配")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    changed = set()
    with zipfile.ZipFile(args.input) as archive_in, zipfile.ZipFile(args.output, "w") as archive_out:
        for info in archive_in.infolist():
            data = archive_in.read(info.filename)
            if info.filename == "ppt/slides/slide1.xml":
                data = rewrite_sentence(data)
                changed.add(info.filename)
            elif info.filename == "ppt/slides/_rels/slide1.xml.rels":
                data = rewrite_relationships(data)
                changed.add(info.filename)
            archive_out.writestr(info, data)
    expected = {"ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels"}
    if changed != expected:
        raise ValueError(f"未完成两个必需部件的改写：{changed}")
    print(f"{args.output} {sha256(args.output)}")


if __name__ == "__main__":
    main()
