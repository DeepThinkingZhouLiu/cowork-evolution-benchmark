#!/usr/bin/env python3
"""LLM-DOCX-Editing F6 的逐题参考解法：在 OOXML 层接受全部修订。"""

from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


INPUT_NAME = "contract-with-redlines.docx"
OUTPUT_NAME = "contract-accepted.docx"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

UNWRAP_ON_ACCEPT = ("ins", "moveTo")
REMOVE_ON_ACCEPT = ("del", "moveFrom")
REMOVE_MARKERS = {
    "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
    "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd", "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
    "pPrChange", "rPrChange", "tblPrChange", "tblGridChange", "trPrChange", "tcPrChange", "sectPrChange", "numberingChange",
}


def local_name(element) -> str:
    return etree.QName(element).localname


def remove_elements(root, names: tuple[str, ...] | set[str]) -> None:
    targets = [element for element in root.iter() if local_name(element) in names]
    for element in targets:
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)


def unwrap_elements(root, names: tuple[str, ...]) -> None:
    targets = [element for element in root.iter() if local_name(element) in names]
    for element in targets:
        parent = element.getparent()
        if parent is None:
            continue
        index = parent.index(element)
        for child in list(element):
            element.remove(child)
            parent.insert(index, child)
            index += 1
        parent.remove(element)


def accept_document_xml(data: bytes) -> bytes:
    root = etree.fromstring(data)
    insertions = len(root.xpath(".//w:ins", namespaces=NS))
    deletions = len(root.xpath(".//w:del", namespaces=NS))
    if insertions != 5 or deletions != 4:
        raise ValueError(f"F6 固定输入应有 5 个插入和 4 个删除，实际为 {insertions}/{deletions}")
    remove_elements(root, REMOVE_ON_ACCEPT)
    unwrap_elements(root, UNWRAP_ON_ACCEPT)
    remove_elements(root, REMOVE_MARKERS)
    remaining = [local_name(element) for element in root.iter() if local_name(element) in {*UNWRAP_ON_ACCEPT, *REMOVE_ON_ACCEPT, *REMOVE_MARKERS}]
    if remaining:
        raise ValueError(f"仍有修订节点：{remaining}")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def clean_settings_xml(data: bytes) -> bytes:
    root = etree.fromstring(data)
    remove_elements(root, {"trackRevisions"})
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def solve(input_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(input_path, "r") as source, ZipFile(output_path, "w", compression=ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "word/document.xml":
                data = accept_document_xml(data)
            elif info.filename == "word/settings.xml":
                data = clean_settings_xml(data)
            target.writestr(info, data)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir / INPUT_NAME, args.output_dir / OUTPUT_NAME)


if __name__ == "__main__":
    main()
