#!/usr/bin/env python3
"""仅为 PPTArena Case 16 从冻结输入重建本地 Golden。"""

from __future__ import annotations

import argparse
import hashlib
import zipfile
from pathlib import Path

from lxml import etree


INPUT_SHA256 = "c17f8b37992c462006bf62a151b3f346df2dade7253f97027f69d715ab7f657c"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
SLIDE_PART = "ppt/slides/slide2.xml"
EMU_PER_POINT = 12_700

# 三张图的 ID、名称与水平位置均由 Case 16 输入和 Query 逐项确认。
TARGETS = {
    "27": {"name": "Picture 26", "left": 75.0},       # WATERFALL
    "5": {"name": "Picture 4", "left": 354.39},      # STAIRCASE
    "19": {"name": "Picture 18", "left": 632.83},    # CHESS
}
TOP_PT = 135.0
SIZE_PT = 225.0


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def build_slide(data: bytes) -> bytes:
    root = etree.fromstring(data)
    namespaces = {"p": P_NS, "a": A_NS}
    common_slide = root.find("p:cSld", namespaces)
    if common_slide is None:
        raise ValueError("slide2.xml 缺少 p:cSld")
    if common_slide.find("p:bg", namespaces) is not None:
        raise ValueError("冻结输入的第 2 页不应已有显式背景")

    background = etree.Element(qn(P_NS, "bg"))
    background_properties = etree.SubElement(background, qn(P_NS, "bgPr"))
    solid_fill = etree.SubElement(background_properties, qn(A_NS, "solidFill"))
    etree.SubElement(solid_fill, qn(A_NS, "schemeClr"), val="bg1")
    etree.SubElement(background_properties, qn(A_NS, "effectLst"))
    common_slide.insert(0, background)

    seen: set[str] = set()
    for picture in root.xpath(".//p:pic", namespaces=namespaces):
        ids = picture.xpath("./p:nvPicPr/p:cNvPr/@id", namespaces=namespaces)
        names = picture.xpath("./p:nvPicPr/p:cNvPr/@name", namespaces=namespaces)
        if not ids or ids[0] not in TARGETS:
            continue
        shape_id = ids[0]
        expected = TARGETS[shape_id]
        if names != [expected["name"]]:
            raise ValueError(f"图片 {shape_id} 名称与审计记录不符：{names}")
        transform = picture.find("./p:spPr/a:xfrm", namespaces)
        offset = transform.find("a:off", namespaces) if transform is not None else None
        extent = transform.find("a:ext", namespaces) if transform is not None else None
        if transform is None or offset is None or extent is None:
            raise ValueError(f"图片 {shape_id} 缺少几何变换")
        transform.attrib.pop("rot", None)
        offset.set("x", str(round(float(expected["left"]) * EMU_PER_POINT)))
        offset.set("y", str(round(TOP_PT * EMU_PER_POINT)))
        extent.set("cx", str(round(SIZE_PT * EMU_PER_POINT)))
        extent.set("cy", str(round(SIZE_PT * EMU_PER_POINT)))
        seen.add(shape_id)

    if seen != set(TARGETS):
        raise ValueError(f"未完整定位三张目标图片：{sorted(seen)}")
    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if sha256(args.input) != INPUT_SHA256:
        raise ValueError("冻结输入 SHA256 不匹配")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(args.input) as archive_in, zipfile.ZipFile(args.output, "w") as archive_out:
        for info in archive_in.infolist():
            data = archive_in.read(info.filename)
            if info.filename == SLIDE_PART:
                data = build_slide(data)
            archive_out.writestr(info, data)
    print(args.output)
    print(sha256(args.output))


if __name__ == "__main__":
    main()
