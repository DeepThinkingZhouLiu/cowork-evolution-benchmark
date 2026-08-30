#!/usr/bin/env python3
"""
PPTArena Case 41 的本地人工 Golden 重建器。

上游 Ground Truth 在添加渐变和可选图片边框时，使 LibreOffice
将中央 TIFF 热力图严重错渲染。本题从冻结输入出发，只在幻灯片
背景写入经人工审核的两端主题 accent1 浅蓝灰渐变，不触碰图片、
文字或关系部件。
"""

from __future__ import annotations

import argparse
import hashlib
import zipfile
from pathlib import Path

from lxml import etree


INPUT_NAME = "MatchBackgroundtoPicture_TestA.pptx"
OUTPUT_NAME = "MatchBackgroundtoPicture_GroundTruthA.pptx"
INPUT_SHA256 = "8c789d83d7063549ab3a198a7363e7e6d9feda84ea20865ae00a6307d3be52cd"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def make_background() -> etree._Element:
    background = etree.Element(qn(P_NS, "bg"))
    properties = etree.SubElement(background, qn(P_NS, "bgPr"))
    gradient = etree.SubElement(properties, qn(A_NS, "gradFill"))
    stops = etree.SubElement(gradient, qn(A_NS, "gsLst"))

    top = etree.SubElement(stops, qn(A_NS, "gs"), pos="0")
    top_color = etree.SubElement(top, qn(A_NS, "schemeClr"), val="accent1")
    etree.SubElement(top_color, qn(A_NS, "lumMod"), val="5000")
    etree.SubElement(top_color, qn(A_NS, "lumOff"), val="95000")

    bottom = etree.SubElement(stops, qn(A_NS, "gs"), pos="100000")
    bottom_color = etree.SubElement(
        bottom, qn(A_NS, "schemeClr"), val="accent1"
    )
    etree.SubElement(bottom_color, qn(A_NS, "lumMod"), val="30000")
    etree.SubElement(bottom_color, qn(A_NS, "lumOff"), val="70000")

    etree.SubElement(
        gradient,
        qn(A_NS, "lin"),
        ang="5400000",
        scaled="1",
    )
    etree.SubElement(properties, qn(A_NS, "effectLst"))
    return background


def rewrite_slide(data: bytes) -> bytes:
    root = etree.fromstring(data)
    common_slide_data = root.find(qn(P_NS, "cSld"))
    if common_slide_data is None:
        raise ValueError("第 1 页缺失 p:cSld")
    for old_background in common_slide_data.findall(qn(P_NS, "bg")):
        common_slide_data.remove(old_background)
    common_slide_data.insert(0, make_background())
    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=True,
    )


def build(source: Path, output: Path) -> None:
    if not source.is_file() or sha256(source) != INPUT_SHA256:
        raise ValueError("冻结输入缺失或 SHA256 不匹配")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as archive_in, zipfile.ZipFile(output, "w") as archive_out:
        for info in archive_in.infolist():
            data = archive_in.read(info.filename)
            if info.filename == "ppt/slides/slide1.xml":
                data = rewrite_slide(data)
            archive_out.writestr(info, data)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    build(args.source, args.output)
    print(f"{args.output}\t{sha256(args.output)}")


if __name__ == "__main__":
    main()
