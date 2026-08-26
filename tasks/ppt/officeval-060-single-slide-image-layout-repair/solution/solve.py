#!/usr/bin/env python3
"""officeval_060 逐题解法：仅修复第 4 页背景图的比例与全出血布局。"""

from __future__ import annotations

import argparse
import hashlib
import zipfile
from copy import copy
from pathlib import Path

from lxml import etree


INPUT_NAME = "PPT.pptx"
INPUT_SHA256 = "83d6462ef25e2682a6a46055bcc6a002c2c36d7aacaa98ea99ca449faf46a51b"
OUTPUT_NAME = "PPT_slide4_relayout.pptx"
SLIDE4_XML = "ppt/slides/slide4.xml"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
SLIDE_WIDTH_EMU = 12_192_000
SLIDE_HEIGHT_EMU = 6_858_000


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def repair_slide4(xml_bytes: bytes) -> bytes:
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(xml_bytes, parser)
    pictures = root.xpath(".//p:pic", namespaces={"p": P_NS})
    if len(pictures) != 1:
        raise ValueError(f"第 4 页预期仅 1 张背景图，实际为 {len(pictures)} 张")
    picture = pictures[0]
    xfrm = picture.find(f"{{{P_NS}}}spPr/{{{A_NS}}}xfrm")
    if xfrm is None:
        raise ValueError("第 4 页背景图缺少 xfrm")
    off = xfrm.find(f"{{{A_NS}}}off")
    ext = xfrm.find(f"{{{A_NS}}}ext")
    if off is None or ext is None:
        raise ValueError("第 4 页背景图缺少位置或尺寸")
    off.set("x", "0"); off.set("y", "0")
    ext.set("cx", str(SLIDE_WIDTH_EMU)); ext.set("cy", str(SLIDE_HEIGHT_EMU))

    # 原始位图是 1467×941，比 16:9 更窄。对上下各裁 6.154%，
    # 再铺满幻灯片，可在不拉伸设备的前提下消除两侧留白。
    blip_fill = picture.find(f"{{{P_NS}}}blipFill")
    if blip_fill is None:
        raise ValueError("第 4 页背景图缺少 blipFill")
    old_src = blip_fill.find(f"{{{A_NS}}}srcRect")
    if old_src is not None:
        blip_fill.remove(old_src)
    src_rect = etree.Element(f"{{{A_NS}}}srcRect")
    src_rect.set("t", "6154")
    src_rect.set("b", "6154")
    stretch = blip_fill.find(f"{{{A_NS}}}stretch")
    blip_fill.insert(blip_fill.index(stretch) if stretch is not None else len(blip_fill), src_rect)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def solve(source: Path, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source, "r") as reader, zipfile.ZipFile(output, "w") as writer:
        names = set(reader.namelist())
        if SLIDE4_XML not in names:
            raise ValueError("输入 PPTX 缺少第 4 页 XML")
        for info in reader.infolist():
            data = reader.read(info.filename)
            if info.filename == SLIDE4_XML:
                data = repair_slide4(data)
            copied_info = copy(info)
            writer.writestr(copied_info, data)


def main() -> None:
    parser = argparse.ArgumentParser(); parser.add_argument("--input-dir", type=Path, required=True); parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args(); source = args.input_dir / INPUT_NAME
    if not source.is_file() or sha256(source) != INPUT_SHA256:
        raise ValueError("源 PPTX 缺失或 SHA256 不匹配")
    output = args.output_dir / OUTPUT_NAME; solve(source, output); print(output)


if __name__ == "__main__": main()
