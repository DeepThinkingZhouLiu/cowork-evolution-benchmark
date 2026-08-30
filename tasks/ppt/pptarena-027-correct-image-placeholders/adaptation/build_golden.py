#!/usr/bin/env python3
"""PPTArena Case 27 任务专属 Golden：只修改三个已填充图片占位符的几何。"""

from __future__ import annotations

import argparse
import hashlib
import zipfile
from pathlib import Path

from lxml import etree


INPUT_SHA256 = "026285dad008ed551128b04548909d8e4377324b67f8428cb6a68d2c87d9555f"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}
EMU_PER_POINT = 12_700
TARGETS = {
    "10": (86.4, 170.34, 229.88),
    "4": (366.87, 170.34, 229.88),
    "7": (647.33, 170.34, 229.88),
}


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def set_geometry(picture, left: float, top: float, size: float) -> None:
    properties = picture.find("p:spPr", NS)
    transform = properties.find("a:xfrm", NS)
    if transform is None:
        transform = etree.SubElement(properties, qn(A_NS, "xfrm"))
    offset = transform.find("a:off", NS)
    if offset is None:
        offset = etree.SubElement(transform, qn(A_NS, "off"))
    extent = transform.find("a:ext", NS)
    if extent is None:
        extent = etree.SubElement(transform, qn(A_NS, "ext"))
    offset.set("x", str(round(left * EMU_PER_POINT)))
    offset.set("y", str(round(top * EMU_PER_POINT)))
    extent.set("cx", str(round(size * EMU_PER_POINT)))
    extent.set("cy", str(round(size * EMU_PER_POINT)))
    transform.attrib.pop("rot", None)


def rewrite_slide(data: bytes) -> bytes:
    root = etree.fromstring(data)
    for shape_id, geometry in TARGETS.items():
        nodes = root.xpath(
            f"./p:cSld/p:spTree/p:pic[p:nvPicPr/p:cNvPr/@id='{shape_id}']",
            namespaces=NS,
        )
        if len(nodes) != 1:
            raise ValueError(f"未唯一定位图片占位符 {shape_id}")
        set_geometry(nodes[0], *geometry)
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


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
    print(f"{args.output}\t{args.output.stat().st_size}\t{sha256(args.output)}")


if __name__ == "__main__":
    main()
