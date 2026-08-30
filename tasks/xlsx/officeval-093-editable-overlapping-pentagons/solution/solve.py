#!/usr/bin/env python3
"""将两个可编辑 custGeom 五边形写入固定工作簿的 C32 锚点。"""
from __future__ import annotations

import argparse
import zipfile
from pathlib import Path

OUTPUT_NAME = "日常信息处理能力观察表_图形组织可编辑版.xlsx"
INPUT_NAME = "日常信息处理能力观察表.xlsx"

NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

LEFT = [(489, 1247), (1528, 1847), (1500, 836), (939, 0), (0, 357)]
RIGHT = [(1528 - x, y) for x, y in LEFT]


def shape(name: str, ident: int, x: int, points: list[tuple[int, int]]) -> str:
    point_xml = [f'<a:moveTo><a:pt x="{points[0][0]}" y="{points[0][1]}"/></a:moveTo>']
    point_xml.extend(f'<a:lnTo><a:pt x="{px}" y="{py}"/></a:lnTo>' for px, py in points[1:])
    point_xml.append("<a:close/>")
    return f'''<xdr:sp>
  <xdr:nvSpPr><xdr:cNvPr id="{ident}" name="{name}"/><xdr:cNvSpPr/></xdr:nvSpPr>
  <xdr:spPr>
    <a:xfrm><a:off x="{x}" y="0"/><a:ext cx="450000" cy="360000"/></a:xfrm>
    <a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>
      <a:pathLst><a:path w="2000" h="2000">{''.join(point_xml)}</a:path></a:pathLst>
    </a:custGeom>
    <a:noFill/><a:ln w="19050"><a:solidFill><a:srgbClr val="404040"/></a:solidFill><a:prstDash val="solid"/></a:ln>
  </xdr:spPr><xdr:style/><xdr:txBody><a:bodyPr/><a:lstStyle/></xdr:txBody>
</xdr:sp>'''


def drawing_xml() -> bytes:
    content = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="{NS_XDR}" xmlns:a="{NS_A}">
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>100000</xdr:colOff><xdr:row>31</xdr:row><xdr:rowOff>10000</xdr:rowOff></xdr:from>
    <xdr:ext cx="900000" cy="360000"/>
    <xdr:grpSp>
      <xdr:nvGrpSpPr><xdr:cNvPr id="1" name="Overlapping pentagons"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>
      <xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="900000" cy="360000"/><a:chOff x="0" y="0"/><a:chExt cx="900000" cy="360000"/></a:xfrm></xdr:grpSpPr>
      {shape('Left freeform pentagon', 2, 100000, LEFT)}
      {shape('Right freeform pentagon', 3, 500000, RIGHT)}
    </xdr:grpSp>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>'''
    return content.encode("utf-8")


def patch_sheet(xml: bytes) -> bytes:
    text = xml.decode("utf-8")
    drawing = '<drawing xmlns:r="%s" r:id="rId1"/>' % NS_R
    if "</worksheet>" not in text:
        raise ValueError("工作表 XML 缺少 worksheet 根元素")
    return text.replace("</worksheet>", drawing + "</worksheet>").encode("utf-8")


def patch_rels(xml: bytes | None) -> bytes:
    if xml is None:
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{NS_REL}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'''.encode()
    text = xml.decode("utf-8")
    return text.replace("</Relationships>", '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>').encode()


def patch_types(xml: bytes) -> bytes:
    text = xml.decode("utf-8")
    override = '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
    return text.replace("</Types>", override + "</Types>").encode()


def solve(input_dir: Path, output_dir: Path) -> None:
    source = input_dir / INPUT_NAME
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_NAME
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        names = set(zin.namelist())
        for name in zin.namelist():
            data = zin.read(name)
            if name == "xl/worksheets/sheet1.xml":
                data = patch_sheet(data)
            elif name == "xl/worksheets/_rels/sheet1.xml.rels":
                data = patch_rels(data)
            elif name == "[Content_Types].xml":
                data = patch_types(data)
            zout.writestr(name, data)
        if "xl/worksheets/_rels/sheet1.xml.rels" not in names:
            zout.writestr("xl/worksheets/_rels/sheet1.xml.rels", patch_rels(None))
        zout.writestr("xl/drawings/drawing1.xml", drawing_xml())


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input-dir", type=Path, required=True)
    ap.add_argument("--output-dir", type=Path, required=True)
    args = ap.parse_args()
    solve(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
