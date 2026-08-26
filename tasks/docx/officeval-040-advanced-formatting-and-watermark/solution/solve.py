#!/usr/bin/env python3
"""officeval_040 的逐题参考解法。"""

from __future__ import annotations

import argparse
import hashlib
import tempfile
import zipfile
from pathlib import Path

from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Pt
from lxml import etree


DOC_NAME = "简介素材.docx"
IMAGE_NAME = "Image.png"
OUTPUT_NAME = "简介素材_格式设置_水印版.docx"
DOC_SHA256 = "c498fd663bbbb1873ef436da8ebeb1d6a0fcffa6f3a98dd41e12839a3cc7acb9"
IMAGE_SHA256 = "d2f28e14e1b83b69d77339941cb28b5217fc6d407d26b1270331ac5227e7081f"
OLD_NAME = "智能工程系"
NEW_NAME = "智能软件工程系"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def set_font(run, name: str, size: float, *, bold: bool = False, italic: bool = False, color: str = "000000") -> None:
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rpr.rFonts.set(qn("w:eastAsia"), name)
    rpr.rFonts.set(qn("w:ascii"), name)
    rpr.rFonts.set(qn("w:hAnsi"), name)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = __import__("docx").shared.RGBColor.from_string(color)


def set_paragraph_format(paragraph, *, indent: bool = True) -> None:
    ppr = paragraph._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:after"), "120")
    spacing.set(qn("w:afterAutospacing"), "0")
    spacing.set(qn("w:line"), "264")
    spacing.set(qn("w:lineRule"), "auto")
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    if indent:
        ind.set(qn("w:firstLineChars"), "200")
        ind.set(qn("w:hanging"), "0")
    else:
        ind.set(qn("w:firstLineChars"), "0")


def clear_paragraph(paragraph) -> None:
    paragraph._p.clear_content()


def add_special_name(paragraph) -> None:
    run = paragraph.add_run(NEW_NAME)
    set_font(run, "宋体", 14, bold=True, italic=True, color="00B050")
    rpr = run._element.get_or_add_rPr()
    emphasis = OxmlElement("w:em")
    emphasis.set(qn("w:val"), "dot")
    rpr.append(emphasis)


def set_body_text(paragraph, text: str) -> None:
    clear_paragraph(paragraph)
    parts = text.replace(OLD_NAME, NEW_NAME).split(NEW_NAME)
    for index, part in enumerate(parts):
        if part:
            run = paragraph.add_run(part)
            set_font(run, "楷体", 12)
        if index < len(parts) - 1:
            add_special_name(paragraph)
    set_paragraph_format(paragraph)


def set_title(paragraph) -> None:
    clear_paragraph(paragraph)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ppr = paragraph._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    spacing.set(qn("w:after"), "400")
    spacing.set(qn("w:afterAutospacing"), "0")
    run = paragraph.add_run("智能软件工程系简介")
    set_font(run, "隶书", 22, bold=True, color="0070C0")
    rpr = run._element.get_or_add_rPr()
    glow = OxmlElement("w14:glow")
    glow.set(qn("w14:rad"), "88900")
    scheme = OxmlElement("w14:schemeClr")
    scheme.set(qn("w14:val"), "accent3")
    glow.append(scheme)
    rpr.append(glow)
    shadow = OxmlElement("w14:shadow")
    shadow.set(qn("w14:blurRad"), "50800")
    shadow.set(qn("w14:dist"), "38100")
    shadow.set(qn("w14:dir"), "16200000")
    shadow.set(qn("w14:sx"), "100000")
    shadow.set(qn("w14:sy"), "100000")
    shadow.set(qn("w14:algn"), "ctr")
    shadow_color = OxmlElement("w14:srgbClr")
    shadow_color.set(qn("w14:val"), "000000")
    alpha = OxmlElement("w14:alpha")
    alpha.set(qn("w14:val"), "40000")
    shadow_color.append(alpha)
    shadow.append(shadow_color)
    rpr.append(shadow)


def set_drop_cap(first_paragraph, remainder_paragraph) -> None:
    clear_paragraph(first_paragraph)
    first_run = first_paragraph.add_run("专")
    set_font(first_run, "楷体", 24, bold=True)
    first_ppr = first_paragraph._p.get_or_add_pPr()
    frame = OxmlElement("w:framePr")
    frame.set(qn("w:dropCap"), "margin")
    frame.set(qn("w:lines"), "2")
    frame.set(qn("w:hSpace"), "170")
    frame.set(qn("w:wrap"), "around")
    frame.set(qn("w:vAnchor"), "text")
    frame.set(qn("w:hAnchor"), "text")
    first_ppr.insert(0, frame)

    clear_paragraph(remainder_paragraph)
    run = remainder_paragraph.add_run("业设置如下：")
    set_font(run, "楷体", 12)
    set_paragraph_format(remainder_paragraph, indent=False)


def add_watermark(document: Document, image_path: Path) -> None:
    section = document.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    paragraph = header.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clear_paragraph(paragraph)

    with tempfile.TemporaryDirectory(prefix="officeval040-watermark-") as tmp:
        prepared = Path(tmp) / "campus-watermark.png"
        with Image.open(image_path) as image:
            source_rgb = image.convert("RGB")
            white = Image.new("RGB", source_rgb.size, "white")
            # Word 会读取下方的 VML 冲蚀属性；LibreOffice 不执行该滤镜，
            # 因此同时保存一份由真实输入确定性淡化得到的像素版本，确保跨套件可读。
            washed = Image.blend(source_rgb, white, 0.78)
            washed.save(prepared, format="PNG", dpi=(240, 240))
            pixel_width, pixel_height = image.size
        rid, _ = header.part.get_or_add_image(str(prepared))

    width_pt = pixel_width / 240.0 * 72.0
    height_pt = pixel_height / 240.0 * 72.0
    run = paragraph.add_run()
    pict = parse_xml(
        f'<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'xmlns:w10="urn:schemas-microsoft-com:office:word">'
        f'<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s1025" stroked="f" '
        f'style="position:absolute;z-index:-251658752;mso-position-horizontal:center;'
        f'mso-position-horizontal-relative:page;mso-position-vertical:center;'
        f'mso-position-vertical-relative:page;width:{width_pt:.2f}pt;height:{height_pt:.2f}pt">'
        f'<v:imagedata r:id="{rid}" o:title="campus-watermark" o:gain="19661f" o:blacklevel="22938f"/>'
        f'<w10:wrap type="none"/>'
        f'</v:shape></w:pict>'
    )
    run._r.append(pict)


def add_missing_office_theme(path: Path) -> None:
    """为上游极简 DOCX 补回标准 Office 主题，使 accent3 可被正常解析。"""
    with tempfile.TemporaryDirectory(prefix="officeval040-theme-") as tmp:
        default_docx = Path(tmp) / "default.docx"
        Document().save(default_docx)
        with zipfile.ZipFile(default_docx) as template_zip:
            theme_xml = template_zip.read("word/theme/theme1.xml")

        with zipfile.ZipFile(path) as source_zip:
            members = [(item, source_zip.read(item.filename)) for item in source_zip.infolist()]
        payloads = {item.filename: payload for item, payload in members}

        rels = etree.fromstring(payloads["word/_rels/document.xml.rels"])
        relationship_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        existing_ids = {rel.get("Id") for rel in rels}
        relationship_id = "rIdTheme"
        suffix = 1
        while relationship_id in existing_ids:
            relationship_id = f"rIdTheme{suffix}"
            suffix += 1
        relation = etree.SubElement(rels, f"{{{relationship_ns}}}Relationship")
        relation.set("Id", relationship_id)
        relation.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        relation.set("Target", "theme/theme1.xml")
        payloads["word/_rels/document.xml.rels"] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone="yes")

        content_types = etree.fromstring(payloads["[Content_Types].xml"])
        content_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        if not content_types.xpath("./ct:Override[@PartName='/word/theme/theme1.xml']", namespaces={"ct": content_ns}):
            override = etree.SubElement(content_types, f"{{{content_ns}}}Override")
            override.set("PartName", "/word/theme/theme1.xml")
            override.set("ContentType", "application/vnd.openxmlformats-officedocument.theme+xml")
        payloads["[Content_Types].xml"] = etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone="yes")
        payloads["word/theme/theme1.xml"] = theme_xml

        rebuilt = Path(tmp) / "rebuilt.docx"
        with zipfile.ZipFile(rebuilt, "w") as output_zip:
            for item, _ in members:
                output_zip.writestr(item, payloads[item.filename])
            output_zip.writestr("word/theme/theme1.xml", theme_xml)
        # 临时目录可能位于另一挂载点，按字节回写生成物，避免跨设备 rename。
        path.write_bytes(rebuilt.read_bytes())


def build_document(source_doc: Path, source_image: Path) -> Document:
    document = Document(source_doc)
    # 上游素材由极简 OOXML 生成，正文末尾没有 sectPr；水印页眉需要显式节属性。
    if not document.sections:
        document._element.body.get_or_add_sectPr()
    original = [paragraph.text for paragraph in document.paragraphs]
    if len(original) != 11 or "专业设置如下：" not in original[2]:
        raise ValueError("来源简介段落结构与已审读版本不一致")
    prefix, marker = original[2].split("专业设置如下：", 1)
    if marker:
        raise ValueError("专业设置标记后出现未预期文本")

    document.add_paragraph()
    paragraphs = document.paragraphs
    set_title(paragraphs[0])
    set_body_text(paragraphs[1], original[1] + prefix)
    set_drop_cap(paragraphs[2], paragraphs[3])
    for destination, text in zip(paragraphs[4:], original[3:]):
        set_body_text(destination, text)
    add_watermark(document, source_image)
    return document


def solve(input_dir: Path, output_dir: Path) -> Path:
    source_doc = input_dir / DOC_NAME
    source_image = input_dir / IMAGE_NAME
    if sha256(source_doc) != DOC_SHA256 or sha256(source_image) != IMAGE_SHA256:
        raise ValueError("输入 DOCX 或水印图片 SHA256 不匹配")
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_NAME
    build_document(source_doc, source_image).save(output)
    add_missing_office_theme(output)
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    print(solve(args.input_dir, args.output_dir))


if __name__ == "__main__":
    main()
