from __future__ import annotations

import argparse
import io
import re
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

from PIL import Image
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from lxml import etree
from openpyxl import load_workbook


OUTPUT_NAME = "产品分析.docx"
SOURCE_FILES = [
    "product_yunxi.docx",
    "product_leader.docx",
    "product_littleswan.docx",
    "product_mijia.docx",
    "product_midea.docx",
]
CROP_BOXES = {
    "product_yunxi.docx": (65, 215, 198, 392),
    "product_leader.docx": (232, 150, 468, 487),
    "product_littleswan.docx": (92, 341, 365, 666),
}


@dataclass(frozen=True)
class Product:
    source: str
    name: str
    model: str
    price: int
    parameters: tuple[str, ...]
    scenario: str
    media: bytes


def source_text(document: Document) -> list[str]:
    return [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]


def value_after(lines: list[str], label: str) -> str:
    matches = [line.split("：", 1)[1].strip() for line in lines if line.startswith(label + "：")]
    assert len(matches) == 1, (label, matches)
    return matches[0]


def first_media(path: Path) -> bytes:
    with zipfile.ZipFile(path) as package:
        names = [name for name in package.namelist() if name.startswith("word/media/")]
        assert len(names) == 1, (path.name, names)
        return package.read(names[0])


def read_product(path: Path) -> Product:
    document = Document(path)
    lines = source_text(document)
    name = value_after(lines, "产品名称")
    model = value_after(lines, "产品型号")
    price_line = value_after(lines, "指导价格")
    price_match = re.search(r"(\d+)\s*元", price_line)
    assert price_match, price_line
    labels = ("洗涤容量", "宽度", "类型", "能效等级", "深度", "高度")
    parameters = tuple(line for line in lines if any(line.startswith(label + "：") for label in labels))
    scenarios = [line for line in lines if line.startswith("本产品适配多种场景")]
    assert len(scenarios) == 1
    return Product(path.name, name, model, int(price_match.group(1)), parameters, scenarios[0], first_media(path))


def verify_catalog_template(path: Path) -> None:
    workbook = load_workbook(path, data_only=False, read_only=True)
    sheet = workbook.active
    headers = [sheet.cell(1, col).value for col in range(1, 6)]
    workbook.close()
    assert headers == ["产品名称", "核心参数", "适用场景", "指导价格", "展示图"]


def shade(cell, color: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    element = properties.find(qn("w:shd")) or OxmlElement("w:shd")
    element.set(qn("w:fill"), color)
    if element.getparent() is None:
        properties.append(element)


def ensure_styles(document: Document) -> None:
    styles = document.styles
    if "Title" not in styles:
        style = styles.add_style("Title", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = "Noto Sans CJK SC"; style.font.size = Pt(22); style.font.bold = True; style.font.color.rgb = RGBColor(31, 78, 121)
    for name, size, level in [("Heading 1", 15, 0), ("Heading 2", 11, 1)]:
        if name not in styles:
            style = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            style.font.name = "Noto Sans CJK SC"; style.font.size = Pt(size); style.font.bold = True; style.font.color.rgb = RGBColor(31, 78, 121)
            properties = style.element.get_or_add_pPr(); outline = OxmlElement("w:outlineLvl"); outline.set(qn("w:val"), str(level)); properties.append(outline)
    if "Table Grid" not in styles:
        styles.add_style("Table Grid", WD_STYLE_TYPE.TABLE)


def table_grid(table) -> None:
    properties = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}"); border.set(qn("w:val"), "single"); border.set(qn("w:sz"), "4"); border.set(qn("w:space"), "0"); border.set(qn("w:color"), "A6A6A6"); borders.append(border)
    properties.append(borders)


def no_split(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def repeat_header(row) -> None:
    element = OxmlElement("w:tblHeader")
    element.set(qn("w:val"), "true")
    row._tr.get_or_add_trPr().append(element)


def bookmark(paragraph, name: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def link(paragraph, label: str, anchor: str) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.extend([color, underline])
    run.append(properties)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def field(paragraph, instruction: str, display: str) -> None:
    simple = OxmlElement("w:fldSimple")
    simple.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = display
    run.append(text)
    simple.append(run)
    paragraph._p.append(simple)


def bordered_header(paragraph) -> None:
    properties = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "1F4E79")
    borders.append(bottom)
    properties.append(borders)


def crop_machine(product: Product, destination: Path) -> None:
    with Image.open(io.BytesIO(product.media)) as image:
        image.convert("RGB").crop(CROP_BOXES[product.source]).save(destination, format="PNG", optimize=False)


def configure_running_parts(document: Document) -> None:
    header = document.sections[0].header
    header.is_linked_to_previous = False
    paragraph = header.paragraphs[0]
    paragraph.text = "产品分析 · 受控发布稿"
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.runs[0].font.color.rgb = RGBColor(0, 0, 0)
    paragraph.runs[0].font.size = Pt(9)
    bordered_header(paragraph)
    footer = document.sections[0].footer
    footer.is_linked_to_previous = False
    page = footer.paragraphs[0]
    page.alignment = WD_ALIGN_PARAGRAPH.CENTER
    field(page, " PAGE ", "1")
    page.add_run(" / ")
    field(page, " NUMPAGES ", "4")
    for section in document.sections[1:]:
        section.header.is_linked_to_previous = True
        section.footer.is_linked_to_previous = True


def put_product_row(row, ordinal: int, product: Product, image_path: Path, bookmark_id: int) -> str:
    left, right = row.cells
    left.text = ""
    right.text = ""
    left.vertical_alignment = right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    title = left.paragraphs[0]
    title.style = "Heading 2"
    title.add_run(f"产品{ordinal}：{product.name}")
    bookmark(title, f"product_{ordinal}", bookmark_id)
    for label, value in [
        ("产品名称", product.name), ("产品型号", product.model),
        ("核心参数", "；".join(product.parameters)), ("适用场景", product.scenario),
        ("指导价格", f"{product.price}元"),
    ]:
        paragraph = left.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(2)
        paragraph.add_run(label + "：").bold = True
        paragraph.add_run(value)
    navigation = left.add_paragraph()
    stem = product.source.removeprefix("product_").removesuffix(".docx")
    link(navigation, "查看全量排名证据", f"evidence_{stem}")
    picture_paragraph = right.paragraphs[0]
    picture_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    shape = picture_paragraph.add_run().add_picture(str(image_path), height=Cm(5.0))
    shape._inline.docPr.set("descr", f"{product.name}产品主体裁剪图，已移除营销文字")
    caption = right.add_paragraph("图片源：" + product.source + " 内嵌图；仅保留产品主体")
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption.runs[0].italic = True
    caption.runs[0].font.size = Pt(7.5)
    row.height = Cm(5.8)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    no_split(row)
    return f"指导价格：{product.price}元"


def inject_comments(path: Path, comments: list[tuple[str, str]]) -> None:
    wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns = {"w": wns}
    temporary = path.with_suffix(".comments.docx")
    with zipfile.ZipFile(path) as source:
        parts = {info.filename: source.read(info.filename) for info in source.infolist()}
        infos = {info.filename: info for info in source.infolist()}
    root = etree.fromstring(parts["word/document.xml"])
    comments_root = etree.Element(qn("w:comments"), nsmap={"w": wns})
    for comment_id, (anchor, note) in enumerate(comments):
        paragraphs = root.xpath("//w:p[contains(string(.), $anchor)]", namespaces=ns, anchor=anchor)
        assert len(paragraphs) == 1, (anchor, len(paragraphs))
        paragraph = paragraphs[0]
        runs = paragraph.xpath("./w:r[w:t]", namespaces=ns)
        assert runs
        start = OxmlElement("w:commentRangeStart"); start.set(qn("w:id"), str(comment_id))
        end = OxmlElement("w:commentRangeEnd"); end.set(qn("w:id"), str(comment_id))
        ref_run = OxmlElement("w:r")
        ref = OxmlElement("w:commentReference"); ref.set(qn("w:id"), str(comment_id)); ref_run.append(ref)
        first_index, last_index = paragraph.index(runs[0]), paragraph.index(runs[-1])
        paragraph.insert(first_index, start); paragraph.insert(last_index + 2, end); paragraph.insert(last_index + 3, ref_run)
        comment = etree.SubElement(comments_root, qn("w:comment")); comment.set(qn("w:id"), str(comment_id)); comment.set(qn("w:author"), "选品复核")
        cp = etree.SubElement(comment, qn("w:p")); cr = etree.SubElement(cp, qn("w:r")); etree.SubElement(cr, qn("w:t")).text = note
    parts["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    parts["word/comments.xml"] = etree.tostring(comments_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    content_types = etree.fromstring(parts["[Content_Types].xml"])
    if not content_types.xpath("//*[local-name()='Override'][@PartName='/word/comments.xml']"):
        override = etree.SubElement(content_types, f"{{{ct_ns}}}Override"); override.set("PartName", "/word/comments.xml"); override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    parts["[Content_Types].xml"] = etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone="yes")
    relationships = etree.fromstring(parts["word/_rels/document.xml.rels"])
    if not relationships.xpath("//*[local-name()='Relationship'][contains(@Type, '/comments')]"):
        used = {rel.get("Id") for rel in relationships}; number = 1
        while f"rId{number}" in used: number += 1
        rel = etree.SubElement(relationships, f"{{{rel_ns}}}Relationship"); rel.set("Id", f"rId{number}"); rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"); rel.set("Target", "comments.xml")
    parts["word/_rels/document.xml.rels"] = etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone="yes")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for name, payload in parts.items(): target.writestr(infos[name] if name in infos else name, payload)
    temporary.replace(path)


def build(input_dir: Path, output_dir: Path) -> None:
    verify_catalog_template(input_dir / "product_catalog_template.xlsx")
    products = [read_product(input_dir / name) for name in SOURCE_FILES]
    products.sort(key=lambda item: (-item.price, item.name))
    assert [item.price for item in products] == [4500, 3999, 3039, 999, 846]
    top_three = products[:3]
    document = Document(input_dir / "product_analysis_template.docx")
    assert [p.text for p in document.paragraphs][:1] == ["产品摘要"]
    assert len(document.tables) == 1 and len(document.tables[0].rows) == 3 and len(document.tables[0].columns) == 2
    ensure_styles(document)
    document.styles["Normal"].font.name = "Noto Sans CJK SC"; document.styles["Normal"].font.size = Pt(9)
    document.paragraphs[0].style = "Title"; document.paragraphs[0].text = "产品摘要"; document.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.paragraphs[1].text = "选品决策与文档发布工作稿"; document.paragraphs[1].alignment = WD_ALIGN_PARAGRAPH.CENTER; document.paragraphs[1].runs[0].bold = True
    document.paragraphs[2].text = "五份产品说明逐份核验；按指导价格降序取前三；价格不等于最终成交价。"; document.paragraphs[2].alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.paragraphs[3].text = ""; field(document.paragraphs[3], ' TOC \\o "1-2" \\h \\z \\u ', "目录（在 Word 中更新域）")
    product_table = document.tables[0]; product_table.style = "Table Grid"; table_grid(product_table)
    comments: list[tuple[str, str]] = []
    with tempfile.TemporaryDirectory(prefix="product-release-") as directory:
        temporary = Path(directory)
        for ordinal, (row, product) in enumerate(zip(product_table.rows, top_three), start=1):
            cropped = temporary / f"product-{ordinal}.png"; crop_machine(product, cropped)
            anchor = put_product_row(row, ordinal, product, cropped, ordinal)
            comments.append((anchor, f"排名依据：{product.source} 的“指导价格”段落；按五份来源数值降序后位列第{ordinal}。"))
        document.add_heading("1. 发布控制", 1)
        control = document.add_table(rows=4, cols=2); control.style = "Table Grid"; table_grid(control)
        for row, values in zip(control.rows, [("发布文件名", OUTPUT_NAME), ("决策规则", "指导价格降序，取前三；同价时按产品名称排序"), ("证据范围", "5 份产品说明 + 1 份表头模板 + 1 份 Word 模板"), ("声明", "只报告来源中的指导价格和参数；不把活动优惠视为确定成交条件")]):
            row.cells[0].text, row.cells[1].text = values; row.cells[0].paragraphs[0].runs[0].bold = True; shade(row.cells[0], "D9EAF7"); no_split(row)
        landscape = document.add_section(WD_ORIENT.LANDSCAPE); landscape.orientation = WD_ORIENT.LANDSCAPE; landscape.page_width, landscape.page_height = Inches(11.69), Inches(8.27); landscape.top_margin = landscape.bottom_margin = Inches(0.55); landscape.left_margin = landscape.right_margin = Inches(0.55)
        document.add_heading("2. 五产品排名与证据矩阵", 1)
        evidence = document.add_table(rows=1, cols=7); evidence.style = "Table Grid"; table_grid(evidence)
        for cell, label in zip(evidence.rows[0].cells, ["排名", "产品", "型号", "指导价格", "决策", "来源定位", "导航"]):
            cell.text = label; cell.paragraphs[0].runs[0].bold = True; cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255); shade(cell, "1F4E79")
        repeat_header(evidence.rows[0])
        for rank, product in enumerate(products, start=1):
            row = evidence.add_row(); decision = "入选 Word 前三" if rank <= 3 else "未入选：价格排名第 4/5"; locator = f"{product.source} → 产品基本信息/指导价格"
            for cell, value in zip(row.cells[:6], [str(rank), product.name, product.model, f"{product.price}元", decision, locator]):
                cell.text = value
                for run in cell.paragraphs[0].runs: run.font.size = Pt(7.5)
            stem = product.source.removeprefix("product_").removesuffix(".docx"); bookmark(row.cells[1].paragraphs[0], f"evidence_{stem}", 100 + rank)
            if rank <= 3: link(row.cells[6].paragraphs[0], f"返回产品{rank}", f"product_{rank}")
            else: row.cells[6].text = "排名留痕"
            no_split(row)
        document.add_heading("3. 决策复算", 1)
        gaps = document.add_table(rows=1, cols=4); gaps.style = "Table Grid"; table_grid(gaps)
        for cell, value in zip(gaps.rows[0].cells, ["比较", "高价", "低价", "差额"]): cell.text = value; cell.paragraphs[0].runs[0].bold = True; shade(cell, "D9EAF7")
        repeat_header(gaps.rows[0])
        for first, second in zip(products, products[1:]):
            row = gaps.add_row()
            for cell, value in zip(row.cells, [f"{first.name} vs {second.name}", f"{first.price}元", f"{second.price}元", f"{first.price - second.price}元"]): cell.text = value
            no_split(row)
        document.add_paragraph("入选边界是第 3 名 3039元 与第 4 名 999元，差额 2040元。没有使用来源以外的评分、销量或促销假设。")
        boundary = document.add_paragraph(); boundary.add_run("证据边界：").bold = True; boundary.add_run("来源中“建议入手价”后的补贴、以旧换新和活动表述都是条件性信息，本稿不将其改写为保证价。")
        configure_running_parts(document)
        document.core_properties.title = "产品分析——选品决策与发布工作稿"; document.core_properties.subject = "五份洗衣机产品说明的指导价排名、前三摘要和证据复核"; document.core_properties.author = "选品复核组"; document.core_properties.keywords = "产品分析,指导价格,证据矩阵,受控发布"; document.core_properties.comments = "从 OfficeVal 088 授权输入构建的可编辑 Word 发布工作稿"
        update = OxmlElement("w:updateFields"); update.set(qn("w:val"), "true"); document.settings.element.append(update)
        output_dir.mkdir(parents=True, exist_ok=True); destination = output_dir / OUTPUT_NAME; document.save(destination)
    inject_comments(destination, comments)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(); parser.add_argument("--input-dir", required=True, type=Path); parser.add_argument("--output-dir", required=True, type=Path); arguments = parser.parse_args(); build(arguments.input_dir, arguments.output_dir)
