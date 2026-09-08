from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

from PIL import Image
from docx import Document
from lxml import etree
from openpyxl import load_workbook


TASK_ID = "cowork-evo/officeval-088-product-release-dossier"
RUBRIC_VERSION = "officeval-088-product-release-dossier-v0.3"
JUDGE_VERSION = "officeval-088-product-release-dossier-deterministic-v0.3"
OUTPUT_NAME = "产品分析.docx"
WEIGHTS = {"R001": .10, "R002": .10, "R003": .22, "R004": .14, "R005": .12, "R006": .08, "R007": .08, "R008": .08, "R009": .05, "R010": .03}
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WPNS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
ANS = "http://schemas.openxmlformats.org/drawingml/2006/main"
RELNS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": WNS, "r": RNS, "wp": WPNS, "a": ANS}
SOURCE_FILES = ["product_yunxi.docx", "product_leader.docx", "product_littleswan.docx", "product_mijia.docx", "product_midea.docx"]
CROP_BOXES = {"product_yunxi.docx": (65, 215, 198, 392), "product_leader.docx": (232, 150, 468, 487), "product_littleswan.docx": (92, 341, 365, 666)}


@dataclass(frozen=True)
class Product:
    source: str
    name: str
    model: str
    price: int
    parameters: tuple[str, ...]
    scenario: str
    media: bytes


def criterion(identifier: str, score: float, evidence: str) -> dict:
    return {"criterion_id": identifier, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic"}


def value_after(lines: list[str], label: str) -> str:
    matches = [line.split("：", 1)[1].strip() for line in lines if line.startswith(label + "：")]
    if len(matches) != 1:
        raise ValueError(f"{label}条目数={len(matches)}")
    return matches[0]


def read_product(path: Path) -> Product:
    document = Document(path)
    lines = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
    price_line = value_after(lines, "指导价格")
    price_match = re.search(r"(\d+)\s*元", price_line)
    if not price_match:
        raise ValueError(price_line)
    labels = ("洗涤容量", "宽度", "类型", "能效等级", "深度", "高度")
    parameters = tuple(line for line in lines if any(line.startswith(label + "：") for label in labels))
    scenarios = [line for line in lines if line.startswith("本产品适配多种场景")]
    with zipfile.ZipFile(path) as package:
        media_names = [name for name in package.namelist() if name.startswith("word/media/")]
        if len(media_names) != 1:
            raise ValueError(f"{path.name}内嵌图数={len(media_names)}")
        media = package.read(media_names[0])
    return Product(path.name, value_after(lines, "产品名称"), value_after(lines, "产品型号"), int(price_match.group(1)), parameters, scenarios[0], media)


def derive_truth(source_dir: Path) -> list[Product]:
    workbook = load_workbook(source_dir / "product_catalog_template.xlsx", data_only=False, read_only=True)
    headers = [workbook.active.cell(1, col).value for col in range(1, 6)]
    workbook.close()
    if headers != ["产品名称", "核心参数", "适用场景", "指导价格", "展示图"]:
        raise ValueError("表头模板不符合冻结证据")
    template = Document(source_dir / "product_analysis_template.docx")
    if not template.paragraphs or template.paragraphs[0].text != "产品摘要" or len(template.tables) != 1:
        raise ValueError("Word 模板骨架不符合冻结证据")
    products = [read_product(source_dir / name) for name in SOURCE_FILES]
    return sorted(products, key=lambda item: (-item.price, item.name))


def package_parts(path: Path) -> tuple[dict[str, bytes], etree._Element, dict[str, str]]:
    with zipfile.ZipFile(path) as package:
        if package.testzip():
            raise ValueError("压缩包 CRC 损坏")
        parts = {name: package.read(name) for name in package.namelist()}
    root = etree.fromstring(parts["word/document.xml"])
    relationships: dict[str, str] = {}
    if "word/_rels/document.xml.rels" in parts:
        rel_root = etree.fromstring(parts["word/_rels/document.xml.rels"])
        for rel in rel_root.findall(f"{{{RELNS}}}Relationship"):
            relationships[rel.get("Id", "")] = rel.get("Target", "")
    return parts, root, relationships


def paragraphs_text(root: etree._Element) -> str:
    return "\n".join("".join(node.text or "" for node in p.xpath(".//w:t", namespaces=NS)) for p in root.xpath("//w:p", namespaces=NS))


def locate_table(document: Document, headers: list[str]):
    for table in document.tables:
        if not table.rows:
            continue
        row = [cell.text.strip() for cell in table.rows[0].cells]
        if row[:len(headers)] == headers:
            return table
    return None


def score_product_blocks(document: Document, products: list[Product]) -> tuple[float, str]:
    if not document.tables or len(document.tables[0].rows) != 3 or len(document.tables[0].columns) != 2:
        return 0.0, "首表非 3×2 产品模板"
    hits, total, details = 0, 0, []
    for rank, (row, product) in enumerate(zip(document.tables[0].rows, products[:3]), start=1):
        text = row.cells[0].text
        checks = [f"产品{rank}：{product.name}" in text, product.name in text, product.model in text, f"指导价格：{product.price}元" in text, all(value in text for value in product.parameters), product.scenario in text]
        hits += sum(checks); total += len(checks); details.append(f"#{rank}={sum(checks)}/{len(checks)}")
    order_text = [row.cells[0].text for row in document.tables[0].rows]
    order_ok = all(product.name in order_text[index] for index, product in enumerate(products[:3]))
    hits += int(order_ok); total += 1
    return hits / total, ", ".join(details) + f", order={order_ok}"


def score_evidence(document: Document, products: list[Product]) -> tuple[float, str]:
    table = locate_table(document, ["排名", "产品", "型号", "指导价格", "决策", "来源定位", "导航"])
    if table is None or len(table.rows) != 6:
        return 0.0, "未找到 5 行证据矩阵"
    hits, total, details = 0, 0, []
    for rank, (row, product) in enumerate(zip(table.rows[1:], products), start=1):
        cells = [cell.text for cell in row.cells]
        decision_ok = "入选" in cells[4] if rank <= 3 else "未入选" in cells[4]
        checks = [cells[0].strip() == str(rank), product.name in cells[1], product.model in cells[2], cells[3].strip() == f"{product.price}元", decision_ok, product.source in cells[5], "指导价格" in cells[5]]
        hits += sum(checks); total += len(checks); details.append(f"#{rank}={sum(checks)}/7")
    return hits / total, ", ".join(details)


def dhash(image: Image.Image) -> int:
    pixels = list(image.convert("L").resize((9, 8)).getdata())
    result = 0
    for row in range(8):
        for col in range(8):
            result = (result << 1) | int(pixels[row * 9 + col] > pixels[row * 9 + col + 1])
    return result


def hamming(first: int, second: int) -> int:
    return (first ^ second).bit_count()


def score_images(parts: dict[str, bytes], root: etree._Element, rels: dict[str, str], products: list[Product]) -> tuple[float, str]:
    product_rows = root.xpath("//w:tbl[1]/w:tr", namespaces=NS)
    hits, total, details = 0, 0, []
    for index, product in enumerate(products[:3]):
        total += 5
        if index >= len(product_rows):
            details.append(f"#{index + 1}=0/5"); continue
        row = product_rows[index]
        drawings = row.xpath(".//wp:inline|.//wp:anchor", namespaces=NS)
        if len(drawings) != 1:
            details.append(f"#{index + 1}=0/5"); continue
        drawing = drawings[0]
        blips = drawing.xpath(".//a:blip", namespaces=NS)
        rid = blips[0].get(f"{{{RNS}}}embed") if blips else ""
        target = rels.get(rid, "")
        media_name = "word/" + target if target and not target.startswith("word/") else target
        media_name = media_name.replace("word/../", "")
        raw = parts.get(media_name, b"")
        extents = drawing.xpath("./wp:extent", namespaces=NS)
        descr_nodes = drawing.xpath("./wp:docPr", namespaces=NS)
        checks = [bool(raw)]
        if raw:
            with Image.open(io.BytesIO(raw)) as output_image, Image.open(io.BytesIO(product.media)) as source_image:
                expected = source_image.convert("RGB").crop(CROP_BOXES[product.source])
                ratio = output_image.width / output_image.height
                expected_ratio = expected.width / expected.height
                checks.append(abs(ratio - expected_ratio) <= 0.035)
                checks.append(hamming(dhash(output_image), dhash(expected)) <= 18)
                checks.append(raw != product.media)
        else:
            checks.extend([False, False, False])
        height_cm = int(extents[0].get("cy", "0")) / 360000 if extents else 0.0
        checks.append(4.8 <= height_cm <= 5.2 and bool(descr_nodes) and product.name in (descr_nodes[0].get("descr", "") if descr_nodes else ""))
        hits += sum(checks); details.append(f"#{index + 1}={sum(checks)}/5")
    return hits / total, ", ".join(details)


def score_comments(parts: dict[str, bytes], root: etree._Element, products: list[Product]) -> tuple[float, str]:
    if "word/comments.xml" not in parts:
        return 0.0, "无原生批注部件"
    comments_root = etree.fromstring(parts["word/comments.xml"])
    comments = comments_root.xpath("//w:comment", namespaces=NS)
    hits, total = 0, 6
    for rank, product in enumerate(products[:3], start=1):
        notes = ["".join(node.text or "" for node in comment.xpath(".//w:t", namespaces=NS)) for comment in comments]
        hits += int(any(product.source in note and f"第{rank}" in note for note in notes))
        ranges = root.xpath("//w:commentRangeStart", namespaces=NS)
        anchor_ok = False
        for start in ranges:
            paragraph = start.getparent()
            text = "".join(node.text or "" for node in paragraph.xpath(".//w:t", namespaces=NS))
            if f"指导价格：{product.price}元" in text:
                anchor_ok = True
        hits += int(anchor_ok)
    return hits / total, f"comments={len(comments)}, checks={hits}/{total}"


def score_navigation(root: etree._Element, products: list[Product]) -> tuple[float, str]:
    bookmarks = {node.get(f"{{{WNS}}}name", "") for node in root.xpath("//w:bookmarkStart", namespaces=NS)}
    anchors = [node.get(f"{{{WNS}}}anchor", "") for node in root.xpath("//w:hyperlink[@w:anchor]", namespaces=NS)]
    required = {"product_1", "product_2", "product_3"} | {f"evidence_{product.source.removeprefix('product_').removesuffix('.docx')}" for product in products}
    toc_nodes = [node.text or "" for node in root.xpath("//w:instrText", namespaces=NS)]
    toc_attrs = [str(value) for value in root.xpath("//w:fldSimple/@w:instr", namespaces=NS)]
    toc = " ".join(toc_nodes + toc_attrs).upper()
    checks = [required <= bookmarks, all(anchor in bookmarks for anchor in anchors), len(anchors) >= 6, "TOC" in toc]
    return sum(checks) / len(checks), f"bookmarks={len(bookmarks)}, required={len(required)}, links={len(anchors)}, toc={'TOC' in toc}"


def score_word_objects(parts: dict[str, bytes], root: etree._Element, document: Document) -> tuple[float, str]:
    sections = root.xpath("//w:sectPr", namespaces=NS)
    landscape = any(node.xpath("./w:pgSz[@w:orient='landscape']", namespaces=NS) for node in sections)
    headers = " ".join(paragraphs_text(etree.fromstring(payload)) for name, payload in parts.items() if name.startswith("word/header") and name.endswith(".xml"))
    header_border = any(etree.fromstring(payload).xpath("//w:pBdr/w:bottom", namespaces=NS) for name, payload in parts.items() if name.startswith("word/header") and name.endswith(".xml"))
    footer_values: list[str] = []
    for name, payload in parts.items():
        if name.startswith("word/footer") and name.endswith(".xml"):
            footer_root = etree.fromstring(payload)
            footer_values.extend(node.text or "" for node in footer_root.xpath("//w:instrText", namespaces=NS))
            footer_values.extend(str(value) for value in footer_root.xpath("//w:fldSimple/@w:instr", namespaces=NS))
    footer_instr = " ".join(footer_values).upper()
    product_rows = root.xpath("//w:tbl[1]/w:tr", namespaces=NS)
    no_split = len(product_rows) == 3 and all(row.xpath("./w:trPr/w:cantSplit", namespaces=NS) for row in product_rows)
    repeat_headers = bool(root.xpath("//w:tbl/w:tr[1]/w:trPr/w:tblHeader", namespaces=NS))
    checks = [len(sections) == 2, landscape, "产品分析" in headers and header_border, "PAGE" in footer_instr and "NUMPAGES" in footer_instr, no_split, repeat_headers]
    return sum(checks) / len(checks), f"sections={len(sections)}, landscape={landscape}, header={('产品分析' in headers) and header_border}, fields={footer_instr.strip()}, no_split={no_split}, repeat_header={repeat_headers}"


def score_integrity(parts: dict[str, bytes], root: etree._Element, document: Document) -> tuple[float, str]:
    text = paragraphs_text(root)
    external = False
    for name, payload in parts.items():
        if name.endswith(".rels"):
            rel_root = etree.fromstring(payload)
            external = external or any(rel.get("TargetMode") == "External" for rel in rel_root.findall(f"{{{RELNS}}}Relationship"))
    unsafe = any(name.endswith("vbaProject.bin") or "oleObject" in name or name.endswith("/afchunk.mht") for name in parts)
    properties = document.core_properties
    forbidden = [term for term in ["销量第一", "五星评分", "现货保证", "最低成交价保证", "供应商承诺必然"] if term in text]
    checks = [len(document.tables) >= 4 and len(text) > 1200, not external, not unsafe, not forbidden, "产品分析" in (properties.title or ""), "选品复核组" == (properties.author or "")]
    return sum(checks) / len(checks), f"tables={len(document.tables)}, text={len(text)}, external={external}, unsafe={unsafe}, forbidden={forbidden}, title={properties.title!r}, author={properties.author!r}"


def score_recalculation(document: Document, products: list[Product]) -> tuple[float, str]:
    table = locate_table(document, ["比较", "高价", "低价", "差额"])
    if table is None or len(table.rows) != 5:
        return 0.0, "无完整决策复算表"
    hits, total = 0, 8
    for row, first, second in zip(table.rows[1:], products, products[1:]):
        text = " | ".join(cell.text for cell in row.cells)
        hits += int(first.name in text and second.name in text)
        hits += int(f"{first.price - second.price}元" in text)
    return hits / total, f"pair_checks={hits}/{total}"


def main() -> None:
    parser = argparse.ArgumentParser(); parser.add_argument("--artifact-dir", required=True, type=Path); parser.add_argument("--result", required=True, type=Path); parser.add_argument("--reward-file", required=True, type=Path); args = parser.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    out = args.artifact_dir / OUTPUT_NAME
    artifacts = [path for path in args.artifact_dir.iterdir() if not path.name.startswith("~$")] if args.artifact_dir.is_dir() else []
    output_docx = [path for path in artifacts if path.suffix.lower() == ".docx"]
    results: list[dict] = []
    try:
        products = derive_truth(source_dir)
        file_set_ok = len(artifacts) == 1 and len(output_docx) == 1 and output_docx[0].name == OUTPUT_NAME
        if not file_set_ok or not out.is_file() or not zipfile.is_zipfile(out):
            raise ValueError(f"交付文件={[path.name for path in output_docx]}")
        parts, root, rels = package_parts(out)
        required = {"[Content_Types].xml", "word/document.xml", "word/styles.xml"}
        if not required <= parts.keys():
            raise ValueError("缺少核心 OOXML 部件")
        document = Document(out)
        results.append(criterion("R001", 1.0, f"唯一交付={OUTPUT_NAME}，OOXML 可解析，members={len(parts)}"))
        first_table = document.tables[0] if document.tables else None
        template_checks = [bool(document.paragraphs and document.paragraphs[0].text == "产品摘要"), bool(first_table and len(first_table.rows) == 3 and len(first_table.columns) == 2), bool(first_table and all(f"产品{i}：" in row.cells[0].text for i, row in enumerate(first_table.rows, start=1)))]
        results.append(criterion("R002", sum(template_checks) / len(template_checks), f"模板标题/3×2骨架/产品标签={template_checks}"))
        score, evidence = score_product_blocks(document, products); results.append(criterion("R003", score, evidence))
        score, evidence = score_evidence(document, products); results.append(criterion("R004", score, evidence))
        score, evidence = score_images(parts, root, rels, products); results.append(criterion("R005", score, evidence))
        score, evidence = score_comments(parts, root, products); results.append(criterion("R006", score, evidence))
        score, evidence = score_navigation(root, products); results.append(criterion("R007", score, evidence))
        score, evidence = score_word_objects(parts, root, document); results.append(criterion("R008", score, evidence))
        score, evidence = score_integrity(parts, root, document); results.append(criterion("R009", score, evidence))
        score, evidence = score_recalculation(document, products); results.append(criterion("R010", score, evidence))
    except Exception as exc:
        results = [criterion("R001", 0.0, f"读取或文件集门槛失败：{exc}")]
    values = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[key] * values.get(key, 0.0) for key in WEIGHTS)
    hurdle_failures = []
    if values.get("R001", 0.0) < 1.0:
        hurdle_failures.append("R001"); reward = 0.0
    caps = {"R002": 0.35, "R003": 0.55, "R004": 0.72, "R005": 0.78, "R006": 0.78, "R008": 0.79, "R009": 0.75, "R010": 0.78}
    for identifier, cap in caps.items():
        if values.get(identifier, 0.0) < 1.0:
            hurdle_failures.append(identifier); reward = min(reward, cap)
    reward = round(reward, 6)
    payload = {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": reward >= 0.8 and not hurdle_failures, "hurdle_failures": hurdle_failures}
    args.result.parent.mkdir(parents=True, exist_ok=True); args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.parent.mkdir(parents=True, exist_ok=True); args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if payload["passed"] else 1)


if __name__ == "__main__":
    main()
