#!/usr/bin/env python3
"""选品发布题的任务特有负例校准。"""

from __future__ import annotations

import argparse
import base64
import json
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Inches
from lxml import etree


TASK_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_NAME = "产品分析.docx"
PYTHON = sys.executable
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WNS}


def score(case: Path) -> float:
    result = case.parent / f"{case.name}-result.json"
    reward = case.parent / f"{case.name}-reward.txt"
    subprocess.run([PYTHON, str(TASK_ROOT / "tests/judge.py"), "--artifact-dir", str(case), "--result", str(result), "--reward-file", str(reward)], check=False, timeout=360)
    return float(json.loads(result.read_text(encoding="utf-8"))["reward"])


def clone(golden: Path, root: Path, name: str) -> Path:
    case = root / name; case.mkdir(); shutil.copy2(golden, case / OUTPUT_NAME); return case


def rewrite(path: Path, mutator) -> None:
    temporary = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(path) as source:
        parts = {info.filename: source.read(info.filename) for info in source.infolist()}
        infos = {info.filename: info for info in source.infolist()}
    mutator(parts)
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for name, payload in parts.items(): target.writestr(infos[name] if name in infos else name, payload)
    temporary.replace(path)


def replace_text(parts: dict[str, bytes], before: str, after: str) -> None:
    root = etree.fromstring(parts["word/document.xml"])
    count = 0
    for node in root.xpath("//w:t", namespaces=NS):
        if before in (node.text or ""):
            node.text = (node.text or "").replace(before, after); count += 1
    assert count
    parts["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def document_mutation(parts: dict[str, bytes], mutator) -> None:
    root = etree.fromstring(parts["word/document.xml"]); mutator(root); parts["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def remove_navigation(root) -> None:
    for node in root.xpath("//w:bookmarkStart | //w:bookmarkEnd", namespaces=NS): node.getparent().remove(node)
    for hyperlink in root.xpath("//w:hyperlink[@w:anchor]", namespaces=NS):
        parent = hyperlink.getparent(); position = parent.index(hyperlink)
        for child in list(hyperlink): parent.insert(position, child); position += 1
        parent.remove(hyperlink)


def run(retain_dir: Path | None = None) -> None:
    with tempfile.TemporaryDirectory(prefix="product-release-negative-") as temporary:
        root = Path(temporary); fresh = root / "fresh"; fresh.mkdir()
        subprocess.run([PYTHON, str(TASK_ROOT / "solution/solve.py"), "--input-dir", str(TASK_ROOT / "data/input_files"), "--output-dir", str(fresh)], check=True, timeout=180)
        golden = fresh / OUTPUT_NAME; cases: dict[str, float] = {}

        missing = root / "missing_output"; missing.mkdir(); cases[missing.name] = score(missing)
        corrupt = root / "corrupt_docx"; corrupt.mkdir(); (corrupt / OUTPUT_NAME).write_bytes(b"not a docx"); cases[corrupt.name] = score(corrupt)
        extra = clone(golden, root, "extra_artifact"); (extra / "selection.txt").write_text("extra", encoding="utf-8"); cases[extra.name] = score(extra)
        wrong_name = root / "wrong_release_filename"; wrong_name.mkdir(); shutil.copy2(golden, wrong_name / "产品分析_副本.docx"); cases[wrong_name.name] = score(wrong_name)

        title = clone(golden, root, "template_title_destroyed")
        rewrite(title / OUTPUT_NAME, lambda parts: replace_text(parts, "产品摘要", "选品简报")); cases[title.name] = score(title)

        order = clone(golden, root, "top_order_swapped")
        def swap_rows(root_xml):
            table = root_xml.xpath("//w:body/w:tbl[1]", namespaces=NS)[0]; rows = table.xpath("./w:tr", namespaces=NS); rows[0].addprevious(rows[1])
        rewrite(order / OUTPUT_NAME, lambda parts: document_mutation(parts, swap_rows)); cases[order.name] = score(order)

        price = clone(golden, root, "leader_price_wrong")
        rewrite(price / OUTPUT_NAME, lambda parts: replace_text(parts, "3999元", "3599元")); cases[price.name] = score(price)

        parameter = clone(golden, root, "yunxi_height_parameter_missing")
        rewrite(parameter / OUTPUT_NAME, lambda parts: replace_text(parts, "；高度：850mm", "")); cases[parameter.name] = score(parameter)

        matrix = clone(golden, root, "fifth_evidence_row_missing")
        def delete_last_evidence(root_xml):
            table = root_xml.xpath("//w:body/w:tbl[3]", namespaces=NS)[0]; rows = table.xpath("./w:tr", namespaces=NS); table.remove(rows[-1])
        rewrite(matrix / OUTPUT_NAME, lambda parts: document_mutation(parts, delete_last_evidence)); cases[matrix.name] = score(matrix)

        decision = clone(golden, root, "mijia_wrongly_included")
        rewrite(decision / OUTPUT_NAME, lambda parts: replace_text(parts, "未入选：价格排名第 4/5", "入选 Word 前三")); cases[decision.name] = score(decision)

        difference = clone(golden, root, "boundary_difference_wrong")
        rewrite(difference / OUTPUT_NAME, lambda parts: replace_text(parts, "2040元", "2000元")); cases[difference.name] = score(difference)

        poster = clone(golden, root, "uncropped_marketing_poster")
        def restore_poster(parts):
            with zipfile.ZipFile(TASK_ROOT / "data/input_files/product_yunxi.docx") as source:
                name = next(item for item in source.namelist() if item.startswith("word/media/")); parts["word/media/image1.png"] = source.read(name)
        rewrite(poster / OUTPUT_NAME, restore_poster); cases[poster.name] = score(poster)

        swapped = clone(golden, root, "product_images_swapped")
        def swap_images(parts): parts["word/media/image1.png"], parts["word/media/image2.png"] = parts["word/media/image2.png"], parts["word/media/image1.png"]
        rewrite(swapped / OUTPUT_NAME, swap_images); cases[swapped.name] = score(swapped)

        comments = clone(golden, root, "price_comments_removed")
        def strip_comments(parts):
            document_mutation(parts, lambda root_xml: [node.getparent().remove(node) for node in root_xml.xpath("//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference", namespaces=NS)])
        rewrite(comments / OUTPUT_NAME, strip_comments); cases[comments.name] = score(comments)

        footer = clone(golden, root, "page_fields_removed")
        def strip_fields(parts):
            for name in [item for item in parts if item.startswith("word/footer") and item.endswith(".xml")]:
                root_xml = etree.fromstring(parts[name])
                for node in root_xml.xpath("//w:fldSimple | //w:instrText", namespaces=NS): node.getparent().remove(node)
                parts[name] = etree.tostring(root_xml, xml_declaration=True, encoding="UTF-8", standalone="yes")
        rewrite(footer / OUTPUT_NAME, strip_fields); cases[footer.name] = score(footer)

        portrait = clone(golden, root, "ranking_section_portrait")
        def make_portrait(root_xml):
            section = root_xml.xpath("//w:sectPr", namespaces=NS)[-1]; size = section.xpath("./w:pgSz", namespaces=NS)[0]; size.attrib.pop(f"{{{WNS}}}orient", None); size.set(f"{{{WNS}}}w", "11906"); size.set(f"{{{WNS}}}h", "16838")
        rewrite(portrait / OUTPUT_NAME, lambda parts: document_mutation(parts, make_portrait)); cases[portrait.name] = score(portrait)

        fabricated = clone(golden, root, "fabricated_supplier_commitment")
        rewrite(fabricated / OUTPUT_NAME, lambda parts: replace_text(parts, "没有使用来源以外的评分、销量或促销假设。", "供应商承诺必然现货保证，且销量第一。")); cases[fabricated.name] = score(fabricated)

        flattened = root / "flattened_screenshot_only"; flattened.mkdir()
        pixel = root / "pixel.png"; pixel.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="))
        document = Document(); document.add_heading("产品摘要", 0); document.add_picture(str(pixel), width=Inches(6)); document.save(flattened / OUTPUT_NAME); cases[flattened.name] = score(flattened)

        medium = clone(golden, root, "medium_missing_navigation")
        rewrite(medium / OUTPUT_NAME, lambda parts: document_mutation(parts, remove_navigation)); cases[medium.name] = score(medium)

        assert all(cases[name] == 0 for name in ["missing_output", "corrupt_docx", "extra_artifact", "wrong_release_filename"]), cases
        for name in ["template_title_destroyed", "top_order_swapped", "leader_price_wrong", "yunxi_height_parameter_missing", "fifth_evidence_row_missing", "mijia_wrongly_included", "boundary_difference_wrong", "uncropped_marketing_poster", "product_images_swapped", "price_comments_removed", "page_fields_removed", "ranking_section_portrait", "fabricated_supplier_commitment", "flattened_screenshot_only"]:
            assert 0 <= cases[name] < 0.8, (name, cases)
        assert 0.8 < cases["medium_missing_navigation"] < 1.0, cases
        if retain_dir is not None: shutil.copytree(root, retain_dir)
        print(json.dumps(cases, ensure_ascii=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(); parser.add_argument("--retain-dir", type=Path); arguments = parser.parse_args(); run(arguments.retain_dir)
