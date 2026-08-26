#!/usr/bin/env python3
"""officeval_033 的逐题参考解法。

只识别这份价目册里五类已人工核对的价格表及其零售价列；不对文档中的
任意数字做全局替换。除 word/document.xml 中 1435 个价格文本外，压缩包
其他成员逐字节复制。
"""

from __future__ import annotations

import argparse
import hashlib
import re
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from lxml import etree


INPUT_NAME = "家居零售价目册.docx"
OUTPUT_NAME = "家居零售价目册_零售价乘以1.1.docx"
INPUT_SHA256 = "8511f025397f248ef8345b22e79390df010921f13e757838a34f936b8e37d035"
EXPECTED_COUNTS = {"combo": 206, "recommended": 153, "custom": 1057, "conversion": 8, "order": 11}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def normalize(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def cell_text(cell) -> str:
    return "".join(cell.xpath("./w:p//w:t/text()", namespaces=NS)).strip()


def classify(headers: list[str]) -> str:
    joined = "|".join(normalize(value) for value in headers)
    if all(value in joined for value in ["定制料号", "物料描述", "尺寸", "A1", "A2", "A3", "A4", "A5", "A6", "标准零售价"]):
        return "custom"
    if all(value in joined for value in ["ERP料号", "规格说明", "物料描述", "零售价（元）", "备注"]):
        return "combo"
    if all(value in joined for value in ["料号", "物料描述", "零售价（元）", "备注"]) and "ERP料号" not in joined:
        return "recommended"
    if all(value in joined for value in ["产品组", "基础零售价（元）", "材质调价", "尺寸调价", "核算后零售价（元）"]):
        return "conversion"
    if all(value in joined for value in ["序号", "产品编码", "产品名称", "零售价（元）"]):
        return "order"
    return "other"


def target_headers(kind: str) -> set[str]:
    return {
        "combo": {"零售价（元）"},
        "recommended": {"零售价（元）"},
        "custom": {"A1", "A2", "A3", "A4", "A5", "A6", "标准零售价"},
        "conversion": {"基础零售价（元）", "核算后零售价（元）"},
        "order": {"零售价（元）"},
    }.get(kind, set())


def parse_integer(text: str) -> int:
    compact = re.sub(r"[,，￥¥元\s]", "", text)
    if not compact.isdigit():
        raise ValueError(f"价格单元格不是纯整数：{text!r}")
    return int(compact)


def uplift(value: int) -> int:
    return int((Decimal(value) * Decimal("1.1")).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def transform(document_xml: bytes) -> tuple[bytes, dict[str, int]]:
    root = etree.fromstring(document_xml)
    counts: dict[str, int] = {}
    for table in root.xpath(".//w:tbl", namespaces=NS):
        rows = table.xpath("./w:tr", namespaces=NS)
        if not rows:
            continue
        header_cells = rows[0].xpath("./w:tc", namespaces=NS)
        headers = [cell_text(cell) for cell in header_cells]
        kind = classify(headers)
        wanted = target_headers(kind)
        if not wanted:
            continue
        indexes = [index for index, header in enumerate(headers) if normalize(header) in {normalize(value) for value in wanted}]
        for row in rows[1:]:
            cells = row.xpath("./w:tc", namespaces=NS)
            for index in indexes:
                if index >= len(cells):
                    raise ValueError(f"{kind} 表价格列越界")
                texts = cells[index].xpath("./w:p//w:t", namespaces=NS)
                if len(texts) != 1:
                    raise ValueError(f"{kind} 表价格单元格应只有一个文本节点，实际 {len(texts)}")
                old_value = parse_integer(texts[0].text or "")
                texts[0].text = str(uplift(old_value))
                counts[kind] = counts.get(kind, 0) + 1
    if counts != EXPECTED_COUNTS:
        raise ValueError(f"价格单元格计数变化：{counts}，预期 {EXPECTED_COUNTS}")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"), counts


def write_docx(source: Path, output: Path, updated_xml: bytes) -> None:
    with zipfile.ZipFile(source, "r") as source_zip, zipfile.ZipFile(output, "w") as output_zip:
        for item in source_zip.infolist():
            payload = updated_xml if item.filename == "word/document.xml" else source_zip.read(item.filename)
            output_zip.writestr(item, payload)


def solve(input_dir: Path, output_dir: Path) -> Path:
    source = input_dir / INPUT_NAME
    if not source.is_file() or sha256(source) != INPUT_SHA256:
        raise ValueError("价目册输入缺失或 SHA256 不匹配")
    with zipfile.ZipFile(source) as archive:
        updated_xml, counts = transform(archive.read("word/document.xml"))
    if sum(counts.values()) != 1435:
        raise ValueError(f"实际改价数量异常：{sum(counts.values())}")
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_NAME
    write_docx(source, output, updated_xml)
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    print(solve(args.input_dir, args.output_dir))


if __name__ == "__main__":
    main()
