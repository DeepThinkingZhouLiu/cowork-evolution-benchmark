#!/usr/bin/env python3
"""officeval_033 的独立 Harbor Judge：逐个语义价格单元格与来源对照。"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import tempfile
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from lxml import etree

import upstream_verifier
from upstream_verifier import classify_table, norm_text, table_rows


TASK_ID = "cowork-evo/officeval-033-retail-price-catalog-uplift"
OUTPUT_NAME = "家居零售价目册_零售价乘以1.1.docx"
SOURCE_NAME = "家居零售价目册.docx"
RUBRIC_VERSION = "omegause-officeval-033-adaptation-v0.1"
JUDGE_VERSION = "officeval-033-adaptation-v0.1"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
EXPECTED_COUNTS = {"combo": 206, "recommended": 153, "custom": 1057, "conversion": 8, "order": 11}
SPECS = {
    "R001": ("指定 DOCX 可打开且保留完整 Word 包结构", 3, "hurdle"),
    "R002": ("1435 个零售价单元格均按 1.1 四舍五入更新", 8, "hurdle"),
    "R003": ("上游抽样的产品组合、推荐配套与定制价格全部正确", 6, "positive"),
    "R004": ("第 65 页换算提醒与第 66 页订购记录价格正确", 3, "positive"),
    "R005": ("所有非目标文本节点逐字保留", 6, "hurdle"),
    "R006": ("除 document.xml 外所有原始包成员逐字节保留", 4, "positive"),
    "R007": ("没有整页图片化，价格说明文字样式保持", 3, "positive"),
    "R008": ("LibreOffice 实际渲染仍为 66 页且末页价格可提取", 3, "positive"),
    "R009": ("若误改非价格数字、漏改价格或压平页面，扣 10 分", -10, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    description, weight, kind = SPECS[identifier]
    return {"criterion_id": identifier, "description": description, "weight": weight, "type": kind, "passed": bool(passed), "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}


def direct_cells(row):
    return row.findall("w:tc", NS)


def direct_text(cell) -> str:
    return "".join(node.text or "" for node in cell.findall("./w:p//w:t", NS)).strip()


def targets(root):
    found = []
    counts = {}
    wanted_map = {"combo": {"零售价（元）"}, "recommended": {"零售价（元）"}, "custom": {"A1", "A2", "A3", "A4", "A5", "A6", "标准零售价"}, "conversion": {"基础零售价（元）", "核算后零售价（元）"}, "order": {"零售价（元）"}}
    for table in root.findall(".//w:tbl", NS):
        rows = table.findall("w:tr", NS)
        if not rows:
            continue
        headers = [direct_text(cell) for cell in direct_cells(rows[0])]
        kind = classify_table([headers])
        wanted = wanted_map.get(kind, set())
        indexes = [index for index, value in enumerate(headers) if norm_text(value) in {norm_text(item) for item in wanted}]
        for row in rows[1:]:
            cells = direct_cells(row)
            for index in indexes:
                nodes = cells[index].findall("./w:p//w:t", NS)
                if len(nodes) == 1:
                    found.append(nodes[0])
                    counts[kind] = counts.get(kind, 0) + 1
    return found, counts


def parse_int(text: str) -> int:
    return int((text or "").strip().replace(",", "").replace("，", "").replace("元", "").replace("￥", "").replace("¥", ""))


def expected(value: int) -> int:
    return int((Decimal(value) * Decimal("1.1")).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def compare_xml(source_xml: bytes, output_xml: bytes) -> tuple[bool, bool, str]:
    source_root = etree.fromstring(source_xml)
    output_root = etree.fromstring(output_xml)
    source_nodes = source_root.findall(".//w:t", NS)
    output_nodes = output_root.findall(".//w:t", NS)
    source_targets, source_counts = targets(source_root)
    output_targets, output_counts = targets(output_root)
    target_ids = {id(node) for node in source_targets}
    target_positions = {index for index, node in enumerate(source_nodes) if id(node) in target_ids}
    if len(source_nodes) != len(output_nodes) or source_counts != EXPECTED_COUNTS or output_counts != EXPECTED_COUNTS:
        return False, False, f"文本节点={len(source_nodes)}/{len(output_nodes)}，目标计数={source_counts}/{output_counts}"
    wrong_prices = []
    wrong_other = []
    for index, (before, after) in enumerate(zip(source_nodes, output_nodes)):
        if index in target_positions:
            try:
                if parse_int(after.text or "") != expected(parse_int(before.text or "")):
                    wrong_prices.append((index, before.text, after.text))
            except Exception:
                wrong_prices.append((index, before.text, after.text))
        elif (before.text or "") != (after.text or ""):
            wrong_other.append((index, before.text, after.text))
    prices_ok = not wrong_prices and len(target_positions) == 1435
    other_ok = not wrong_other
    return prices_ok, other_ok, f"目标数={len(target_positions)}，错价={wrong_prices[:3]}，误改非目标={wrong_other[:3]}"


def package_preserved(source: Path, output: Path) -> tuple[bool, str]:
    with zipfile.ZipFile(source) as src, zipfile.ZipFile(output) as out:
        src_names = set(src.namelist())
        out_names = set(out.namelist())
        changed = [name for name in sorted(src_names & out_names) if name != "word/document.xml" and hashlib.sha256(src.read(name)).digest() != hashlib.sha256(out.read(name)).digest()]
    ok = src_names == out_names and not changed
    return ok, f"成员集合一致={src_names == out_names}，非正文被修改={changed[:5]}"


def render_check(path: Path) -> tuple[bool, str]:
    if not shutil.which("libreoffice") or not shutil.which("pdfinfo") or not shutil.which("pdftotext"):
        return False, "缺少渲染依赖"
    with tempfile.TemporaryDirectory(prefix="officeval033-render-") as tmp:
        proc = subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, str(path)], capture_output=True, text=True, timeout=600)
        pdf = Path(tmp) / f"{path.stem}.pdf"
        if proc.returncode != 0 or not pdf.is_file():
            return False, f"渲染失败：{proc.stderr[-300:]}"
        info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, timeout=30).stdout
        pages = next((int(line.split(":", 1)[1]) for line in info.splitlines() if line.startswith("Pages:")), 0)
        text = subprocess.run(["pdftotext", "-f", "66", "-l", "66", str(pdf), "-"], capture_output=True, text=True, timeout=60).stdout
        ok = pages == 66 and "40943" in text and "9920" in text
        return ok, f"页数={pages}，末页抽样价格可提取={ok}"


def evaluate(artifact_dir: Path, source_dir: Path) -> dict:
    output = artifact_dir / OUTPUT_NAME
    source = source_dir / SOURCE_NAME
    results = []
    try:
        with zipfile.ZipFile(output) as out_zip, zipfile.ZipFile(source) as src_zip:
            output_xml = out_zip.read("word/document.xml")
            source_xml = src_zip.read("word/document.xml")
        results.append(criterion("R001", True, f"输出大小={output.stat().st_size}B"))
    except Exception as exc:
        results.append(criterion("R001", False, f"缺失或损坏：{exc}"))
        return build_result(results, output)

    prices_ok, other_ok, compare_evidence = compare_xml(source_xml, output_xml)
    results.append(criterion("R002", prices_ok, compare_evidence))
    results.append(criterion("R005", other_ok, compare_evidence))

    upstream = upstream_verifier.evaluate(str(artifact_dir))
    items = upstream.get("dim2_items", [])
    hits = {item.get("rule", ""): bool(item.get("hit")) for item in items}
    first_three = list(hits.values())[:3]
    p65_p66 = list(hits.values())[3:5]
    results.append(criterion("R003", len(first_three) == 3 and all(first_three), f"上游前三组={first_three}，总分={upstream.get('total_score')}"))
    results.append(criterion("R004", len(p65_p66) == 2 and all(p65_p66), f"上游第65/66页={p65_p66}"))

    package_ok, package_evidence = package_preserved(source, output)
    results.append(criterion("R006", package_ok, package_evidence))
    negative_hits = list(hits.values())[5:]
    preservation_ok = len(negative_hits) == 3 and not any(negative_hits)
    results.append(criterion("R007", preservation_ok, f"上游扣分项命中={negative_hits}"))
    render_pass, render_evidence = render_check(output)
    results.append(criterion("R008", render_pass, render_evidence))
    severe = not prices_ok or not other_ok or not preservation_ok
    results.append(criterion("R009", severe, f"漏改/误改/压平={severe}"))
    return build_result(results, output)


def build_result(results: list[dict], output: Path) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])
    total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"]) + sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, earned / total))
    hurdles = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
    if hurdles:
        reward = 0.0
    normalized = [{"criterion_id": item["criterion_id"], "score": (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"], "evidence": item["evidence"], "method": item["method"], "raw": {"description": item["description"], "weight": item["weight"], "type": item["type"], "passed": item["passed"]}} for item in results]
    digest = hashlib.sha256(output.read_bytes() if output.is_file() else b"").hexdigest()
    return {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [digest], "criterion_results": normalized, "reward": round(reward, 6), "passed": reward >= 0.8, "hurdle_failures": hurdles}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--source-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir, args.source_dir)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(rendered + "\n", encoding="utf-8")
    if args.reward_file:
        args.reward_file.parent.mkdir(parents=True, exist_ok=True)
        args.reward_file.write_text(f"{result['reward']:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
