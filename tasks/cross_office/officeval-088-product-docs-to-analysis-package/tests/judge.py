#!/usr/bin/env python3
"""OfficeVal 088 任务专属 v2 原子部分评分；原严格口径保存在 judge_v1.py。"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import re
import sys
from pathlib import Path
from typing import Any, Callable


TASK_ID = "cowork-evo/officeval-088-product-docs-to-analysis-package"
RUBRIC_VERSION = "omegause-officeval-088-adaptation-v0.2"
JUDGE_VERSION = "officeval-088-criterion-partial-v2.0"
OUTPUT_XLSX = "选货汇总表_已整理.xlsx"
OUTPUT_DOCX = "产品分析.docx"
REQUIRED_FILES = (OUTPUT_XLSX, OUTPUT_DOCX)
POSITIVE_WEIGHTS = {
    "R001": 3,
    "R002": 6,
    "R003": 3,
    "R004": 4,
    "R005": 3,
    "R006": 5,
    "R007": 3,
    "R008": 3,
}
PENALTY_WEIGHTS = {"R009": -8, "R010": -6, "R011": -8, "R012": -5}


def _load_v1():
    path = Path(__file__).with_name("judge_v1.py")
    spec = importlib.util.spec_from_file_location("officeval_088_judge_v1", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("无法载入 v1 Judge")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


V1 = _load_v1()


def _ratio(findings: list[tuple[bool, str]]) -> tuple[float, str, int, int]:
    total = len(findings)
    earned = sum(bool(ok) for ok, _ in findings)
    failures = [label for ok, label in findings if not ok]
    evidence = "全部原子检查通过" if not failures else "未通过：" + "；".join(failures[:16])
    if len(failures) > 16:
        evidence += f"；另有 {len(failures) - 16} 项"
    return (earned / total if total else 0.0), evidence, earned, total


def _criterion(
    identifier: str,
    score: float,
    evidence: str,
    earned: float,
    total: float,
) -> dict[str, Any]:
    return {
        "criterion_id": identifier,
        "score": round(max(0.0, min(1.0, score)), 6),
        "evidence": evidence,
        "method": "deterministic",
        "raw": {"earned_units": round(earned, 6), "total_units": round(total, 6)},
    }


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _core(artifact: Path) -> tuple[bool, list[str]]:
    errors: list[str] = []
    for name in REQUIRED_FILES:
        path = artifact / name
        if not path.is_file() or path.stat().st_size == 0:
            errors.append(f"核心交付缺失或为空：{name}")
    if errors:
        return False, errors
    try:
        book = V1.load_workbook(artifact / OUTPUT_XLSX, data_only=False)
        if not book.worksheets:
            raise ValueError("XLSX 无工作表")
        book.close()
    except Exception as error:
        errors.append(f"XLSX 无法解析：{type(error).__name__}: {error}")
    try:
        document = V1.Document(artifact / OUTPUT_DOCX)
        if not document.paragraphs and not document.tables:
            raise ValueError("DOCX 无内容")
    except Exception as error:
        errors.append(f"DOCX 无法解析：{type(error).__name__}: {error}")
    return not errors, errors


def _fatal(artifact: Path, errors: list[str]) -> dict[str, Any]:
    evidence = "；".join(errors)
    rows = [
        _criterion(identifier, 0.0, evidence, 0, 0)
        for identifier in (*POSITIVE_WEIGHTS, *PENALTY_WEIGHTS)
    ]
    digests = [_sha256(artifact / name) for name in REQUIRED_FILES if (artifact / name).is_file()]
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": digests or [hashlib.sha256(b"").hexdigest()],
        "criterion_results": rows,
        "score": 0.0,
        "reward": 0.0,
        "passed": False,
        "hurdle_failures": ["R001"],
        "penalties": [],
    }


def _excel_fact_findings(excel: dict, expectations: dict) -> tuple[list[tuple[bool, str]], int, int]:
    findings: list[tuple[bool, str]] = []
    product_errors = 0
    price_errors = 0
    sheet = excel["sheet"]
    for product in expectations["products"]:
        key = product["key"]
        row = excel["rows"].get(key)
        findings.append((row is not None, f"{key}产品行缺失"))
        if row is None:
            product_errors += 1
            price_errors += 1
            continue
        joined = V1.norm(" ".join(str(sheet.cell(row, column).value or "") for column in range(1, 5)))
        product_ok = True
        for token in [product["name"], product["model"], *product["parameters"]]:
            ok = V1.norm(token) in joined
            findings.append((ok, f"{key}缺少事实 {token}"))
            product_ok &= ok
        digits = re.sub(r"\D", "", V1.norm(sheet.cell(row, 4).value))
        price_ok = str(product["price"]) in digits
        findings.append((price_ok, f"{key}价格错误：{sheet.cell(row, 4).value}"))
        if not product_ok:
            product_errors += 1
        if not price_ok:
            price_errors += 1
    return findings, product_errors, price_errors


def _excel_scene_findings(excel: dict, expectations: dict) -> list[tuple[bool, str]]:
    findings: list[tuple[bool, str]] = []
    sheet = excel["sheet"]
    for product in expectations["products"]:
        key = product["key"]
        row = excel["rows"].get(key)
        if row is None:
            findings.append((False, f"{key}缺行，无法核验场景"))
            continue
        scene = V1.norm(sheet.cell(row, 3).value)
        for token in product["scenes"]:
            findings.append((V1.norm(token) in scene, f"{key}缺少场景 {token}"))
        findings.extend(
            [
                (bool(sheet.cell(row, 2).alignment.wrap_text), f"{key}核心参数未自动换行"),
                (bool(sheet.cell(row, 3).alignment.wrap_text), f"{key}适用场景未自动换行"),
                ((sheet.row_dimensions[row].height or 15) >= 170, f"{key}行高不足"),
            ]
        )
    return findings


def _excel_image_findings(
    excel: dict,
    expectations: dict,
    crops: dict,
) -> tuple[list[tuple[bool, str]], int]:
    del crops  # 源图相似度已由 v1 excel_state 逐图计算，避免重读已关闭图像流。
    findings: list[tuple[bool, str]] = []
    sheet = excel["sheet"]
    issues = [str(item) for item in excel["image_issues"]]
    details = [str(item) for item in excel["image_details"]]
    bad_products = 0
    for product in expectations["products"]:
        key = product["key"]
        product_issues = [item for item in issues if item.startswith(f"{key}:")]
        seen = any(item.startswith(f"{key}:") for item in details)
        ok = seen and not product_issues
        findings.append((ok, f"{key}展示图未正确锨定/裁剪/缩放：{product_issues or '未找到图片'}"))
        if not ok:
            bad_products += 1
    findings.append((len(sheet._images) == 5, f"Excel 图片总数不是 5：{len(sheet._images)}"))
    return findings, bad_products


def _excel_style_findings(excel: dict) -> list[tuple[bool, str]]:
    sheet = excel["sheet"]
    headers = [V1.norm(sheet.cell(1, column).value) for column in range(1, 6)]
    expected_headers = ["产品名称", "核心参数", "适用场景", "指导价格", "展示图片"]
    expected_widths = [23.42, 32.81, 10.82, 12.36, 12.08]
    findings = [(headers == expected_headers, f"表头错误：{headers}")]
    for index, expected in enumerate(expected_widths, 1):
        column = chr(64 + index)
        actual = sheet.column_dimensions[column].width or 8.43
        findings.append((abs(actual - expected) <= 0.1, f"{column}列宽={actual:.2f}"))
    findings.append((sheet.page_setup.orientation == "landscape", "Excel 未设为横向打印"))
    findings.append((all(sheet.cell(1, column).fill.fill_type for column in range(1, 6)), "表头底色不完整"))
    return findings


def _word_fact_findings(word: dict, excel: dict, expectations: dict) -> tuple[list[tuple[bool, str]], int, int]:
    findings: list[tuple[bool, str]] = []
    expected_order = expectations["top_three"]
    actual_order = word["actual_order"]
    wrong_positions = 0
    for index, expected in enumerate(expected_order):
        actual = actual_order[index] if index < len(actual_order) else None
        ok = actual == expected
        findings.append((ok, f"Word 第 {index + 1} 个产品应为 {expected}，实际 {actual}"))
        wrong_positions += int(not ok)
    blocks = {key: text for key, text, _ in word["blocks"]}
    product_by_key = {product["key"]: product for product in expectations["products"]}
    inconsistent_products = 0
    for key in expected_order:
        product = product_by_key[key]
        text = blocks.get(key, "")
        row = excel["rows"].get(key)
        excel_text = V1.norm(" ".join(str(excel["sheet"].cell(row, column).value or "") for column in range(1, 5))) if row else ""
        product_ok = bool(text)
        findings.append((bool(text), f"Word 缺少 {key} 产品块"))
        for token in [product["name"], product["model"], str(product["price"]), *product["parameters"], *product["scenes"]]:
            ok = V1.norm(token) in text and V1.norm(token) in excel_text
            findings.append((ok, f"{key} 在 Word/Excel 中未一致呈现 {token}"))
            product_ok &= ok
        inconsistent_products += int(not product_ok)
    return findings, wrong_positions, inconsistent_products


def _word_image_findings(word: dict, expectations: dict, crops: dict) -> tuple[list[tuple[bool, str]], int]:
    findings: list[tuple[bool, str]] = [(len(word["blocks"]) == 3, f"Word 产品块数不是 3：{len(word['blocks'])}")]
    expected_order = set(expectations["top_three"])
    product_by_key = {product["key"]: product for product in expectations["products"]}
    bad_products = 0
    for key, _, table in word["blocks"]:
        if key not in expected_order:
            continue
        picture = V1.word_image_for_table(word["document"], table)
        findings.append((picture is not None, f"Word {key}产品块缺图"))
        if picture is None:
            bad_products += 1
            continue
        raw, ratio, height_cm = picture
        product = product_by_key[key]
        distance = V1.image_distance(raw, crops[key])
        cant_split = table.rows[0]._tr.trPr is not None and table.rows[0]._tr.trPr.find(V1.qn("w:cantSplit")) is not None
        checks = [
            (distance <= 35, f"Word {key}图片不是指定源裁图"),
            (abs(ratio - product["ratio"]) <= 0.02, f"Word {key}图片比例错误"),
            (4.9 <= height_cm <= 5.1, f"Word {key}图片高度不是 5cm"),
            (cant_split, f"Word {key}产品块可跨页拆分"),
        ]
        findings.extend(checks)
        if not all(ok for ok, _ in checks):
            bad_products += 1
    missing = expected_order - {key for key, _, _ in word["blocks"]}
    for key in sorted(missing):
        findings.append((False, f"Word 缺少 {key}产品块，无法核验图片"))
        bad_products += 1
    return findings, bad_products


def evaluate(artifact_dir: Path, *, run_render: bool = True) -> dict[str, Any]:
    artifact = Path(artifact_dir)
    core_ok, core_errors = _core(artifact)
    if not core_ok:
        return _fatal(artifact, core_errors)

    task_root = Path(__file__).resolve().parents[1]
    input_dir = Path("/app/input_files") if Path("/app/input_files").is_dir() else task_root / "data/input_files"
    expectations = json.loads((Path(__file__).with_name("expectations.json")).read_text(encoding="utf-8"))
    crops = V1.expected_crops(input_dir)
    excel = V1.excel_state(artifact / OUTPUT_XLSX, expectations, crops)
    word = V1.word_state(artifact / OUTPUT_DOCX, expectations, crops, excel)
    results: dict[str, tuple[float, str, int, int]] = {
        "R001": (1.0, "两个核心 OOXML 交付非空且可解析", 2, 2)
    }

    facts, product_errors, price_errors = _excel_fact_findings(excel, expectations)
    results["R002"] = _ratio(facts)
    results["R003"] = _ratio(_excel_scene_findings(excel, expectations))
    excel_images, bad_excel_images = _excel_image_findings(excel, expectations, crops)
    results["R004"] = _ratio(excel_images)
    results["R005"] = _ratio(_excel_style_findings(excel))
    word_facts, wrong_order_positions, inconsistent_products = _word_fact_findings(word, excel, expectations)
    results["R006"] = _ratio(word_facts)
    word_images, bad_word_images = _word_image_findings(word, expectations, crops)
    results["R007"] = _ratio(word_images)

    if run_render:
        rendered_ok, render_evidence = V1.render_package(artifact / OUTPUT_XLSX, artifact / OUTPUT_DOCX)
    else:
        rendered_ok, render_evidence = True, "分层负例快速检查沿用已验证渲染基线"
    r8_findings = [
        (bool(word["header_ok"]), "Word 页眉或下划线不符合要求"),
        (bool(word["footer_ok"]), "Word 居中动态页码不符合要求"),
        (rendered_ok, f"实际渲染未通过：{render_evidence}"),
    ]
    results["R008"] = _ratio(r8_findings)

    product_count = max(1, len(expectations["products"]))
    top_count = max(1, len(expectations["top_three"]))
    p9_score = max(price_errors / product_count, product_errors / product_count)
    results["R009"] = (p9_score, f"价格错误={price_errors}/5，产品事实错误={product_errors}/5", max(price_errors, product_errors), product_count)
    p10_score = min(1.0, bad_excel_images / product_count)
    results["R010"] = (p10_score, f"Excel 错误产品图={bad_excel_images}/5", bad_excel_images, product_count)
    p11_score = max(wrong_order_positions / top_count, inconsistent_products / top_count)
    results["R011"] = (p11_score, f"顺序错位={wrong_order_positions}/3，Word/Excel 不一致产品={inconsistent_products}/3", max(wrong_order_positions, inconsistent_products), top_count)
    p12_failures = sum(not ok for ok, _ in r8_findings) + min(bad_word_images, top_count)
    p12_total = len(r8_findings) + top_count
    p12_score = p12_failures / p12_total
    results["R012"] = (p12_score, f"页眉/页码/渲染与 Word 图片版式失败={p12_failures}/{p12_total}", p12_failures, p12_total)

    excel["workbook"].close()
    positive_total = sum(POSITIVE_WEIGHTS.values())
    positive = sum(POSITIVE_WEIGHTS[key] * results[key][0] for key in POSITIVE_WEIGHTS) / positive_total
    penalties = sum((-PENALTY_WEIGHTS[key]) * results[key][0] for key in PENALTY_WEIGHTS) / positive_total
    reward = round(max(0.0, min(1.0, positive - penalties)), 6)
    criterion_results = [
        _criterion(key, *results[key]) for key in (*POSITIVE_WEIGHTS, *PENALTY_WEIGHTS)
    ]
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [_sha256(artifact / name) for name in REQUIRED_FILES],
        "criterion_results": criterion_results,
        "score": reward,
        "reward": reward,
        "passed": reward >= 0.75,
        "hurdle_failures": [],
        "penalties": [key for key in PENALTY_WEIGHTS if results[key][0] > 0],
        "render_evidence": render_evidence,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    parser.add_argument("--skip-render", action="store_true")
    parser.add_argument("--scoring-version", choices=["v1", "v2"], default="v2")
    arguments = parser.parse_args()
    if arguments.scoring_version == "v1":
        result = V1.evaluate(arguments.artifact_dir)
        result["score"] = result.get("reward", 0.0)
        result["scoring_version"] = "v1"
    else:
        result = evaluate(arguments.artifact_dir, run_render=not arguments.skip_render)
        result["scoring_version"] = "v2"
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if arguments.result:
        arguments.result.parent.mkdir(parents=True, exist_ok=True)
        arguments.result.write_text(rendered + "\n", encoding="utf-8")
    if arguments.reward_file:
        arguments.reward_file.parent.mkdir(parents=True, exist_ok=True)
        arguments.reward_file.write_text(f"{float(result.get('reward', 0.0)):.6f}\n", encoding="utf-8")


if __name__ == "__main__":
    main()
