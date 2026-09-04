#!/usr/bin/env python3
"""退款根因六文件包的 v2 连续评分 Judge。

v1 被原样封存在 ``judge_v1.py``。默认 v2 只在核心文件缺失、为空、
损坏或拿冻结输入冒充交付时硬归零；其余错误按本题退款、商品、汇总、
报告和邮件的原子事实连续计分。
"""

from __future__ import annotations

import argparse
import ast
import hashlib
import importlib.util
import json
import os
import re
import tempfile
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook


TASK_ID = "cowork-evo/cowork-bench-refund-root-cause-package"
RUBRIC_VERSION = "cowork-bench-refund-root-cause-package-v0.2-continuous"
JUDGE_VERSION = "cowork-bench-refund-root-cause-package-v0.2-continuous"
PASS_THRESHOLD = 0.8
OUTPUT_NAMES = [
    "refund_root_cause_processor.py",
    "analysis_results.json",
    "Refund_Analysis.xlsx",
    "Root_Cause_Report.docx",
    "quality_team_email.txt",
    "supplier_relations_email.txt",
]
DETAIL_HEADERS = [
    "Refund_ID", "Order_ID", "Product_Name", "Refund_Amount",
    "Severity", "Refund_Reason", "Date",
]
IMPACT_HEADERS = [
    "Product_Name", "Refund_Count", "Total_Refund_Amount", "Avg_Review_Rating",
    "Review_Count", "Issue_Type", "Requires_Investigation",
]
SUMMARY_KEYS = [
    "Total_Refunds", "Total_Refund_Amount", "Critical_Count", "Major_Count",
    "Minor_Count", "Products_Affected", "Quality_Issues", "Service_Issues",
    "Products_Requiring_Investigation",
]
WORD_SECTIONS = [
    "Executive Summary", "Refund Trend Analysis", "Product-Level Findings",
    "Root Cause Classification", "Recommendations", "Action Plan",
]
SPECS = {
    "R001": ("六个核心交付存在、非空、可解析且不是冻结输入副本", 5, "hurdle"),
    "R002": ("处理器在原始输入上按原子事实正确回放", 8, "positive"),
    "R003": ("处理器在金额与评分变异输入上按原子事实正确跟随", 8, "positive"),
    "R004": ("analysis_results.json 全部任务专属原子字段正确", 12, "positive"),
    "R005": ("Excel Refund Details 的结构、顺序和 13×7 字段正确", 12, "positive"),
    "R006": ("Excel Product Impact 的结构、排序和 24×7 字段正确", 15, "positive"),
    "R007": ("Excel Summary 的结构、顺序和九项值正确", 6, "positive"),
    "R008": ("Word 六章节、汇总、原因、根因和调查商品事实正确", 12, "positive"),
    "R009": ("两封离线邮件的路由、主题、汇总和调查商品事实正确", 8, "positive"),
    "R010": ("JSON、Excel、Word 与邮件中的核心事实相互一致", 5, "positive"),
    "R011": ("Excel 与 Word 保留专业、原生、可编辑结构", 5, "positive"),
    "R012": ("Excel 与 Word 可实际渲染且核心文本可提取", 4, "positive"),
    "R013": ("出现输入不支持的业务数字或断言", -10, "penalty"),
    "R014": ("输出目录出现六个指定交付之外的文件", -2, "penalty"),
}


def _load_legacy():
    path = Path(__file__).with_name("judge_v1.py")
    spec = importlib.util.spec_from_file_location("refund_judge_v1", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载 v1 Judge：{path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


LEGACY = _load_legacy()


def same(actual: Any, expected: Any) -> bool:
    return bool(LEGACY.payload_equal(actual, expected))


def expected_units(expected: Any) -> int:
    if isinstance(expected, dict):
        return 1 + sum(expected_units(value) for value in expected.values())
    if isinstance(expected, list):
        return 1 + sum(expected_units(value) for value in expected)
    return 1


def atomic_units(actual: Any, expected: Any) -> tuple[int, int]:
    """递归计数结构单元与叶子事实，不把一处错误扩成整个载体 0 分。"""
    if isinstance(expected, dict):
        total = 1
        matched = int(isinstance(actual, dict) and set(actual) == set(expected))
        for key, value in expected.items():
            if isinstance(actual, dict) and key in actual:
                child_match, child_total = atomic_units(actual[key], value)
            else:
                child_match, child_total = 0, expected_units(value)
            matched += child_match
            total += child_total
        return matched, total
    if isinstance(expected, list):
        total = 1
        matched = int(isinstance(actual, list) and len(actual) == len(expected))
        for index, value in enumerate(expected):
            if isinstance(actual, list) and index < len(actual):
                child_match, child_total = atomic_units(actual[index], value)
            else:
                child_match, child_total = 0, expected_units(value)
            matched += child_match
            total += child_total
        return matched, total
    return int(same(actual, expected)), 1


def metamorphic_units(actual: Any, expected: dict[str, Any]) -> tuple[int, int]:
    """只评本题变异真正影响的 15 个事实，避免未变字段稀释回放错误。"""
    if not isinstance(actual, dict):
        return 0, 15

    def by_key(payload: Any, section: str, key: str, value: Any) -> dict[str, Any]:
        rows = payload.get(section, []) if isinstance(payload, dict) else []
        return next((row for row in rows if isinstance(row, dict) and row.get(key) == value), {})

    expected_detail = by_key(expected, "refund_details", "Refund_ID", 4)
    actual_detail = by_key(actual, "refund_details", "Refund_ID", 4)
    product_name = expected["product_impact"][0]["Product_Name"]
    expected_product = by_key(expected, "product_impact", "Product_Name", product_name)
    actual_product = by_key(actual, "product_impact", "Product_Name", product_name)
    expected_investigation = by_key(expected, "investigation_products", "Product_Name", product_name)
    actual_investigation = by_key(actual, "investigation_products", "Product_Name", product_name)
    pairs: list[tuple[Any, Any]] = []
    for field in ("Refund_Amount", "Severity"):
        pairs.append((actual_detail.get(field), expected_detail.get(field)))
    for field in ("Total_Refund_Amount", "Avg_Review_Rating", "Issue_Type"):
        pairs.append((actual_product.get(field), expected_product.get(field)))
    for field in ("Total_Refund_Amount", "Avg_Review_Rating", "Issue_Type"):
        pairs.append((actual_investigation.get(field), expected_investigation.get(field)))
    for field in (
        "Total_Refund_Amount", "Critical_Count", "Minor_Count",
        "Quality_Issues", "Service_Issues",
    ):
        actual_summary = actual.get("summary", {})
        pairs.append((actual_summary.get(field) if isinstance(actual_summary, dict) else None, expected["summary"][field]))
    # 退款 4 在两个完整列表中的存在，以及六个顶层区域仍可用，也属于变异回放结构。
    pairs.extend([
        (bool(actual_detail), True),
        (bool(actual_product), True),
    ])
    return sum(int(same(left, right)) for left, right in pairs), len(pairs)


def fraction(matched: int, total: int) -> float:
    return matched / total if total else 0.0


def positive(identifier: str, score: float, evidence: str) -> dict[str, Any]:
    description, weight, kind = SPECS[identifier]
    bounded = max(0.0, min(1.0, float(score)))
    return {
        "criterion_id": identifier,
        "score": round(bounded, 6),
        "evidence": evidence,
        "method": "deterministic",
        "raw": {
            "description": description,
            "weight": weight,
            "type": kind,
            "matched_fraction": bounded,
        },
    }


def penalty(identifier: str, trigger_fraction: float, evidence: str) -> dict[str, Any]:
    description, weight, kind = SPECS[identifier]
    trigger = max(0.0, min(1.0, float(trigger_fraction)))
    return {
        "criterion_id": identifier,
        "score": round(1.0 - trigger, 6),
        "evidence": evidence,
        "method": "deterministic",
        "raw": {
            "description": description,
            "weight": weight,
            "type": kind,
            "trigger_fraction": trigger,
        },
    }


def normalize_cell(value: Any) -> Any:
    return LEGACY.normalize_cell(value)


def sheet_rows(sheet, headers: list[str]) -> tuple[list[Any], list[dict[str, Any]]]:
    actual_headers = [sheet.cell(1, column).value for column in range(1, len(headers) + 1)]
    rows: list[dict[str, Any]] = []
    for row_number in range(2, sheet.max_row + 1):
        values = [
            normalize_cell(sheet.cell(row_number, column).value)
            for column in range(1, len(headers) + 1)
        ]
        if all(value in (None, "") for value in values):
            continue
        rows.append(dict(zip(headers, values)))
    return actual_headers, rows


def table_atomic_score(
    workbook,
    sheet_name: str,
    headers: list[str],
    expected_rows: list[dict[str, Any]],
    key: str,
) -> tuple[float, str, list[dict[str, Any]]]:
    # 单元：工作表 1、表头每列、行数 1、排序 1、每个预期字段。
    total = 1 + len(headers) + 2 + len(expected_rows) * len(headers)
    if sheet_name not in workbook.sheetnames:
        return 0.0, f"缺工作表 {sheet_name}；0/{total}", []
    sheet = workbook[sheet_name]
    actual_headers, actual_rows = sheet_rows(sheet, headers)
    matched = 1
    matched += sum(int(same(actual, expected)) for actual, expected in zip(actual_headers, headers))
    matched += int(len(actual_rows) == len(expected_rows))
    matched += int(
        [row.get(key) for row in actual_rows]
        == [row.get(key) for row in expected_rows]
    )
    actual_by_key = {row.get(key): row for row in actual_rows}
    for expected_row in expected_rows:
        actual_row = actual_by_key.get(expected_row[key], {})
        matched += sum(
            int(field in actual_row and same(actual_row[field], expected_row[field]))
            for field in headers
        )
    return fraction(matched, total), f"原子匹配={matched}/{total}，实际行={len(actual_rows)}", actual_rows


def summary_atomic_score(workbook, expected: dict[str, Any]) -> tuple[float, str, dict[str, Any]]:
    total = 1 + 2 + 2 + len(SUMMARY_KEYS)
    if "Summary" not in workbook.sheetnames:
        return 0.0, f"缺工作表 Summary；0/{total}", {}
    sheet = workbook["Summary"]
    headers = [sheet.cell(1, column).value for column in (1, 2)]
    rows = []
    for row_number in range(2, sheet.max_row + 1):
        key = sheet.cell(row_number, 1).value
        value = normalize_cell(sheet.cell(row_number, 2).value)
        if key not in (None, ""):
            rows.append((str(key), value))
    actual = dict(rows)
    matched = 1
    matched += sum(int(same(actual_value, wanted)) for actual_value, wanted in zip(headers, ["Metric", "Value"]))
    matched += int(len(rows) == len(SUMMARY_KEYS))
    matched += int([key for key, _ in rows] == SUMMARY_KEYS)
    matched += sum(int(key in actual and same(actual[key], expected[key])) for key in SUMMARY_KEYS)
    return fraction(matched, total), f"原子匹配={matched}/{total}，实际指标={len(rows)}", actual


def word_atomic_score(text: str, headings: list[str], truth: dict[str, Any]) -> tuple[float, str]:
    units: list[tuple[str, bool]] = []
    for section in WORD_SECTIONS:
        units.append((f"章节:{section}", any(section.lower() == heading.lower() for heading in headings)))
    summary = truth["summary"]
    units.extend([
        ("总金额", LEGACY.has_money(text, summary["Total_Refund_Amount"])),
        ("退款总数", LEGACY.label_count(text, "refund", summary["Total_Refunds"], 90)),
        ("Critical", LEGACY.label_count(text, "Critical", summary["Critical_Count"], 90)),
        ("Major", LEGACY.label_count(text, "Major", summary["Major_Count"], 90)),
        ("Minor", LEGACY.label_count(text, "Minor", summary["Minor_Count"], 90)),
    ])
    for row in truth["refund_reason_breakdown"]:
        units.append((f"原因:{row['Refund_Reason']}", LEGACY.label_count(text, row["Refund_Reason"], row["Refund_Count"], 45)))
    for row in truth["root_cause_breakdown"]:
        units.append((f"根因:{row['Root_Cause']}", LEGACY.label_count(text, row["Root_Cause"], row["Refund_Count"], 45)))
    for row in truth["investigation_products"]:
        position = text.find(row["Product_Name"])
        block = text[position: position + len(row["Product_Name"]) + 220] if position >= 0 else ""
        units.extend([
            (f"商品:{row['Product_Name'][:24]}", position >= 0),
            (f"商品退款数:{row['Product_Name'][:24]}", bool(block) and number_near_product(block, row["Product_Name"], row["Refund_Count"])),
            (f"商品金额:{row['Product_Name'][:24]}", bool(block) and LEGACY.has_money(block, row["Total_Refund_Amount"])),
        ])
    matched = sum(int(passed) for _, passed in units)
    failed = [label for label, passed in units if not passed]
    return fraction(matched, len(units)), f"原子匹配={matched}/{len(units)}，缺失={failed}"


def _subject(text: str) -> str:
    match = re.search(r"^subject:\s*(.+)$", text, re.IGNORECASE | re.MULTILINE)
    return match.group(1) if match else ""


def number_near_product(block: str, product_name: str, value: Any) -> bool:
    """Word 原生表格提取后行标题不重复，用商品名后的首段判断退款数。"""
    tail = block[len(product_name): len(product_name) + 80]
    return re.search(rf"(?<![0-9.]){re.escape(str(value))}(?![0-9.])", tail) is not None


def email_atomic_score(quality: str, supplier: str, truth: dict[str, Any]) -> tuple[float, str]:
    summary = truth["summary"]
    q_subject = _subject(quality).lower()
    s_subject = _subject(supplier).lower()
    units: list[tuple[str, bool]] = [
        ("质量收件人", "to: quality_team@company.com" in quality.lower()),
        ("质量主题 Refund", "refund" in q_subject),
        ("质量主题 Analysis", "analysis" in q_subject),
        ("质量退款总数", LEGACY.label_count(quality, "refund", summary["Total_Refunds"], 90)),
        ("质量总金额", LEGACY.has_money(quality, summary["Total_Refund_Amount"])),
        ("质量 Critical", LEGACY.label_count(quality, "Critical", summary["Critical_Count"], 45)),
        ("质量 Major", LEGACY.label_count(quality, "Major", summary["Major_Count"], 45)),
        ("质量 Minor", LEGACY.label_count(quality, "Minor", summary["Minor_Count"], 45)),
        ("质量调查商品数", LEGACY.label_count(quality, "investigation", summary["Products_Requiring_Investigation"], 90)),
        ("供应商收件人", "to: supplier_relations@company.com" in supplier.lower()),
        ("供应商主题 Supplier", "supplier" in s_subject),
        ("供应商主题 Investigation", "investigation" in s_subject),
    ]
    for row in truth["refund_reason_breakdown"]:
        units.append((f"质量原因:{row['Refund_Reason']}", LEGACY.label_count(quality, row["Refund_Reason"], row["Refund_Count"], 45)))
    for row in truth["investigation_products"]:
        position = supplier.find(row["Product_Name"])
        block = supplier[position: position + len(row["Product_Name"]) + 180] if position >= 0 else ""
        units.extend([
            (f"供应商商品:{row['Product_Name'][:24]}", position >= 0),
            (f"供应商退款数:{row['Product_Name'][:24]}", bool(block) and LEGACY.label_count(block, "refunds", row["Refund_Count"], 40)),
            (f"供应商金额:{row['Product_Name'][:24]}", bool(block) and LEGACY.has_money(block, row["Total_Refund_Amount"])),
        ])
    matched = sum(int(passed) for _, passed in units)
    failed = [label for label, passed in units if not passed]
    return fraction(matched, len(units)), f"原子匹配={matched}/{len(units)}，缺失={failed}"


def cross_artifact_score(
    delivered_json: dict[str, Any],
    excel_summary: dict[str, Any],
    report_text: str,
    quality_text: str,
    supplier_text: str,
) -> tuple[float, str]:
    units: list[tuple[str, bool]] = []
    json_summary = delivered_json.get("summary", {}) if isinstance(delivered_json, dict) else {}
    for key in SUMMARY_KEYS:
        units.append((f"JSON↔Excel:{key}", key in json_summary and key in excel_summary and same(json_summary[key], excel_summary[key])))
    textual_keys = ["Total_Refunds", "Total_Refund_Amount", "Critical_Count", "Major_Count", "Minor_Count"]
    labels = {
        "Total_Refunds": "refund", "Critical_Count": "Critical",
        "Major_Count": "Major", "Minor_Count": "Minor",
    }
    for key in textual_keys:
        if key not in json_summary:
            units.extend([(f"JSON↔Word:{key}", False), (f"JSON↔质量邮件:{key}", False)])
            continue
        value = json_summary[key]
        if key == "Total_Refund_Amount":
            units.extend([
                (f"JSON↔Word:{key}", LEGACY.has_money(report_text, value)),
                (f"JSON↔质量邮件:{key}", LEGACY.has_money(quality_text, value)),
            ])
        else:
            units.extend([
                (f"JSON↔Word:{key}", LEGACY.label_count(report_text, labels[key], value, 90)),
                (f"JSON↔质量邮件:{key}", LEGACY.label_count(quality_text, labels[key], value, 90)),
            ])
    investigations = delivered_json.get("investigation_products", []) if isinstance(delivered_json, dict) else []
    for row in investigations[:4] if isinstance(investigations, list) else []:
        name = str(row.get("Product_Name", ""))
        count = row.get("Refund_Count")
        amount = row.get("Total_Refund_Amount")
        for carrier, text in [("Word", report_text), ("供应商邮件", supplier_text)]:
            position = text.find(name) if name else -1
            block = text[position: position + len(name) + 220] if position >= 0 else ""
            units.append((f"JSON↔{carrier}商品:{name[:18]}", position >= 0))
            count_ok = (
                number_near_product(block, name, count)
                if carrier == "Word"
                else LEGACY.label_count(block, "refund", count, 50)
            )
            units.append((f"JSON↔{carrier}退款数:{name[:18]}", bool(block) and count_ok))
            units.append((f"JSON↔{carrier}金额:{name[:18]}", bool(block) and LEGACY.has_money(block, amount)))
    matched = sum(int(passed) for _, passed in units)
    failed = [label for label, passed in units if not passed]
    return fraction(matched, len(units)), f"一致单元={matched}/{len(units)}，冲突={failed}"


def native_structure_score(workbook, report_tables: int, report_media: int, headings: list[str]) -> tuple[float, str]:
    freezes = {sheet.title: str(sheet.freeze_panes or "") for sheet in workbook.worksheets}
    units = {
        "三张且仅三张工作表": workbook.sheetnames == ["Refund Details", "Product Impact", "Summary"],
        "三张原生表": sum(len(sheet.tables) for sheet in workbook.worksheets) >= 3,
        "三张冻结首行": freezes == {"Refund Details": "A2", "Product Impact": "A2", "Summary": "A2"},
        "无 Excel 图片压平": sum(len(sheet._images) for sheet in workbook.worksheets) == 0,
        "表头加粗": all(all(cell.font.bold for cell in sheet[1] if cell.value is not None) for sheet in workbook.worksheets),
        "合理列宽": all(any((dimension.width or 0) > 12 for dimension in sheet.column_dimensions.values()) for sheet in workbook.worksheets),
        "Word 至少五张原生表": report_tables >= 5,
        "Word 无图片压平": report_media == 0,
        "Word 原生标题": len(headings) >= 7,
    }
    matched = sum(int(value) for value in units.values())
    return fraction(matched, len(units)), f"结构单元={matched}/{len(units)}，详情={units}"


def input_digests(input_dir: Path) -> set[str]:
    return {
        hashlib.sha256(path.read_bytes()).hexdigest()
        for path in input_dir.iterdir()
        if path.is_file()
    }


def build_result(criteria: list[dict[str, Any]], paths: dict[str, Path], hard_zero_reasons: list[str]) -> dict[str, Any]:
    criteria.sort(key=lambda item: item["criterion_id"])
    positive_total = sum(weight for _, weight, _ in SPECS.values() if weight > 0)
    earned = 0.0
    penalty_points = 0.0
    for item in criteria:
        weight = SPECS[item["criterion_id"]][1]
        if weight > 0:
            earned += weight * item["score"]
        else:
            penalty_points += weight * item["raw"].get("trigger_fraction", 0.0)
    reward = max(0.0, min(1.0, (earned + penalty_points) / positive_total))
    if hard_zero_reasons:
        reward = 0.0
    digests = [
        hashlib.sha256(path.read_bytes()).hexdigest()
        for path in paths.values()
        if path.is_file()
    ] or [hashlib.sha256(b"").hexdigest()]
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "scoring_version": "v2",
        "artifact_sha256": digests,
        "criterion_results": criteria,
        "reward": round(reward, 6),
        "passed": reward >= PASS_THRESHOLD,
        "hurdle_failures": ["R001"] if hard_zero_reasons else [],
        "hard_zero_reasons": hard_zero_reasons,
    }


def fatal_result(paths: dict[str, Path], reasons: list[str], extras: list[str]) -> dict[str, Any]:
    criteria = [positive("R001", 0.0, f"硬归零：{reasons}")]
    for identifier in [f"R{number:03d}" for number in range(2, 13)]:
        criteria.append(positive(identifier, 0.0, "核心完整性失败，未继续给分"))
    criteria.append(penalty("R013", 0.0, "核心完整性失败，未扫描断言"))
    criteria.append(penalty("R014", 1.0 if extras else 0.0, f"额外文件={extras}"))
    return build_result(criteria, paths, reasons)


def evaluate_v2(artifact_dir: Path) -> dict[str, Any]:
    task_root = Path(__file__).resolve().parents[1]
    input_dir = Path("/app/input_files") if Path("/app/input_files").is_dir() else task_root / "data" / "input_files"
    paths = {name: artifact_dir / name for name in OUTPUT_NAMES}
    actual_files = sorted(path.name for path in artifact_dir.iterdir() if path.is_file()) if artifact_dir.is_dir() else []
    ignored_authoring = {"solve.py", "solve.sh"} if artifact_dir.resolve() == (task_root / "solution").resolve() else set()
    extras = sorted(set(actual_files).difference(OUTPUT_NAMES).difference(ignored_authoring))

    reasons: list[str] = []
    for name, path in paths.items():
        if not path.is_file():
            reasons.append(f"缺失 {name}")
        elif path.stat().st_size == 0:
            reasons.append(f"空文件 {name}")
    if reasons:
        return fatal_result(paths, reasons, extras)

    frozen_digests = input_digests(input_dir)
    for name, path in paths.items():
        if hashlib.sha256(path.read_bytes()).hexdigest() in frozen_digests:
            reasons.append(f"{name} 与冻结输入逐字节相同")

    try:
        ast.parse(paths["refund_root_cause_processor.py"].read_text(encoding="utf-8"))
    except Exception as error:
        reasons.append(f"处理脚本不可解析：{error}")
    try:
        delivered_json = json.loads(paths["analysis_results.json"].read_text(encoding="utf-8"))
        if not isinstance(delivered_json, dict):
            raise TypeError("顶层不是对象")
    except Exception as error:
        delivered_json = {}
        reasons.append(f"JSON 不可解析：{error}")
    try:
        workbook = load_workbook(paths["Refund_Analysis.xlsx"], data_only=False)
    except Exception as error:
        workbook = None
        reasons.append(f"XLSX 不可解析：{error}")
    try:
        report_text, headings, report_tables, report_media = LEGACY.docx_payload(paths["Root_Cause_Report.docx"])
        Document(paths["Root_Cause_Report.docx"])
    except Exception as error:
        report_text, headings, report_tables, report_media = "", [], 0, 0
        reasons.append(f"DOCX 不可解析：{error}")
    try:
        quality_text = paths["quality_team_email.txt"].read_text(encoding="utf-8")
        supplier_text = paths["supplier_relations_email.txt"].read_text(encoding="utf-8")
        if not quality_text.strip() or not supplier_text.strip():
            raise ValueError("邮件草稿为空")
    except Exception as error:
        quality_text, supplier_text = "", ""
        reasons.append(f"邮件草稿不可解析：{error}")
    if reasons:
        if workbook is not None:
            workbook.close()
        return fatal_result(paths, reasons, extras)

    truth = LEGACY.derive_truth(input_dir)
    criteria: list[dict[str, Any]] = [positive("R001", 1.0, "六个核心交付存在、非空、可解析，且均非冻结输入副本")]

    canonical_payload, canonical_log = LEGACY.run_processor(paths["refund_root_cause_processor.py"], input_dir)
    matched, total = atomic_units(canonical_payload, truth)
    criteria.append(positive("R002", fraction(matched, total), f"原子匹配={matched}/{total}，日志={canonical_log}"))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-refund-v2-mutation-") as temp_dir:
        mutated_input = Path(temp_dir) / "inputs"
        LEGACY.build_mutated_inputs(input_dir, mutated_input)
        mutated_truth = LEGACY.derive_truth(mutated_input)
        mutated_payload, mutated_log = LEGACY.run_processor(paths["refund_root_cause_processor.py"], mutated_input)
    mutated_match, mutated_total = metamorphic_units(mutated_payload, mutated_truth)
    criteria.append(positive("R003", fraction(mutated_match, mutated_total), f"原子匹配={mutated_match}/{mutated_total}，日志={mutated_log}"))

    json_match, json_total = atomic_units(delivered_json, truth)
    criteria.append(positive("R004", fraction(json_match, json_total), f"原子匹配={json_match}/{json_total}"))

    detail_score, detail_evidence, _ = table_atomic_score(
        workbook, "Refund Details", DETAIL_HEADERS, truth["refund_details"], "Refund_ID"
    )
    impact_score, impact_evidence, _ = table_atomic_score(
        workbook, "Product Impact", IMPACT_HEADERS, truth["product_impact"], "Product_Name"
    )
    summary_score, summary_evidence, excel_summary = summary_atomic_score(workbook, truth["summary"])
    criteria.extend([
        positive("R005", detail_score, detail_evidence),
        positive("R006", impact_score, impact_evidence),
        positive("R007", summary_score, summary_evidence),
    ])

    word_score, word_evidence = word_atomic_score(report_text, headings, truth)
    email_score, email_evidence = email_atomic_score(quality_text, supplier_text, truth)
    criteria.extend([
        positive("R008", word_score, word_evidence),
        positive("R009", email_score, email_evidence),
    ])

    consistency_score, consistency_evidence = cross_artifact_score(
        delivered_json, excel_summary, report_text, quality_text, supplier_text
    )
    criteria.append(positive("R010", consistency_score, consistency_evidence))

    native_score, native_evidence = native_structure_score(
        workbook, report_tables, report_media, headings
    )
    criteria.append(positive("R011", native_score, native_evidence))
    workbook.close()

    if os.environ.get("COWORK_EVO_SKIP_RENDER_FOR_TESTS") == "1":
        render_score = 1.0
        render_evidence = "分层负例回归复用已通过的真实 Golden 渲染；生产评分未设置该变量"
    else:
        with tempfile.TemporaryDirectory(prefix="cowork-evo-refund-v2-render-") as temp_dir:
            rendered, render_log = LEGACY.render_office(
                [paths["Refund_Analysis.xlsx"], paths["Root_Cause_Report.docx"]], Path(temp_dir)
            )
        xlsx_render = rendered.get("Refund_Analysis.xlsx", {})
        docx_render = rendered.get("Root_Cause_Report.docx", {})
        xlsx_ok = (
            xlsx_render.get("pages", 0) >= 3
            and all(token in xlsx_render.get("text", "") for token in ("Refund_ID", "Product_Name", "Total_Refunds"))
        )
        docx_ok = (
            docx_render.get("pages", 0) >= 3
            and all(token in docx_render.get("text", "") for token in ("Executive Summary", "Product-Level Findings", "Action Plan"))
        )
        render_score = (int(xlsx_ok) + int(docx_ok)) / 2
        render_evidence = f"XLSX={xlsx_ok}/{xlsx_render.get('pages', 0)}页，DOCX={docx_ok}/{docx_render.get('pages', 0)}页，日志={render_log}"
    criteria.append(positive("R012", render_score, render_evidence))

    unsupported = sorted(set(LEGACY.unsupported_claims(
        report_text + "\n" + quality_text + "\n" + supplier_text, truth
    )))
    unsupported_fraction = min(1.0, len(unsupported) / 3.0)
    criteria.append(penalty("R013", unsupported_fraction, f"去重发现={unsupported}"))
    criteria.append(penalty("R014", 1.0 if extras else 0.0, f"额外文件={extras}"))
    return build_result(criteria, paths, [])


def evaluate(artifact_dir: Path) -> dict[str, Any]:
    version = os.environ.get("SCORING_VERSION", "v2").strip().lower()
    if version == "v1":
        result = LEGACY.evaluate(artifact_dir)
        result["scoring_version"] = "v1"
        return result
    if version != "v2":
        raise ValueError(f"不支持的 SCORING_VERSION={version!r}，仅支持 v1 或 v2")
    return evaluate_v2(artifact_dir)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    parser.add_argument("--scoring-version", choices=("v1", "v2"))
    args = parser.parse_args()
    if args.scoring_version:
        os.environ["SCORING_VERSION"] = args.scoring_version
    result = evaluate(args.artifact_dir.resolve())
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(rendered + "\n", encoding="utf-8")
    if args.reward_file:
        args.reward_file.parent.mkdir(parents=True, exist_ok=True)
        args.reward_file.write_text(f"{float(result['reward']):.6f}\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
