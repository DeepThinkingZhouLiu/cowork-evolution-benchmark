#!/usr/bin/env python3
"""从六份源输入独立重算真值的确定性 Harbor Judge。"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import defaultdict
from decimal import Decimal
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook


TASK_ID = "cowork-evo/cowork-bench-shipping-tax-update-propagation"
RUBRIC_VERSION = "cowork-bench-shipping-tax-update-propagation-v0.1"
JUDGE_VERSION = "1.0.0"
PASS_THRESHOLD = 0.8

OUTPUT_NAMES = [
    "shipping_tax_update_processor.py",
    "shipping_tax_results.json",
    "Shipping_Tax_Analysis.xlsx",
    "Operations_Report.docx",
    "operations_email.txt",
]
GENERATED_NAMES = OUTPUT_NAMES[1:]
REQUIRED_HEADINGS = [
    "Executive Summary",
    "Shipping Coverage",
    "Tax Structure",
    "Recommendations",
    "Source Boundary",
]

CRITERION_SPECS = {
    "R001": (6, "hurdle", "五个核心交付物完整且可打开"),
    "R002": (9, "hurdle", "processor 在原始输入上可回放并生成完整正确的四类结果"),
    "R003": (10, "hurdle", "Reduced Rate 5->7 后正确联动 JSON、Excel、Word 和邮件"),
    "R004": (8, "hurdle", "JSON 完整保留配送、税率明细及汇总真值"),
    "R005": (8, "hurdle", "Excel Shipping Zones 的区域、地域、方法和排序正确"),
    "R006": (8, "hurdle", "Excel Tax Rates 保留全部税率和空税类行"),
    "R007": (6, "hurdle", "Methods_Count 为引用同行 Methods 的动态公式"),
    "R008": (8, "hurdle", "Word 章节、配送明细、税率明细和汇总全部正确"),
    "R009": (6, "hurdle", "离线邮件的收件人、主题和核心事实正确"),
    "R010": (7, "hurdle", "JSON、Excel、Word 和邮件之间无核心事实冲突"),
    "R011": (5, "hurdle", "Word 与邮件没有无来源数字或业务断言"),
    "R012": (4, "positive", "Excel 和 Word 使用原生可编辑结构且样式完整"),
    "R013": (5, "positive", "Excel 和 Word 可被 LibreOffice 实际渲染"),
    "R014": (10, "hurdle", "六份输入的大小、行数和 SHA256 与溯源清单一致"),
    "R015": (-25, "penalty", "交付缺失、损坏或无法渲染"),
    "R016": (-22, "penalty", "processor 不可回放或对变异输入硬编码"),
    "R017": (-20, "penalty", "来源事实与 Judge 重算真值不一致"),
    "R018": (-18, "penalty", "Methods_Count 为常量、硬编码公式或不依赖 Methods"),
    "R019": (-20, "penalty", "JSON、Excel、Word 或邮件之间存在事实冲突"),
    "R020": (-18, "penalty", "编造无来源百分比、收入、需求、基准、ROI 或预测"),
}

HURDLE_CAPS = {
    "R001": 0.0,
    "R002": 0.35,
    "R003": 0.40,
    "R004": 0.45,
    "R005": 0.50,
    "R006": 0.50,
    "R007": 0.65,
    "R008": 0.55,
    "R009": 0.60,
    "R010": 0.60,
    "R011": 0.70,
    "R014": 0.40,
}


def criterion(identifier: str, passed: bool, evidence: str) -> dict[str, Any]:
    weight, kind, description = CRITERION_SPECS[identifier]
    return {
        "criterion_id": identifier,
        "passed": passed,
        "score": 1.0 if passed else 0.0,
        "weight": weight,
        "type": kind,
        "description": description,
        "method": "deterministic",
        "evidence": evidence,
    }


def read_csv(path: Path, required: tuple[str, ...]) -> list[dict[str, str]]:
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None or not set(required).issubset(reader.fieldnames):
            raise ValueError(f"{path.name} 缺列 {required}")
        return list(reader)


def source_bool(value: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"t", "true", "1", "yes"}:
        return True
    if normalized in {"f", "false", "0", "no"}:
        return False
    raise ValueError(f"无效布尔值 {value!r}")


def json_number(value: Decimal) -> int | float:
    return int(value) if value == value.to_integral() else float(value)


def rate_text(value: int | float | Decimal) -> str:
    return format(Decimal(str(value)).normalize(), "f")


def derive_truth(input_dir: Path) -> dict[str, Any]:
    """独立实现题意；不导入、执行或读取 solution。"""
    zones = read_csv(input_dir / "shipping_zones.csv", ("id", "name", "order"))
    locations = read_csv(
        input_dir / "shipping_zone_locations.csv", ("id", "zone_id", "code", "type")
    )
    zone_methods = read_csv(
        input_dir / "shipping_zone_methods.csv",
        ("id", "zone_id", "method_id", "method_title", "enabled", "settings"),
    )
    method_catalog = read_csv(
        input_dir / "shipping_methods.csv", ("id", "title", "description")
    )
    tax_classes = read_csv(input_dir / "tax_classes.csv", ("slug", "name"))
    tax_rates = read_csv(
        input_dir / "tax_rates.csv",
        (
            "id", "country", "state", "postcode", "city", "rate", "name",
            "priority", "compound", "shipping", "order", "class",
        ),
    )

    zone_ids = {int(row["id"]) for row in zones}
    catalog_ids = {row["id"] for row in method_catalog}
    class_names = {row["slug"]: row["name"] for row in tax_classes}
    if len(zone_ids) != len(zones) or len(catalog_ids) != len(method_catalog):
        raise ValueError("源输入含重复主键")

    locations_by_zone: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in sorted(locations, key=lambda item: int(item["id"])):
        zone_id = int(row["zone_id"])
        if zone_id not in zone_ids:
            raise ValueError(f"地域 {row['id']} 引用未知配送区")
        locations_by_zone[zone_id].append(
            {"location_id": int(row["id"]), "type": row["type"], "code": row["code"]}
        )

    methods_by_zone: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in sorted(zone_methods, key=lambda item: int(item["id"])):
        zone_id = int(row["zone_id"])
        if zone_id not in zone_ids or row["method_id"] not in catalog_ids:
            raise ValueError(f"配送方法 {row['id']} 外键无效")
        methods_by_zone[zone_id].append(
            {
                "record_id": int(row["id"]),
                "method_id": row["method_id"],
                "title": row["method_title"],
                "enabled": source_bool(row["enabled"]),
                "settings": json.loads(row["settings"] or "{}"),
            }
        )

    shipping_rows = []
    for zone in sorted(zones, key=lambda item: (int(item["order"]), int(item["id"]))):
        zone_id = int(zone["id"])
        enabled = [method for method in methods_by_zone[zone_id] if method["enabled"]]
        shipping_rows.append(
            {
                "zone_id": zone_id,
                "zone_name": zone["name"],
                "zone_order": int(zone["order"]),
                "locations": locations_by_zone[zone_id],
                "enabled_methods": enabled,
                "methods_count": len(enabled),
            }
        )

    rates_by_class: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in sorted(tax_rates, key=lambda item: int(item["id"])):
        slug = row["class"] or "standard"
        if slug not in class_names:
            raise ValueError(f"税率 {row['id']} 引用未知税类")
        rates_by_class[slug].append(
            {
                "rate_id": int(row["id"]),
                "country": row["country"],
                "state": row["state"],
                "postcode": row["postcode"],
                "city": row["city"],
                "rate": json_number(Decimal(row["rate"])),
                "rate_name": row["name"],
                "priority": int(row["priority"]),
                "compound": source_bool(row["compound"]),
                "shipping": source_bool(row["shipping"]),
                "order": int(row["order"]),
            }
        )

    class_rows = [
        {"slug": row["slug"], "name": row["name"], "rates": rates_by_class[row["slug"]]}
        for row in sorted(tax_classes, key=lambda item: item["name"])
    ]
    configured = [rate for item in class_rows for rate in item["rates"]]
    highest = max((Decimal(str(row["rate"])) for row in configured), default=Decimal("0"))
    return {
        "shipping_zones": shipping_rows,
        "tax_classes": class_rows,
        "summary": {
            "total_zones": len(shipping_rows),
            "total_enabled_methods": sum(row["methods_count"] for row in shipping_rows),
            "tax_classes": len(class_rows),
            "configured_tax_rates": len(configured),
            "displayed_tax_rows": sum(max(1, len(row["rates"])) for row in class_rows),
            "zones_with_single_method": sum(row["methods_count"] == 1 for row in shipping_rows),
            "highest_configured_rate": json_number(highest),
        },
    }


def close_number(actual: Any, expected: Any, tolerance: Decimal = Decimal("0.00001")) -> bool:
    try:
        return abs(Decimal(str(actual)) - Decimal(str(expected))) <= tolerance
    except Exception:
        return False


def payload_equal(actual: Any, expected: Any) -> bool:
    if isinstance(expected, dict):
        return isinstance(actual, dict) and set(actual) == set(expected) and all(
            payload_equal(actual[key], expected[key]) for key in expected
        )
    if isinstance(expected, list):
        return isinstance(actual, list) and len(actual) == len(expected) and all(
            payload_equal(left, right) for left, right in zip(actual, expected)
        )
    if isinstance(expected, (int, float, Decimal)) and not isinstance(expected, bool):
        return close_number(actual, expected)
    return actual == expected


def expected_shipping_excel(truth: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        {
            "Zone_Name": zone["zone_name"],
            "Methods_Count": zone["methods_count"],
            "Regions": "; ".join(f"{item['type']}:{item['code']}" for item in zone["locations"]),
            "Methods": "; ".join(method["title"] for method in zone["enabled_methods"]),
        }
        for zone in truth["shipping_zones"]
    ]


def expected_tax_excel(truth: dict[str, Any]) -> list[dict[str, Any]]:
    rows = []
    for tax_class in truth["tax_classes"]:
        if not tax_class["rates"]:
            rows.append(
                {"Tax_Class": tax_class["name"], "State": "", "Rate": None, "Compound": None, "Shipping": None}
            )
            continue
        for rate in tax_class["rates"]:
            rows.append(
                {
                    "Tax_Class": tax_class["name"],
                    "State": rate["state"],
                    "Rate": rate["rate"],
                    "Compound": rate["compound"],
                    "Shipping": rate["shipping"],
                }
            )
    return rows


def dynamic_count_formula(formula: Any, row_number: int) -> bool:
    if not isinstance(formula, str) or not formula.startswith("="):
        return False
    normalized = formula.upper().replace("$", "").replace(" ", "")
    if f"D{row_number}" not in normalized:
        return False
    uses_delimiter_count = "LEN(" in normalized and "SUBSTITUTE(" in normalized
    uses_split_count = "TEXTSPLIT(" in normalized and "COUNTA(" in normalized
    return uses_delimiter_count or uses_split_count


def normalize_blank(value: Any) -> Any:
    return "" if value is None else value


def read_excel(path: Path) -> tuple[dict[str, Any] | None, dict[str, Any], str]:
    try:
        workbook = load_workbook(path, data_only=False)
    except Exception as error:
        return None, {}, str(error)

    formatting = {
        "sheet_names": workbook.sheetnames,
        "table_count": sum(len(sheet.tables) for sheet in workbook.worksheets),
        "images": sum(len(sheet._images) for sheet in workbook.worksheets),
        "freezes": {sheet.title: str(sheet.freeze_panes or "") for sheet in workbook.worksheets},
        "header_bold": all(
            all(cell.font.bold for cell in sheet[1] if cell.value is not None)
            for sheet in workbook.worksheets
        ),
        "wide_columns": all(
            any((dimension.width or 0) >= 16 for dimension in sheet.column_dimensions.values())
            for sheet in workbook.worksheets
        ),
    }
    if workbook.sheetnames != ["Shipping Zones", "Tax Rates"]:
        workbook.close()
        return None, formatting, f"工作表错误：{workbook.sheetnames}"

    shipping = workbook["Shipping Zones"]
    shipping_headers = [shipping.cell(1, column).value for column in range(1, 5)]
    if shipping_headers != ["Zone_Name", "Methods_Count", "Regions", "Methods"]:
        workbook.close()
        return None, formatting, f"Shipping Zones 表头错误：{shipping_headers}"
    shipping_rows = []
    formula_checks = []
    for row_number in range(2, shipping.max_row + 1):
        values = [shipping.cell(row_number, column).value for column in range(1, 5)]
        if all(value in (None, "") for value in values):
            continue
        methods = normalize_blank(values[3])
        calculated = 0 if not methods else len(str(methods).split("; "))
        shipping_rows.append(
            {
                "Zone_Name": normalize_blank(values[0]),
                "Methods_Count": calculated,
                "Regions": normalize_blank(values[2]),
                "Methods": methods,
            }
        )
        formula_checks.append(
            {
                "row": row_number,
                "formula": values[1],
                "dynamic": dynamic_count_formula(values[1], row_number),
            }
        )

    taxes = workbook["Tax Rates"]
    tax_headers = [taxes.cell(1, column).value for column in range(1, 6)]
    if tax_headers != ["Tax_Class", "State", "Rate", "Compound", "Shipping"]:
        workbook.close()
        return None, formatting, f"Tax Rates 表头错误：{tax_headers}"
    tax_rows = []
    for row_number in range(2, taxes.max_row + 1):
        values = [taxes.cell(row_number, column).value for column in range(1, 6)]
        if all(value in (None, "") for value in values):
            continue
        tax_rows.append(
            {
                "Tax_Class": normalize_blank(values[0]),
                "State": normalize_blank(values[1]),
                "Rate": values[2],
                "Compound": values[3],
                "Shipping": values[4],
            }
        )
    workbook.close()
    return {
        "shipping_rows": shipping_rows,
        "tax_rows": tax_rows,
        "formula_checks": formula_checks,
    }, formatting, ""


def extract_docx(path: Path) -> tuple[str, list[str], list[list[list[str]]], int]:
    document = Document(path)
    pieces = []
    headings = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            pieces.append(text)
            style_name = (paragraph.style.name or "").lower()
            if style_name == "title" or style_name.startswith("heading"):
                headings.append(text)
    tables: list[list[list[str]]] = []
    for table in document.tables:
        table_rows = []
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            table_rows.append(cells)
            pieces.extend(cell for cell in cells if cell)
        tables.append(table_rows)
    with zipfile.ZipFile(path) as archive:
        media_count = sum(
            1 for name in archive.namelist() if name.startswith("word/media/") and not name.endswith("/")
        )
    return "\n".join(pieces), headings, tables, media_count


def find_table_rows(tables: list[list[list[str]]], first_cell: str) -> list[list[str]]:
    return [
        row
        for table in tables
        for row in table[1:]
        if row and row[0].strip() == first_cell
    ]


def contains_count_near(text: str, label: str, count: int, distance: int = 90) -> bool:
    return bool(
        re.search(
            re.escape(label) + rf".{{0,{distance}}}\b{count}\b",
            text,
            flags=re.IGNORECASE | re.DOTALL,
        )
        or re.search(
            rf"\b{count}\b.{{0,{distance}}}" + re.escape(label),
            text,
            flags=re.IGNORECASE | re.DOTALL,
        )
    )


def word_checks(
    text: str,
    headings: list[str],
    tables: list[list[list[str]]],
    truth: dict[str, Any],
) -> tuple[bool, list[str]]:
    errors = []
    for heading in REQUIRED_HEADINGS:
        if not any(item.casefold() == heading.casefold() for item in headings):
            errors.append(f"缺少章节 {heading}")

    summary = truth["summary"]
    for label, key in [
        ("Shipping zones", "total_zones"),
        ("Enabled methods", "total_enabled_methods"),
        ("Tax classes", "tax_classes"),
        ("Configured tax rates", "configured_tax_rates"),
        ("Displayed tax rows", "displayed_tax_rows"),
    ]:
        if not contains_count_near(text, label, summary[key]):
            errors.append(f"汇总缺少 {label}={summary[key]}")
    highest = f"{rate_text(summary['highest_configured_rate'])}%"
    if highest not in text:
        errors.append(f"缺少最高税率 {highest}")

    for zone in truth["shipping_zones"]:
        candidates = find_table_rows(tables, zone["zone_name"])
        regions = "; ".join(f"{item['type']}:{item['code']}" for item in zone["locations"])
        methods = "; ".join(method["title"] for method in zone["enabled_methods"])
        exact = any(
            len(row) >= 4
            and row[1].strip() == str(zone["methods_count"])
            and row[2].strip() == regions
            and row[3].strip() == methods
            for row in candidates
        )
        if not exact:
            errors.append(f"配送区明细错误 {zone['zone_name']}")

    for tax_class in truth["tax_classes"]:
        if not tax_class["rates"]:
            candidates = find_table_rows(tables, tax_class["name"])
            if not any(
                len(row) >= 2 and row[1].strip().casefold() == "no configured rate"
                for row in candidates
            ):
                errors.append(f"缺少空税类 {tax_class['name']}")
            continue
        candidates = find_table_rows(tables, tax_class["name"])
        for rate in tax_class["rates"]:
            scope = " / ".join(
                value
                for value in (rate["country"], rate["state"], rate["postcode"], rate["city"])
                if value
            )
            expected_rate = f"{rate_text(rate['rate'])}%"
            exact = any(
                len(row) >= 6
                and row[1].strip() == rate["rate_name"]
                and row[2].strip() == scope
                and row[3].strip() == expected_rate
                and row[4].strip().casefold() == str(rate["compound"]).lower()
                and row[5].strip().casefold() == str(rate["shipping"]).lower()
                for row in candidates
            )
            if not exact:
                errors.append(f"税率明细错误 {tax_class['name']}/{rate['rate_name']}")
    return not errors, errors


def email_checks(text: str, truth: dict[str, Any]) -> tuple[bool, list[str]]:
    errors = []
    if not re.search(r"^To:\s*operations@store\.com\s*$", text, re.IGNORECASE | re.MULTILINE):
        errors.append("收件人错误")
    if not re.search(
        r"^Subject:\s*Shipping and Tax Configuration Report\s*$",
        text,
        re.IGNORECASE | re.MULTILINE,
    ):
        errors.append("主题错误")
    summary = truth["summary"]
    for label, key in [
        ("shipping zones", "total_zones"),
        ("enabled shipping methods", "total_enabled_methods"),
        ("tax classes", "tax_classes"),
        ("configured tax-rate records", "configured_tax_rates"),
    ]:
        if not contains_count_near(text, label, summary[key], 70):
            errors.append(f"邮件缺少 {label}={summary[key]}")
    highest = f"{rate_text(summary['highest_configured_rate'])}%"
    if highest not in text:
        errors.append("邮件最高税率错误")
    for zone in truth["shipping_zones"]:
        line = next((item for item in text.splitlines() if zone["zone_name"] in item), "")
        if not re.search(rf"\b{zone['methods_count']}\b", line):
            errors.append(f"邮件配送区方法数错误 {zone['zone_name']}")
    for tax_class in truth["tax_classes"]:
        line = next((item for item in text.splitlines() if tax_class["name"] in item), "")
        if not re.search(rf"\b{len(tax_class['rates'])}\b", line):
            errors.append(f"邮件税类记录数错误 {tax_class['name']}")
            continue
        if tax_class["rates"]:
            for rate in tax_class["rates"]:
                if f"{rate_text(rate['rate'])}%" not in line:
                    errors.append(f"邮件税率错误 {tax_class['name']}/{rate['rate_id']}")
        elif "none" not in line.casefold():
            errors.append(f"邮件未标明空税类 {tax_class['name']}")
    return not errors, errors


def unsupported_claims(text: str, truth: dict[str, Any]) -> list[str]:
    allowed_rates = {
        Decimal(str(rate["rate"]))
        for tax_class in truth["tax_classes"]
        for rate in tax_class["rates"]
    }
    findings = []
    for match in re.finditer(r"(?<![\w.])([0-9]+(?:\.[0-9]+)?)\s*%", text):
        value = Decimal(match.group(1))
        if value not in allowed_rates:
            findings.append(f"无来源百分比 {match.group(1)}%")

    assertion_patterns = [
        (r"high\s+(?:order\s+volume|demand)", "高订单量/高需求"),
        (r"revenue\s+(?:growth|increase|impact|loss)", "收入影响"),
        (r"cost\s+savings?", "成本节省"),
        (r"customer\s+(?:demand|churn|loss)", "客户需求/流失"),
        (r"market\s+benchmark|industry\s+(?:benchmark|average)", "市场/行业基准"),
        (r"\broi\b", "ROI"),
        (r"will\s+(?:increase|decrease|grow|reduce|save)", "预测性断言"),
    ]
    for sentence in re.split(r"(?<=[.!?])\s+|\n+", text):
        lowered = sentence.casefold()
        if not lowered.strip():
            continue
        negated = bool(re.search(r"\b(?:no|not|none|without|isn't|aren't|wasn't|weren't)\b", lowered))
        for pattern, label in assertion_patterns:
            if re.search(pattern, lowered) and not negated:
                findings.append(label)
    return sorted(set(findings))


def compare_tax_rows(actual: list[dict[str, Any]], expected: list[dict[str, Any]]) -> bool:
    if len(actual) != len(expected):
        return False
    for left, right in zip(actual, expected):
        if left["Tax_Class"] != right["Tax_Class"] or left["State"] != right["State"]:
            return False
        if right["Rate"] is None:
            if left["Rate"] is not None:
                return False
        elif not close_number(left["Rate"], right["Rate"]):
            return False
        if left["Compound"] != right["Compound"] or left["Shipping"] != right["Shipping"]:
            return False
    return True


def inspect_package(output_dir: Path, truth: dict[str, Any]) -> dict[str, Any]:
    missing = [name for name in GENERATED_NAMES if not (output_dir / name).is_file()]
    inspection: dict[str, Any] = {"missing": missing}
    if missing:
        inspection.update(
            json_ok=False, shipping_excel_ok=False, tax_excel_ok=False, formula_ok=False,
            word_ok=False, email_ok=False, unsupported=["generated files missing"],
            excel_error="", word_errors=[], email_errors=[], formatting={},
            word_tables=0, word_media=0,
        )
        return inspection

    try:
        payload = json.loads((output_dir / "shipping_tax_results.json").read_text(encoding="utf-8"))
        inspection["json_ok"] = payload_equal(payload, truth)
        inspection["json_error"] = ""
    except Exception as error:
        inspection["json_ok"] = False
        inspection["json_error"] = str(error)

    excel_payload, formatting, excel_error = read_excel(output_dir / "Shipping_Tax_Analysis.xlsx")
    inspection["formatting"] = formatting
    inspection["excel_error"] = excel_error
    inspection["shipping_excel_ok"] = bool(excel_payload) and payload_equal(
        excel_payload["shipping_rows"], expected_shipping_excel(truth)
    )
    inspection["tax_excel_ok"] = bool(excel_payload) and compare_tax_rows(
        excel_payload["tax_rows"], expected_tax_excel(truth)
    )
    inspection["formula_ok"] = bool(excel_payload) and (
        len(excel_payload["formula_checks"]) == len(truth["shipping_zones"])
        and all(row["dynamic"] for row in excel_payload["formula_checks"])
    )
    inspection["formula_checks"] = excel_payload["formula_checks"] if excel_payload else []

    try:
        word_text, headings, tables, media = extract_docx(output_dir / "Operations_Report.docx")
        word_ok, word_errors = word_checks(word_text, headings, tables, truth)
    except Exception as error:
        word_text, headings, tables, media = "", [], [], 0
        word_ok, word_errors = False, [str(error)]
    inspection.update(
        word_ok=word_ok,
        word_errors=word_errors,
        word_text=word_text,
        word_headings=headings,
        word_tables=len(tables),
        word_media=media,
    )

    try:
        email_text = (output_dir / "operations_email.txt").read_text(encoding="utf-8")
        email_ok, email_errors = email_checks(email_text, truth)
    except Exception as error:
        email_text, email_ok, email_errors = "", False, [str(error)]
    inspection.update(email_ok=email_ok, email_errors=email_errors, email_text=email_text)
    inspection["unsupported"] = unsupported_claims(word_text + "\n" + email_text, truth)
    return inspection


def package_content_ok(inspection: dict[str, Any]) -> bool:
    return all(
        inspection.get(key, False)
        for key in (
            "json_ok", "shipping_excel_ok", "tax_excel_ok", "formula_ok", "word_ok", "email_ok"
        )
    ) and not inspection.get("unsupported")


def run_processor(script: Path, input_dir: Path, truth: dict[str, Any]) -> tuple[dict[str, Any], str]:
    with tempfile.TemporaryDirectory(prefix="cowork-evo-shipping-tax-replay-") as temp_dir:
        output_dir = Path(temp_dir) / "output"
        try:
            completed = subprocess.run(
                [
                    sys.executable,
                    str(script.resolve()),
                    "--input-dir",
                    str(input_dir.resolve()),
                    "--output-dir",
                    str(output_dir),
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=120,
                check=False,
            )
        except Exception as error:
            return {"missing": GENERATED_NAMES}, f"执行异常：{error}"
        if completed.returncode != 0:
            return {"missing": GENERATED_NAMES}, f"退出={completed.returncode}，日志={completed.stdout[-900:]}"
        inspection = inspect_package(output_dir, truth)
        return inspection, completed.stdout[-900:]


def build_mutated_inputs(source: Path, target: Path) -> None:
    shutil.copytree(source, target)
    rate_path = target / "tax_rates.csv"
    with rate_path.open(encoding="utf-8", newline="") as handle:
        rows = list(csv.DictReader(handle))
        fields = list(rows[0])
    matched = 0
    for row in rows:
        if row["id"] == "7":
            if row["rate"] != "5.0000" or row["class"] != "reduced-rate":
                raise ValueError("变异锚点与预期不一致")
            row["rate"] = "7.0000"
            matched += 1
    if matched != 1:
        raise ValueError(f"变异锚点数量错误：{matched}")
    with rate_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def render_office(paths: list[Path], target: Path) -> tuple[dict[str, dict[str, Any]], str]:
    rendered = {}
    logs = []
    target.mkdir(parents=True, exist_ok=True)
    profile = target / "lo-profile"
    for path in paths:
        command = [
            "libreoffice",
            "--headless",
            f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(target),
            str(path),
        ]
        try:
            completed = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=120,
                check=False,
            )
        except Exception as error:
            logs.append(f"{path.name}: {error}")
            continue
        pdf_path = target / f"{path.stem}.pdf"
        logs.append(
            f"{path.name}: rc={completed.returncode}, pdf={pdf_path.is_file()}, {completed.stdout[-220:]}"
        )
        if completed.returncode != 0 or not pdf_path.is_file():
            continue
        info = subprocess.run(
            ["pdfinfo", str(pdf_path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            check=False,
        )
        page_match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
        text_result = subprocess.run(
            ["pdftotext", "-layout", str(pdf_path), "-"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            check=False,
        )
        rendered[path.name] = {
            "pages": int(page_match.group(1)) if page_match else 0,
            "text": text_result.stdout,
            "pdf": str(pdf_path),
        }
    return rendered, " | ".join(logs)


def verify_manifest(task_root: Path) -> tuple[bool, list[str]]:
    errors = []
    try:
        manifest = json.loads((task_root / "metadata" / "input-manifest.json").read_text(encoding="utf-8"))
        for item in manifest["files"]:
            path = task_root / item["path"]
            if not path.is_file():
                errors.append(f"缺输入 {item['path']}")
                continue
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            size = path.stat().st_size
            with path.open(encoding="utf-8", newline="") as handle:
                rows = sum(1 for _ in csv.DictReader(handle))
            if digest != item["sha256"]:
                errors.append(f"SHA256 错误 {item['path']}")
            if size != item["bytes"]:
                errors.append(f"大小错误 {item['path']}")
            if rows != item["rows"]:
                errors.append(f"行数错误 {item['path']}")
    except Exception as error:
        errors.append(str(error))
    return not errors, errors


def build_result(results: list[dict[str, Any]], artifact_paths: list[Path]) -> dict[str, Any]:
    present = {item["criterion_id"] for item in results}
    for identifier in CRITERION_SPECS:
        if identifier not in present:
            weight, kind, _ = CRITERION_SPECS[identifier]
            results.append(
                criterion(
                    identifier,
                    kind == "penalty" and False,
                    "因交付缺失或前置门槛失败，本判据未执行",
                )
            )
    results.sort(key=lambda item: item["criterion_id"])
    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdle_failures = [
        item["criterion_id"]
        for item in results
        if item["type"] == "hurdle" and not item["passed"]
    ]
    if hurdle_failures:
        reward = min(reward, min(HURDLE_CAPS[item] for item in hurdle_failures))
    digests = [
        hashlib.sha256(path.read_bytes()).hexdigest()
        for path in artifact_paths
        if path.is_file()
    ] or [hashlib.sha256(b"").hexdigest()]
    criterion_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        criterion_results.append(
            {
                "criterion_id": item["criterion_id"],
                "score": score,
                "evidence": item["evidence"],
                "method": item["method"],
                "raw": {
                    "description": item["description"],
                    "weight": item["weight"],
                    "type": item["type"],
                    "passed": item["passed"],
                },
            }
        )
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": digests,
        "criterion_results": criterion_results,
        "reward": round(reward, 6),
        "passed": reward >= PASS_THRESHOLD,
        "hurdle_failures": hurdle_failures,
    }


def evaluate(artifact_dir: Path, render: bool = True) -> dict[str, Any]:
    task_root = Path(__file__).resolve().parents[1]
    input_dir = task_root / "data" / "input_files"
    paths = {name: artifact_dir / name for name in OUTPUT_NAMES}
    results = []
    missing = [name for name, path in paths.items() if not path.is_file()]
    results.append(criterion("R001", not missing, f"缺失={missing}"))
    if missing:
        results.append(criterion("R015", True, f"缺失={missing}"))
        return build_result(results, list(paths.values()))

    truth = derive_truth(input_dir)
    delivered = inspect_package(artifact_dir, truth)

    canonical_replay, canonical_log = run_processor(
        paths["shipping_tax_update_processor.py"], input_dir, truth
    )
    canonical_ok = package_content_ok(canonical_replay)
    results.append(
        criterion(
            "R002",
            canonical_ok,
            f"原始回放={canonical_ok}，缺失={canonical_replay.get('missing')}，"
            f"错误={canonical_replay.get('excel_error', '')}/{canonical_replay.get('word_errors', [])}/"
            f"{canonical_replay.get('email_errors', [])}，日志={canonical_log}",
        )
    )

    with tempfile.TemporaryDirectory(prefix="cowork-evo-shipping-tax-mutation-") as temp_dir:
        mutated_input = Path(temp_dir) / "inputs"
        build_mutated_inputs(input_dir, mutated_input)
        mutated_truth = derive_truth(mutated_input)
        mutated_replay, mutated_log = run_processor(
            paths["shipping_tax_update_processor.py"], mutated_input, mutated_truth
        )
    canonical_reduced = next(
        rate["rate"]
        for item in truth["tax_classes"] if item["slug"] == "reduced-rate"
        for rate in item["rates"] if rate["rate_id"] == 7
    )
    mutated_reduced = next(
        rate["rate"]
        for item in mutated_truth["tax_classes"] if item["slug"] == "reduced-rate"
        for rate in item["rates"] if rate["rate_id"] == 7
    )
    mutation_changed = close_number(canonical_reduced, 5) and close_number(mutated_reduced, 7)
    mutation_ok = mutation_changed and package_content_ok(mutated_replay)
    results.append(
        criterion(
            "R003",
            mutation_ok,
            f"真值={canonical_reduced}->{mutated_reduced}，变异回放={package_content_ok(mutated_replay)}，"
            f"Excel={mutated_replay.get('tax_excel_ok')}，Word={mutated_replay.get('word_ok')}，"
            f"JSON={mutated_replay.get('json_ok')}，Email={mutated_replay.get('email_ok')}，日志={mutated_log}",
        )
    )

    results.append(
        criterion("R004", delivered["json_ok"], f"JSON={delivered['json_ok']}，错误={delivered.get('json_error', '')}")
    )
    results.append(
        criterion(
            "R005",
            delivered["shipping_excel_ok"],
            f"Shipping Zones={delivered['shipping_excel_ok']}，错误={delivered.get('excel_error', '')}",
        )
    )
    results.append(
        criterion(
            "R006",
            delivered["tax_excel_ok"],
            f"Tax Rates={delivered['tax_excel_ok']}，错误={delivered.get('excel_error', '')}",
        )
    )
    results.append(
        criterion(
            "R007",
            delivered["formula_ok"],
            f"公式={delivered.get('formula_checks', [])}",
        )
    )
    results.append(
        criterion(
            "R008",
            delivered["word_ok"],
            f"Word={delivered['word_ok']}，错误={delivered.get('word_errors', [])}",
        )
    )
    results.append(
        criterion(
            "R009",
            delivered["email_ok"],
            f"Email={delivered['email_ok']}，错误={delivered.get('email_errors', [])}",
        )
    )
    cross_ok = all(
        delivered[key]
        for key in ("json_ok", "shipping_excel_ok", "tax_excel_ok", "word_ok", "email_ok")
    )
    results.append(
        criterion(
            "R010",
            cross_ok,
            f"JSON={delivered['json_ok']}，Excel={delivered['shipping_excel_ok']}/{delivered['tax_excel_ok']}，"
            f"Word={delivered['word_ok']}，Email={delivered['email_ok']}",
        )
    )
    no_unsupported = not delivered["unsupported"]
    results.append(
        criterion("R011", no_unsupported, f"发现={delivered['unsupported']}")
    )

    formatting = delivered.get("formatting", {})
    native_ok = (
        formatting.get("sheet_names") == ["Shipping Zones", "Tax Rates"]
        and formatting.get("table_count", 0) >= 2
        and formatting.get("images") == 0
        and formatting.get("freezes") == {"Shipping Zones": "A2", "Tax Rates": "A2"}
        and formatting.get("header_bold")
        and formatting.get("wide_columns")
        and delivered.get("word_tables", 0) >= 3
        and delivered.get("word_media") == 0
        and len(delivered.get("word_headings", [])) >= 6
    )
    results.append(
        criterion(
            "R012",
            native_ok,
            json.dumps(
                formatting
                | {
                    "word_tables": delivered.get("word_tables"),
                    "word_media": delivered.get("word_media"),
                    "word_headings": len(delivered.get("word_headings", [])),
                },
                ensure_ascii=False,
            ),
        )
    )

    if render:
        with tempfile.TemporaryDirectory(prefix="cowork-evo-shipping-tax-render-") as temp_dir:
            rendered, render_log = render_office(
                [paths["Shipping_Tax_Analysis.xlsx"], paths["Operations_Report.docx"]],
                Path(temp_dir),
            )
        xlsx_render = rendered.get("Shipping_Tax_Analysis.xlsx", {})
        docx_render = rendered.get("Operations_Report.docx", {})
        render_ok = (
            xlsx_render.get("pages", 0) >= 2
            and all(
                token in xlsx_render.get("text", "")
                for token in ("Zone_Name", "Methods_Count", "Tax_Class", "Reduced Rate")
            )
            and docx_render.get("pages", 0) >= 1
            and all(
                token in docx_render.get("text", "")
                for token in ("Executive Summary", "Shipping Coverage", "Tax Structure", "Recommendations")
            )
        )
        render_evidence = (
            f"pages={{'xlsx': {xlsx_render.get('pages', 0)}, 'docx': {docx_render.get('pages', 0)}}}，"
            f"日志={render_log}"
        )
    else:
        render_ok = True
        render_evidence = "回归负例快速模式：跳过外部渲染，不影响内容判定"
    results.append(criterion("R013", render_ok, render_evidence))

    manifest_ok, manifest_errors = verify_manifest(task_root)
    results.append(criterion("R014", manifest_ok, f"错误={manifest_errors}"))

    damaged = bool(delivered.get("missing")) or not delivered.get("formatting") or not delivered.get("word_text") or not render_ok
    replay_error = not (canonical_ok and mutation_ok)
    fact_error = not all(
        delivered[key]
        for key in ("json_ok", "shipping_excel_ok", "tax_excel_ok", "word_ok", "email_ok")
    )
    formula_error = not delivered["formula_ok"]
    conflict = not cross_ok
    unsupported = bool(delivered["unsupported"])
    results.append(criterion("R015", damaged, f"损坏={damaged}，渲染={render_ok}"))
    results.append(criterion("R016", replay_error, f"原始={canonical_ok}，变异={mutation_ok}"))
    results.append(criterion("R017", fact_error, f"真值不一致={fact_error}"))
    results.append(criterion("R018", formula_error, f"公式错误={formula_error}"))
    results.append(criterion("R019", conflict, f"跨文件冲突={conflict}"))
    results.append(criterion("R020", unsupported, f"无来源断言={delivered['unsupported']}"))
    return build_result(results, list(paths.values()))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", required=True, type=Path)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    parser.add_argument("--skip-render", action="store_true", help="仅用于本题负例快速回归")
    args = parser.parse_args()
    result = evaluate(args.artifact_dir.resolve(), render=not args.skip_render)
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
