#!/usr/bin/env python3
"""从六份原始输入重算真值的确定性 Harbor Judge。"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook


TASK_ID = "cowork-evo/cowork-bench-refund-root-cause-package"
RUBRIC_VERSION = "cowork-bench-refund-root-cause-package-v0.1"
JUDGE_VERSION = "cowork-bench-refund-root-cause-package-v0.1"
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
ROOT_CAUSE_MAP = {
    "Product defective": "Manufacturing Defect",
    "Product arrived damaged": "Shipping Damage",
    "Wrong item received": "Wrong Item",
    "Changed my mind": "Customer Expectations",
    "Item not as described": "Customer Expectations",
    "Duplicate order": "Other",
}
HURDLE_CAPS = {
    "R001": 0.0,
    "R002": 0.3,
    "R003": 0.4,
    "R004": 0.4,
    "R005": 0.45,
    "R006": 0.45,
    "R007": 0.5,
    "R008": 0.55,
    "R009": 0.55,
    "R010": 0.6,
    "R013": 0.7,
}
CRITERION_SPECS = {
    "R001": ("六个指定交付物均存在且可打开", 6, "hurdle"),
    "R002": ("处理器在原始输入上可回放且输出等于输入重算真值", 8, "hurdle"),
    "R003": ("变更退款金额与评分后处理器结果正确跟随，不是固定答案", 9, "hurdle"),
    "R004": ("交付 JSON 全部明细、商品聚合、汇总、原因与根因正确", 8, "hurdle"),
    "R005": ("Excel Refund Details 的表头、13 条记录、排序与数值正确", 7, "hurdle"),
    "R006": ("Excel Product Impact 的 24 个商品、分摊金额、评分与调查标记正确", 8, "hurdle"),
    "R007": ("Excel Summary 九项指标与输入重算真值一致", 6, "hurdle"),
    "R008": ("Word 报告结构完整，核心数值、原因、根因和四个调查商品正确", 8, "hurdle"),
    "R009": ("两封离线邮件草稿的收件人、主题和事实完整正确", 7, "hurdle"),
    "R010": ("JSON、Excel、Word 和两封邮件的核心事实无跨文件冲突", 6, "hurdle"),
    "R011": ("Excel 和 Word 保持原生可编辑结构，格式达到专业交付要求", 4, "positive"),
    "R012": ("Excel 和 Word 可由 LibreOffice 逐页渲染且核心内容可提取", 3, "positive"),
    "R013": ("报告与邮件没有输入不支持的金额、百分比、基准或效果断言", 5, "hurdle"),
    "R014": ("任一来源派生数值与输入重算真值冲突", -18, "penalty"),
    "R015": ("处理脚本不可回放或不随变异输入改变", -18, "penalty"),
    "R016": ("指定交付物缺失、损坏、截图化或无法渲染", -20, "penalty"),
    "R017": ("核心数值或商品事实在交付物之间冲突", -18, "penalty"),
    "R018": ("出现输入不支持的金额、百分比、行业基准或业务效果断言", -18, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str = "") -> dict:
    description, weight, kind = CRITERION_SPECS[identifier]
    return {
        "criterion_id": identifier,
        "description": description,
        "weight": weight,
        "type": kind,
        "passed": bool(passed),
        "score": 1.0 if passed else 0.0,
        "evidence": evidence,
        "method": "deterministic",
    }


def round_money(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def json_number(value: Decimal) -> float:
    return float(round_money(value))


def derive_truth(input_dir: Path) -> dict[str, Any]:
    with (input_dir / "refunds.csv").open(encoding="utf-8", newline="") as handle:
        refunds = list(csv.DictReader(handle))
    orders_payload = json.loads((input_dir / "refund_orders.json").read_text(encoding="utf-8"))
    orders = {int(row["id"]): row for row in orders_payload}
    ratings: dict[int, list[int]] = defaultdict(list)
    with (input_dir / "product_reviews.csv").open(encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            ratings[int(row["product_id"])].append(int(row["rating"]))

    details = []
    impact: dict[str, dict[str, Any]] = defaultdict(
        lambda: {"count": 0, "amount": Decimal("0"), "product_id": None}
    )
    severities: Counter[str] = Counter()
    reasons: Counter[str] = Counter()
    roots: Counter[str] = Counter()
    for refund in sorted(refunds, key=lambda row: int(row["id"])):
        refund_id = int(refund["id"])
        order_id = int(refund["order_id"])
        amount = Decimal(refund["amount"])
        items = orders[order_id]["line_items"]
        severity = "Critical" if amount > 50 else "Major" if amount >= 20 else "Minor"
        reason = refund["reason"]
        details.append(
            {
                "Refund_ID": refund_id,
                "Order_ID": order_id,
                "Product_Name": "; ".join(str(item["name"]) for item in items),
                "Refund_Amount": json_number(amount),
                "Severity": severity,
                "Refund_Reason": reason,
                "Date": refund["date_created"][:10],
            }
        )
        severities[severity] += 1
        reasons[reason] += 1
        roots[ROOT_CAUSE_MAP.get(reason, "Other")] += 1
        share = amount / Decimal(len(items))
        for item in items:
            aggregate = impact[str(item["name"])]
            aggregate["count"] += 1
            aggregate["amount"] += share
            aggregate["product_id"] = int(item["product_id"])

    products = []
    for name, aggregate in impact.items():
        values = ratings.get(int(aggregate["product_id"]), [])
        average = round_money(Decimal(sum(values)) / len(values)) if values else None
        issue = "No Reviews" if average is None else "Quality Issue" if average < 3 else "Service Issue"
        products.append(
            {
                "Product_Name": name,
                "Refund_Count": aggregate["count"],
                "Total_Refund_Amount": json_number(aggregate["amount"]),
                "Avg_Review_Rating": None if average is None else float(average),
                "Review_Count": len(values),
                "Issue_Type": issue,
                "Requires_Investigation": "Yes" if aggregate["count"] >= 2 else "No",
            }
        )
    products.sort(key=lambda row: (-row["Refund_Count"], row["Product_Name"]))
    summary = {
        "Total_Refunds": len(details),
        "Total_Refund_Amount": json_number(sum(Decimal(row["amount"]) for row in refunds)),
        "Critical_Count": severities["Critical"],
        "Major_Count": severities["Major"],
        "Minor_Count": severities["Minor"],
        "Products_Affected": len(products),
        "Quality_Issues": sum(row["Issue_Type"] == "Quality Issue" for row in products),
        "Service_Issues": sum(row["Issue_Type"] == "Service Issue" for row in products),
        "Products_Requiring_Investigation": sum(row["Requires_Investigation"] == "Yes" for row in products),
    }
    reason_rows = [
        {"Refund_Reason": reason, "Refund_Count": count}
        for reason, count in sorted(reasons.items(), key=lambda item: (-item[1], item[0]))
    ]
    root_rows = [
        {"Root_Cause": root, "Refund_Count": count}
        for root, count in sorted(roots.items(), key=lambda item: (-item[1], item[0]))
    ]
    return {
        "refund_details": details,
        "product_impact": products,
        "summary": summary,
        "refund_reason_breakdown": reason_rows,
        "root_cause_breakdown": root_rows,
        "investigation_products": [row for row in products if row["Requires_Investigation"] == "Yes"],
    }


def close(actual: Any, expected: Any, tolerance: float = 0.011) -> bool:
    try:
        return abs(float(actual) - float(expected)) <= tolerance
    except (TypeError, ValueError):
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
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return close(actual, expected)
    return actual == expected


def run_processor(script: Path, input_dir: Path) -> tuple[dict | None, str]:
    with tempfile.TemporaryDirectory(prefix="cowork-evo-refund-processor-") as temp_dir:
        output = Path(temp_dir) / "result.json"
        try:
            completed = subprocess.run(
                [sys.executable, str(script.resolve()), "--input-dir", str(input_dir.resolve()), "--output", str(output)],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=90,
                check=False,
            )
        except Exception as error:
            return None, f"执行异常：{error}"
        if completed.returncode != 0 or not output.is_file():
            return None, f"退出={completed.returncode}，输出存在={output.is_file()}，日志={completed.stdout[-800:]}"
        try:
            return json.loads(output.read_text(encoding="utf-8")), completed.stdout[-800:]
        except Exception as error:
            return None, f"JSON 无法读取：{error}"


def build_mutated_inputs(source: Path, target: Path) -> None:
    shutil.copytree(source, target)
    refund_path = target / "refunds.csv"
    with refund_path.open(encoding="utf-8", newline="") as handle:
        rows = list(csv.DictReader(handle))
        fields = list(rows[0])
    for row in rows:
        if row["id"] == "4":
            row["amount"] = "16.45"
    with refund_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)

    review_path = target / "product_reviews.csv"
    with review_path.open(encoding="utf-8", newline="") as handle:
        review_rows = list(csv.DictReader(handle))
        review_fields = list(review_rows[0])
    for row in review_rows:
        if row["product_id"] == "34":
            row["rating"] = "1"
    with review_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=review_fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(review_rows)


def normalize_cell(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def read_excel(path: Path) -> tuple[dict[str, Any] | None, dict[str, Any], str]:
    try:
        workbook = load_workbook(path, data_only=False)
    except Exception as error:
        return None, {}, str(error)
    formatting = {
        "sheet_names": workbook.sheetnames,
        "tables": sum(len(sheet.tables) for sheet in workbook.worksheets),
        "freezes": {sheet.title: str(sheet.freeze_panes or "") for sheet in workbook.worksheets},
        "images": sum(len(sheet._images) for sheet in workbook.worksheets),
        "header_bold": all(
            all(cell.font.bold for cell in sheet[1] if cell.value is not None)
            for sheet in workbook.worksheets
        ),
        "column_widths": all(
            any((dimension.width or 0) > 12 for dimension in sheet.column_dimensions.values())
            for sheet in workbook.worksheets
        ),
    }
    expected_sheets = ["Refund Details", "Product Impact", "Summary"]
    if workbook.sheetnames != expected_sheets:
        workbook.close()
        return None, formatting, f"工作表={workbook.sheetnames}"

    def table_rows(sheet_name: str, headers: list[str]) -> list[dict[str, Any]] | None:
        sheet = workbook[sheet_name]
        actual_headers = [sheet.cell(1, col).value for col in range(1, len(headers) + 1)]
        if actual_headers != headers:
            return None
        rows = []
        for row_number in range(2, sheet.max_row + 1):
            values = [normalize_cell(sheet.cell(row_number, col).value) for col in range(1, len(headers) + 1)]
            if all(value in (None, "") for value in values):
                continue
            rows.append(dict(zip(headers, values)))
        return rows

    details = table_rows("Refund Details", DETAIL_HEADERS)
    products = table_rows("Product Impact", IMPACT_HEADERS)
    summary_rows = table_rows("Summary", ["Metric", "Value"])
    workbook.close()
    if details is None or products is None or summary_rows is None:
        return None, formatting, "表头不匹配"
    summary = {str(row["Metric"]): row["Value"] for row in summary_rows}
    return {"details": details, "products": products, "summary": summary}, formatting, ""


def rows_equal(actual: list[dict], expected: list[dict]) -> bool:
    return payload_equal(actual, expected)


def docx_payload(path: Path) -> tuple[str, list[str], int, int]:
    document = Document(path)
    pieces: list[str] = []
    headings: list[str] = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text:
            pieces.append(text)
            if paragraph.style.name == "Title" or paragraph.style.name.lower().startswith("heading"):
                headings.append(text)
    table_count = len(document.tables)
    for table in document.tables:
        for row in table.rows:
            pieces.extend(cell.text.strip() for cell in row.cells if cell.text.strip())
    with zipfile.ZipFile(path) as archive:
        media = sum(1 for name in archive.namelist() if name.startswith("word/media/") and not name.endswith("/"))
    return "\n".join(pieces), headings, table_count, media


def has_money(text: str, value: float) -> bool:
    formatted = f"{value:,.2f}"
    plain = f"{value:.2f}"
    return formatted in text or plain in text


def label_count(text: str, label: str, count: int, distance: int = 50) -> bool:
    label_then_count = re.search(
        re.escape(label) + rf".{{0,{distance}}}\b{count}\b",
        text,
        re.IGNORECASE | re.DOTALL,
    )
    count_then_label = re.search(
        rf"\b{count}\b.{{0,{distance}}}" + re.escape(label),
        text,
        re.IGNORECASE | re.DOTALL,
    )
    return label_then_count is not None or count_then_label is not None


def word_checks(text: str, headings: list[str], truth: dict) -> tuple[bool, list[str]]:
    summary = truth["summary"]
    errors = []
    required = [
        "Executive Summary", "Refund Trend Analysis", "Product-Level Findings",
        "Root Cause Classification", "Recommendations", "Action Plan",
    ]
    for section in required:
        if not any(section.lower() == heading.lower() for heading in headings):
            errors.append(f"缺章节 {section}")
    if not has_money(text, summary["Total_Refund_Amount"]):
        errors.append("缺总金额")
    for label, key in [
        ("refund", "Total_Refunds"), ("Critical", "Critical_Count"),
        ("Major", "Major_Count"), ("Minor", "Minor_Count"),
    ]:
        if not label_count(text, label, summary[key], 90):
            errors.append(f"缺 {label}={summary[key]}")
    for row in truth["refund_reason_breakdown"]:
        if not label_count(text, row["Refund_Reason"], row["Refund_Count"], 45):
            errors.append(f"缺原因 {row}")
    for row in truth["root_cause_breakdown"]:
        if not label_count(text, row["Root_Cause"], row["Refund_Count"], 45):
            errors.append(f"缺根因 {row}")
    for row in truth["investigation_products"]:
        if row["Product_Name"] not in text or not has_money(text, row["Total_Refund_Amount"]):
            errors.append(f"缺调查商品 {row['Product_Name'][:45]}")
    return not errors, errors


def email_checks(quality: str, supplier: str, truth: dict) -> tuple[bool, list[str]]:
    summary = truth["summary"]
    errors = []
    q_lower = quality.lower()
    s_lower = supplier.lower()
    if "to: quality_team@company.com" not in q_lower:
        errors.append("质量邮件收件人错误")
    q_subject = re.search(r"^subject:\s*(.+)$", quality, re.IGNORECASE | re.MULTILINE)
    if not q_subject or not all(token in q_subject.group(1).lower() for token in ("refund", "analysis")):
        errors.append("质量邮件主题错误")
    if not has_money(quality, summary["Total_Refund_Amount"]):
        errors.append("质量邮件总金额错误")
    for label, key in [("Critical", "Critical_Count"), ("Major", "Major_Count"), ("Minor", "Minor_Count")]:
        if not label_count(quality, label, summary[key], 30):
            errors.append(f"质量邮件缺 {label}")
    if not label_count(quality, "refund", summary["Total_Refunds"], 90):
        errors.append("质量邮件退款数错误")

    if "to: supplier_relations@company.com" not in s_lower:
        errors.append("供应商邮件收件人错误")
    s_subject = re.search(r"^subject:\s*(.+)$", supplier, re.IGNORECASE | re.MULTILINE)
    if not s_subject or not all(token in s_subject.group(1).lower() for token in ("supplier", "investigation")):
        errors.append("供应商邮件主题错误")
    for row in truth["investigation_products"]:
        if row["Product_Name"] not in supplier:
            errors.append(f"供应商邮件缺商品 {row['Product_Name'][:45]}")
        product_position = supplier.find(row["Product_Name"])
        product_block = supplier[product_position : product_position + len(row["Product_Name"]) + 140]
        if not label_count(product_block, "refunds", row["Refund_Count"], 30):
            errors.append(f"供应商邮件退款数错误 {row['Product_Name'][:30]}")
        if not has_money(product_block, row["Total_Refund_Amount"]):
            errors.append(f"供应商邮件金额错误 {row['Product_Name'][:30]}")
    return not errors, errors


def unsupported_claims(text: str, truth: dict) -> list[str]:
    allowed = {Decimal("20.00"), Decimal("50.00")}
    allowed.add(round_money(Decimal(str(truth["summary"]["Total_Refund_Amount"]))))
    allowed.update(round_money(Decimal(str(row["Refund_Amount"]))) for row in truth["refund_details"])
    allowed.update(round_money(Decimal(str(row["Total_Refund_Amount"]))) for row in truth["product_impact"])
    findings = []
    for match in re.finditer(r"\$\s*([0-9][0-9,]*(?:\.\d{1,2})?)", text):
        value = round_money(Decimal(match.group(1).replace(",", "")))
        if value not in allowed:
            findings.append(f"无来源金额 ${match.group(1)}")
    if "%" in text:
        findings.append("输入不含百分比事实")
    lowered = text.lower()
    for pattern, label in [
        (r"industry\s+benchmark|industry\s+average", "行业基准"),
        (r"revenue\s+impact|revenue\s+loss|cost\s+savings|\broi\b", "业务效果金额"),
        (r"customer\s+loss|churn\s+rate", "客户流失"),
        (r"supplier\s+(?:fault|defect)\s+rate", "供应商缺陷率"),
    ]:
        for match in re.finditer(pattern, lowered):
            prefix = lowered[max(0, match.start() - 100) : match.start()]
            if re.search(
                r"(?:do|does)\s+not\s+support|not\s+supported|not\s+calculated|"
                r"not\s+estimated|no\s+[^.]{0,50}(?:is|are)\s+(?:asserted|calculated|estimated)",
                prefix,
            ):
                continue
            findings.append(label)
            break
    return findings


def render_office(paths: list[Path], target: Path) -> tuple[dict[str, dict[str, Any]], str]:
    rendered: dict[str, dict[str, Any]] = {}
    logs = []
    target.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["HOME"] = str(target / "lo-home")
    env["NO_PROXY"] = "localhost,127.0.0.1,::1"
    for path in paths:
        try:
            completed = subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(target), str(path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=120,
                check=False,
                env=env,
            )
        except Exception as error:
            logs.append(f"{path.name}: {error}")
            continue
        pdf_path = target / f"{path.stem}.pdf"
        logs.append(f"{path.name}: rc={completed.returncode}, pdf={pdf_path.is_file()}, {completed.stdout[-240:]}")
        if completed.returncode != 0 or not pdf_path.is_file():
            continue
        info = subprocess.run(["pdfinfo", str(pdf_path)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, check=False)
        page_match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
        text_result = subprocess.run(["pdftotext", "-layout", str(pdf_path), "-"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, check=False)
        rendered[path.name] = {
            "pages": int(page_match.group(1)) if page_match else 0,
            "text": text_result.stdout,
            "pdf": str(pdf_path),
        }
    return rendered, " | ".join(logs)


def build_result(results: list[dict], artifact_paths: list[Path]) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in CRITERION_SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "因交付缺失或前置门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])
    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdle_failures = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
    if hurdle_failures:
        reward = min(reward, min(HURDLE_CAPS[item] for item in hurdle_failures))
    digests = [hashlib.sha256(path.read_bytes()).hexdigest() for path in artifact_paths if path.is_file()]
    if not digests:
        digests = [hashlib.sha256(b"").hexdigest()]
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


def evaluate(artifact_dir: Path) -> dict:
    task_root = Path(__file__).resolve().parents[1]
    input_dir = task_root / "data" / "input_files"
    paths = {name: artifact_dir / name for name in OUTPUT_NAMES}
    results = []
    missing = [name for name, path in paths.items() if not path.is_file()]
    structure_ok = not missing
    results.append(criterion("R001", structure_ok, f"缺失={missing}"))
    if missing:
        results.append(criterion("R016", True, f"缺失={missing}"))
        return build_result(results, list(paths.values()))

    truth = derive_truth(input_dir)
    canonical_payload, canonical_log = run_processor(paths["refund_root_cause_processor.py"], input_dir)
    processor_ok = canonical_payload is not None and payload_equal(canonical_payload, truth)
    results.append(criterion("R002", processor_ok, f"匹配={processor_ok}，日志={canonical_log}"))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-refund-mutation-") as temp_dir:
        mutated_input = Path(temp_dir) / "inputs"
        build_mutated_inputs(input_dir, mutated_input)
        mutated_truth = derive_truth(mutated_input)
        mutated_payload, mutated_log = run_processor(paths["refund_root_cause_processor.py"], mutated_input)
    changed = (
        mutated_truth["summary"]["Total_Refund_Amount"] != truth["summary"]["Total_Refund_Amount"]
        and mutated_truth["summary"]["Critical_Count"] != truth["summary"]["Critical_Count"]
        and mutated_truth["summary"]["Quality_Issues"] != truth["summary"]["Quality_Issues"]
        and mutated_truth["product_impact"][0]["Avg_Review_Rating"] != truth["product_impact"][0]["Avg_Review_Rating"]
    )
    mutation_ok = mutated_payload is not None and payload_equal(mutated_payload, mutated_truth) and changed
    results.append(
        criterion(
            "R003",
            mutation_ok,
            f"匹配={mutated_payload is not None and payload_equal(mutated_payload, mutated_truth)}，真值变化={changed}，日志={mutated_log}",
        )
    )

    try:
        delivered_json = json.loads(paths["analysis_results.json"].read_text(encoding="utf-8"))
        json_ok = payload_equal(delivered_json, truth)
        json_error = ""
    except Exception as error:
        delivered_json = None
        json_ok = False
        json_error = str(error)
    results.append(criterion("R004", json_ok, f"匹配={json_ok}，错误={json_error}"))

    excel_payload, formatting, excel_error = read_excel(paths["Refund_Analysis.xlsx"])
    details_ok = bool(excel_payload) and rows_equal(excel_payload["details"], truth["refund_details"])
    products_ok = bool(excel_payload) and rows_equal(excel_payload["products"], truth["product_impact"])
    summary_ok = bool(excel_payload) and set(excel_payload["summary"]) == set(SUMMARY_KEYS) and all(
        payload_equal(excel_payload["summary"][key], truth["summary"][key]) for key in SUMMARY_KEYS
    )
    results.append(criterion("R005", details_ok, f"行数={len(excel_payload['details']) if excel_payload else 0}，错误={excel_error}"))
    results.append(criterion("R006", products_ok, f"行数={len(excel_payload['products']) if excel_payload else 0}，错误={excel_error}"))
    results.append(criterion("R007", summary_ok, f"键={list(excel_payload['summary']) if excel_payload else []}，错误={excel_error}"))

    try:
        report_text, report_headings, report_tables, report_media = docx_payload(paths["Root_Cause_Report.docx"])
        report_ok, report_errors = word_checks(report_text, report_headings, truth)
    except Exception as error:
        report_text, report_headings, report_tables, report_media = "", [], 0, 0
        report_ok, report_errors = False, [str(error)]
    results.append(criterion("R008", report_ok, f"错误={report_errors}，标题={len(report_headings)}，表={report_tables}"))

    try:
        quality_text = paths["quality_team_email.txt"].read_text(encoding="utf-8")
        supplier_text = paths["supplier_relations_email.txt"].read_text(encoding="utf-8")
        emails_ok, email_errors = email_checks(quality_text, supplier_text, truth)
    except Exception as error:
        quality_text, supplier_text = "", ""
        emails_ok, email_errors = False, [str(error)]
    results.append(criterion("R009", emails_ok, f"错误={email_errors}"))

    cross_ok = json_ok and details_ok and products_ok and summary_ok and report_ok and emails_ok
    results.append(criterion("R010", cross_ok, f"JSON={json_ok}，Excel={details_ok}/{products_ok}/{summary_ok}，Word={report_ok}，Emails={emails_ok}"))

    native_ok = (
        formatting.get("images") == 0
        and formatting.get("tables", 0) >= 3
        and formatting.get("freezes") == {"Refund Details": "A2", "Product Impact": "A2", "Summary": "A2"}
        and formatting.get("header_bold")
        and formatting.get("column_widths")
        and report_media == 0
        and report_tables >= 4
        and len(report_headings) >= 7
    )
    results.append(criterion("R011", native_ok, json.dumps(formatting | {"doc_tables": report_tables, "doc_media": report_media, "doc_headings": len(report_headings)}, ensure_ascii=False)))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-refund-render-") as temp_dir:
        rendered, render_log = render_office(
            [paths["Refund_Analysis.xlsx"], paths["Root_Cause_Report.docx"]], Path(temp_dir)
        )
    xlsx_render = rendered.get("Refund_Analysis.xlsx", {})
    docx_render = rendered.get("Root_Cause_Report.docx", {})
    render_ok = (
        xlsx_render.get("pages", 0) >= 3
        and all(token in xlsx_render.get("text", "") for token in ("Refund_ID", "Product_Name", "Total_Refunds"))
        and docx_render.get("pages", 0) >= 3
        and all(token in docx_render.get("text", "") for token in ("Executive Summary", "Product-Level Findings", "Action Plan"))
    )
    render_pages = {"xlsx": xlsx_render.get("pages", 0), "docx": docx_render.get("pages", 0)}
    results.append(criterion("R012", render_ok, f"渲染={render_pages}，日志={render_log}"))

    unsupported = unsupported_claims(report_text + "\n" + quality_text + "\n" + supplier_text, truth)
    no_unsupported = not unsupported
    results.append(criterion("R013", no_unsupported, f"发现={unsupported}"))

    source_error = not (processor_ok and json_ok and details_ok and products_ok and summary_ok and report_ok and emails_ok)
    script_error = not (processor_ok and mutation_ok)
    damaged_error = not (structure_ok and excel_payload is not None and bool(report_text) and render_ok)
    conflict_error = not cross_ok
    unsupported_error = bool(unsupported)
    results.append(criterion("R014", source_error, f"处理器={processor_ok}，JSON={json_ok}，Excel={details_ok}/{products_ok}/{summary_ok}，Word={report_ok}，Emails={emails_ok}"))
    results.append(criterion("R015", script_error, f"原始回放={processor_ok}，变异回放={mutation_ok}"))
    results.append(criterion("R016", damaged_error, f"结构={structure_ok}，Excel={excel_payload is not None}，Word={bool(report_text)}，渲染={render_ok}"))
    results.append(criterion("R017", conflict_error, f"跨文件一致={cross_ok}"))
    results.append(criterion("R018", unsupported_error, f"发现={unsupported}"))
    return build_result(results, list(paths.values()))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir.resolve())
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
