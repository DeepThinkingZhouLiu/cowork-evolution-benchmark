#!/usr/bin/env python3
"""FORTE Finance-018 多工作簿整合任务的独立确定性 Judge。"""

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
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

sys.dont_write_bytecode = True

from openpyxl import load_workbook
from pypinyin import lazy_pinyin


TASK_ID = "cowork-evo/forte-finance-018-receivables-payables-consolidation"
RUBRIC_VERSION = "forte-finance-018-v0.1"
JUDGE_VERSION = "1.0.0"
PASS_THRESHOLD = 0.8
OUTPUT_NAMES = [
    "reconcile_workbooks.py",
    "未付统计.csv",
    "未收统计.csv",
    "僵尸账款核查.txt",
    "往来款跨期整合.xlsx",
]
PERIOD_FILES = (
    ("2025上半年", "2025往来明细-上半年.xlsx"),
    ("2025下半年", "2025往来明细-下半年.xlsx"),
    ("2026", "2026往来明细.xlsx"),
)
EXPECTED_HEADERS = (
    "科目名称",
    "客商名称",
    "方向",
    "期初余额",
    "本期借方",
    "本期贷方",
    "借方累计",
    "贷方累计",
    "方向",
    "期末余额",
)
REQUIRED_SHEETS = ["跨期明细", "未付统计", "未收统计", "僵尸账款核查", "校验摘要"]
INPUT_HASHES = {
    "2025往来明细-上半年.xlsx": (13873, "4ffb7c2d6d3efe6fcae45c2aff755bedd727d9c3467d00e1765a147e5ecab5f3"),
    "2025往来明细-下半年.xlsx": (12850, "4ad841ee36acc3b52aea0c6931a260b69ead7493f0c33056c429c36d70575c31"),
    "2026往来明细.xlsx": (12746, "4ee1f69d3346eea3b99660ab8a2efc946c93f00123ba83e121277bc7be66f4f3"),
}
UPSTREAM_HASHES = {
    "upstream-task.md": "01506a107ae61814e220cd6aca06605172be3e8687005fc1e275813120baf393",
    "upstream-rubrics.md": "5f1b1f7af5e06d9de60dfc420e3235eda1ebe1d94f8163682752321e2385f61c",
    "upstream-unpaid-golden.csv": "ecc14fd234ebecb51380e3deab5e459871f777d0f9de858df546f57c15eb10d5",
    "upstream-unreceived-golden.csv": "8d4a9c0e988afa7099512e9ed0c6eb041bc0a11da7baa286041e2d2c81c8d6d3",
    "LICENSE-FORTE": "0746b4b0c31d032479c7fdf1aa17566eb1978392126bcde26eefcfceb6893f9c",
}
CENT = Decimal("0.01")

CRITERION_SPECS = {
    "R001": (8, "hurdle", "5 个指定交付物完整可解析"),
    "R002": (12, "hurdle", "未付 CSV 与 FORTE 上游 Golden 一致"),
    "R003": (12, "hurdle", "未收 CSV 与 FORTE 上游 Golden 一致"),
    "R004": (15, "hurdle", "业务真值与僵尸账结论正确"),
    "R005": (15, "hurdle", "processor 原始输入 fresh replay 正确"),
    "R006": (10, "hurdle", "processor 内容变异 replay 正确"),
    "R007": (8, "hurdle", "整合工作簿完整且跨交付物一致"),
    "R008": (7, "hurdle", "表头、排序和业务口径一致"),
    "R009": (5, "hurdle", "输入与上游审计文件哈希正确"),
    "R010": (4, "positive", "LibreOffice 可导出非空 PDF"),
    "R011": (4, "positive", "工作簿保持原生可编辑结构"),
    "R012": (-25, "penalty", "交付缺失、损坏或不可解析"),
    "R013": (-25, "penalty", "未付或未收业务结果错误"),
    "R014": (-20, "penalty", "僵尸账结论或业务断言错误"),
    "R015": (-25, "penalty", "processor 不可回放或写死原始答案"),
    "R016": (-18, "penalty", "XLSX 与 CSV/文本互相冲突"),
    "R017": (-15, "penalty", "跨期明细或原生表格结构丢失"),
}
HURDLE_CAPS = {
    "R001": 0.0,
    "R002": 0.45,
    "R003": 0.45,
    "R004": 0.40,
    "R005": 0.30,
    "R006": 0.35,
    "R007": 0.55,
    "R008": 0.65,
    "R009": 0.50,
}


@dataclass(frozen=True)
class Record:
    period: str
    source_file: str
    source_row: int
    subject: str
    counterparty: str
    opening_direction: str
    opening_balance: Decimal | None
    debit: Decimal | None
    credit: Decimal | None
    debit_total: Decimal | None
    credit_total: Decimal | None
    closing_direction: str
    closing_balance: Decimal | None
    closing_number_format: str


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def decimal(value: object) -> Decimal | None:
    if value in (None, ""):
        return None
    return Decimal(str(value)).quantize(CENT, rounding=ROUND_HALF_UP)


def read_records(input_dir: Path) -> tuple[list[Record], list[int]]:
    records: list[Record] = []
    counts: list[int] = []
    for period, filename in PERIOD_FILES:
        workbook = load_workbook(input_dir / filename, data_only=True, read_only=False)
        try:
            sheet = workbook.active
            headers = tuple(cell.value for cell in sheet[1][:10])
            if headers != EXPECTED_HEADERS:
                raise ValueError(f"{filename} 表头错误：{headers}")
            count = 0
            for row_number, row in enumerate(sheet.iter_rows(min_row=2, max_col=10), 2):
                values = [cell.value for cell in row]
                if not any(value is not None for value in values):
                    continue
                records.append(
                    Record(
                        period,
                        filename,
                        row_number,
                        str(values[0]),
                        str(values[1]),
                        str(values[2] or ""),
                        decimal(values[3]),
                        decimal(values[4]),
                        decimal(values[5]),
                        decimal(values[6]),
                        decimal(values[7]),
                        str(values[8] or ""),
                        decimal(values[9]),
                        row[9].number_format,
                    )
                )
                count += 1
            counts.append(count)
        finally:
            workbook.close()
    return records, counts


def sort_key(record: Record) -> tuple[tuple[str, ...], Decimal, str]:
    return tuple(lazy_pinyin(record.counterparty)), -(record.closing_balance or Decimal("0")), record.subject


def truth(records: list[Record]) -> dict[str, Any]:
    current = [record for record in records if record.period == "2026"]
    unpaid = sorted(
        [record for record in current if record.closing_direction == "贷" and record.closing_balance not in (None, 0)],
        key=sort_key,
    )
    unreceived = sorted(
        [record for record in current if record.closing_direction == "借" and record.closing_balance not in (None, 0)],
        key=sort_key,
    )
    history: dict[tuple[str, str], dict[str, Record]] = {}
    for record in records:
        history.setdefault((record.subject, record.counterparty), {})[record.period] = record
    audit = []
    for item in unreceived:
        trace = history[(item.subject, item.counterparty)]
        period_rows = [trace.get(period) for period, _ in PERIOD_FILES]
        balances = [record.closing_balance if record else None for record in period_rows]
        directions = [record.closing_direction if record else None for record in period_rows]
        audit.append(
            {
                "subject": item.subject,
                "counterparty": item.counterparty,
                "balances": balances,
                "is_zombie": (
                    all(value not in (None, Decimal("0")) for value in balances)
                    and all(value == "借" for value in directions)
                    and len(set(balances)) == 1
                ),
            }
        )
    return {"unpaid": unpaid, "unreceived": unreceived, "audit": audit}


def csv_amount(record: Record) -> str:
    value = record.closing_balance
    if value is None:
        return ""
    if record.closing_number_format == "General" and value == value.to_integral_value():
        return str(int(value))
    return f"{value:,.2f}"


def expected_csv_rows(records: Iterable[Record], header: str) -> list[list[str]]:
    return [["科目名称", "客商名称", header]] + [
        [record.subject, record.counterparty, csv_amount(record)] for record in records
    ]


def read_csv(path: Path) -> tuple[list[list[str]], str]:
    payload = path.read_bytes()
    encoding = "gb18030"
    try:
        text = payload.decode(encoding)
    except UnicodeDecodeError:
        encoding = "utf-8-sig"
        text = payload.decode(encoding)
    return list(csv.reader(text.splitlines())), encoding


def inspect_business_outputs(artifact_dir: Path, input_dir: Path) -> tuple[dict[str, bool], dict[str, Any]]:
    records, counts = read_records(input_dir)
    expected = truth(records)
    unpaid_rows, unpaid_encoding = read_csv(artifact_dir / "未付统计.csv")
    unreceived_rows, unreceived_encoding = read_csv(artifact_dir / "未收统计.csv")
    unpaid_ok = unpaid_rows == expected_csv_rows(expected["unpaid"], "未付款项")
    unreceived_ok = unreceived_rows == expected_csv_rows(expected["unreceived"], "未收款项")
    text = (artifact_dir / "僵尸账款核查.txt").read_text(encoding="utf-8")
    zombies = [item for item in expected["audit"] if item["is_zombie"]]
    conclusion_ok = ("存在僵尸账款" in text) if zombies else ("无僵尸账款" in text and "存在僵尸账款" not in text)
    zombie_facts_ok = True
    for item in zombies:
        amount = item["balances"][-1]
        zombie_facts_ok &= item["counterparty"] in text and item["subject"] in text and f"{amount:,.2f}" in text
    bullet_lines = [line.strip() for line in text.splitlines() if line.strip().startswith("-")]
    allowed_bullets = {
        f"- {item['counterparty']} | {item['subject']} | {item['balances'][-1]:,.2f}" for item in zombies
    }
    unsupported_ok = set(bullet_lines) == allowed_bullets
    return (
        {
            "unpaid": unpaid_ok,
            "unreceived": unreceived_ok,
            "zombie": conclusion_ok and zombie_facts_ok and unsupported_ok,
            "sort_and_headers": unpaid_ok and unreceived_ok,
        },
        {
            "records": records,
            "counts": counts,
            "truth": expected,
            "unpaid_encoding": unpaid_encoding,
            "unreceived_encoding": unreceived_encoding,
            "text": text,
            "bullet_lines": bullet_lines,
        },
    )


def inspect_workbook(path: Path, context: dict[str, Any]) -> tuple[bool, bool, str]:
    workbook = load_workbook(path, data_only=True, read_only=False)
    errors: list[str] = []
    native_errors: list[str] = []
    expected = context["truth"]
    records: list[Record] = context["records"]
    try:
        if workbook.sheetnames != REQUIRED_SHEETS:
            errors.append(f"工作表={workbook.sheetnames}")
            return False, False, "; ".join(errors)
        detail = workbook["跨期明细"]
        expected_detail = [
            [
                record.period,
                record.source_file,
                record.source_row,
                record.subject,
                record.counterparty,
                record.opening_direction,
                record.opening_balance,
                record.debit,
                record.credit,
                record.debit_total,
                record.credit_total,
                record.closing_direction,
                record.closing_balance,
            ]
            for record in records
        ]
        actual_detail = []
        for row in detail.iter_rows(min_row=2, values_only=True):
            actual_detail.append(
                [
                    decimal(value) if index in (6, 7, 8, 9, 10, 12) else value
                    for index, value in enumerate(row[:13])
                ]
            )
        if actual_detail != expected_detail:
            errors.append("跨期明细与 3 份输入不一致")

        for sheet_name, key, amount_header in (
            ("未付统计", "unpaid", "未付款项"),
            ("未收统计", "unreceived", "未收款项"),
        ):
            sheet = workbook[sheet_name]
            header = [cell.value for cell in sheet[1][:3]]
            rows = [[row[0], row[1], decimal(row[2])] for row in sheet.iter_rows(min_row=2, max_col=3, values_only=True)]
            wanted = [[r.subject, r.counterparty, r.closing_balance] for r in expected[key]]
            if header != ["科目名称", "客商名称", amount_header] or rows != wanted:
                errors.append(f"{sheet_name} 与 CSV 真值不一致")

        zombie_sheet = workbook["僵尸账款核查"]
        actual_audit = []
        for row in zombie_sheet.iter_rows(min_row=2, max_col=6, values_only=True):
            actual_audit.append([row[0], row[1], decimal(row[2]), decimal(row[3]), decimal(row[4]), row[5]])
        wanted_audit = [
            [item["subject"], item["counterparty"], *item["balances"], "是" if item["is_zombie"] else "否"]
            for item in expected["audit"]
        ]
        if actual_audit != wanted_audit:
            errors.append("僵尸账轨迹与独立重算不一致")

        summary = {row[0]: row[1] for row in workbook["校验摘要"].iter_rows(min_row=2, max_col=2, values_only=True)}
        expected_summary = {
            "原始记录总数": len(records),
            "2026 未付笔数": len(expected["unpaid"]),
            "2026 未付总额": sum((r.closing_balance or Decimal("0")) for r in expected["unpaid"]),
            "2026 未收笔数": len(expected["unreceived"]),
            "2026 未收总额": sum((r.closing_balance or Decimal("0")) for r in expected["unreceived"]),
            "僵尸账笔数": sum(bool(item["is_zombie"]) for item in expected["audit"]),
        }
        for key, value in expected_summary.items():
            actual = summary.get(key)
            if isinstance(value, Decimal):
                actual = decimal(actual)
            if actual != value:
                errors.append(f"校验摘要 {key}={actual!r} != {value!r}")

        for sheet in workbook.worksheets:
            if sheet.max_row > 1 and sheet.freeze_panes != "A2":
                native_errors.append(f"{sheet.title} 未冻结表头")
            if sheet.max_row > 1 and not sheet.auto_filter.ref:
                native_errors.append(f"{sheet.title} 未设置筛选")
        if getattr(workbook, "_images", None):
            native_errors.append("工作簿包含整体图片代替表格")
    finally:
        workbook.close()
    return not errors, not native_errors, f"内容问题={errors[:8]}，原生结构问题={native_errors[:8]}"


def run_processor(processor: Path, input_dir: Path, output_dir: Path) -> tuple[bool, str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    try:
        completed = subprocess.run(
            [sys.executable, str(processor), "--input-dir", str(input_dir), "--output-dir", str(output_dir)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=120,
            check=False,
            env=env,
        )
    except Exception as error:
        return False, str(error)
    ok = completed.returncode == 0 and all((output_dir / name).is_file() for name in OUTPUT_NAMES)
    return ok, completed.stdout[-1000:]


def build_mutation(source_dir: Path, target_dir: Path) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)
    for _, filename in PERIOD_FILES:
        shutil.copy2(source_dir / filename, target_dir / filename)
    path = target_dir / "2026往来明细.xlsx"
    workbook = load_workbook(path)
    sheet = workbook.active
    found = False
    for row in sheet.iter_rows(min_row=2, max_col=10):
        if row[0].value == "其他应收款\其他应收往来" and row[1].value == "【绵阳长城发展融资担保有限公司】":
            row[4].value = None
            row[6].value = None
            row[9].value = 1500000
            row[9].number_format = "#,##0.00"
            found = True
            break
    if not found:
        raise AssertionError("未找到变异目标应收款")
    sheet.append(
        [
            "应付账款\应付其他",
            "【北京星河科技有限公司】",
            "平",
            None,
            None,
            12345.67,
            None,
            12345.67,
            "贷",
            12345.67,
        ]
    )
    for column in (6, 8, 10):
        sheet.cell(sheet.max_row, column).number_format = "#,##0.00"
    workbook.save(path)
    workbook.close()


def inspect_source_package(task_root: Path) -> tuple[bool, str]:
    errors = []
    manifest = json.loads((task_root / "metadata" / "input-manifest.json").read_text(encoding="utf-8"))
    manifest_map = {Path(item["path"]).name: item for item in manifest["files"]}
    for filename, (size, digest) in INPUT_HASHES.items():
        path = task_root / "data" / "input_files" / filename
        if path.stat().st_size != size or sha256(path) != digest:
            errors.append(filename)
        if manifest_map.get(filename, {}).get("sha256") != digest or manifest_map.get(filename, {}).get("size_bytes") != size:
            errors.append(f"manifest:{filename}")
    for filename, digest in UPSTREAM_HASHES.items():
        path = task_root / "adaptation" / filename
        if not path.is_file() or sha256(path) != digest:
            errors.append(filename)
    return not errors, f"错误={errors}"


def render_workbook(path: Path) -> tuple[bool, str]:
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    pdfinfo = shutil.which("pdfinfo")
    if not executable or not pdfinfo:
        return False, "缺少 libreoffice 或 pdfinfo"
    with tempfile.TemporaryDirectory(prefix="cowork-evo-forte-render-") as directory:
        target = Path(directory)
        env = os.environ.copy()
        env["HOME"] = str(target / "home")
        completed = subprocess.run(
            [executable, "--headless", "--convert-to", "pdf", "--outdir", str(target), str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=120,
            check=False,
            env=env,
        )
        pdf = target / f"{path.stem}.pdf"
        if completed.returncode != 0 or not pdf.is_file() or pdf.stat().st_size < 1000:
            return False, f"rc={completed.returncode}，日志={completed.stdout[-500:]}"
        info = subprocess.run([pdfinfo, str(pdf)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, check=False)
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
        pages = int(match.group(1)) if match else 0
        return pages > 0, f"PDF={pdf.stat().st_size} bytes，pages={pages}"


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


def build_result(results: list[dict[str, Any]], paths: list[Path]) -> dict[str, Any]:
    results.sort(key=lambda item: item["criterion_id"])
    total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / total))
    hurdle_failures = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
    if hurdle_failures:
        reward = min(reward, min(HURDLE_CAPS[item] for item in hurdle_failures))
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
    digests = [sha256(path) for path in paths if path.is_file()]
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": digests or [hashlib.sha256(b"").hexdigest()],
        "criterion_results": criterion_results,
        "reward": round(reward, 6),
        "passed": reward >= PASS_THRESHOLD,
        "hurdle_failures": hurdle_failures,
    }


def evaluate(artifact_dir: Path, render: bool = True) -> dict[str, Any]:
    task_root = Path(__file__).resolve().parents[1]
    input_dir = task_root / "data" / "input_files"
    paths = {name: artifact_dir / name for name in OUTPUT_NAMES}
    results: list[dict[str, Any]] = []

    missing = [name for name, path in paths.items() if not path.is_file() or path.stat().st_size == 0]
    integrity_ok = not missing
    if integrity_ok:
        try:
            compile(paths["reconcile_workbooks.py"].read_text(encoding="utf-8"), str(paths["reconcile_workbooks.py"]), "exec")
            read_csv(paths["未付统计.csv"])
            read_csv(paths["未收统计.csv"])
            paths["僵尸账款核查.txt"].read_text(encoding="utf-8")
            with zipfile.ZipFile(paths["往来款跨期整合.xlsx"]) as archive:
                integrity_ok = archive.testzip() is None
        except Exception as error:
            integrity_ok = False
            missing.append(f"解析失败：{error}")
    results.append(criterion("R001", integrity_ok, f"缺失或损坏={missing}"))

    if not integrity_ok:
        for identifier in ("R002", "R003", "R004", "R005", "R006", "R007", "R008", "R009", "R010", "R011"):
            results.append(criterion(identifier, False, "前置完整性失败"))
        results.extend(
            [
                criterion("R012", True, f"缺失或损坏={missing}"),
                criterion("R013", True, "无法验证业务结果"),
                criterion("R014", True, "无法验证核查结论"),
                criterion("R015", True, "无法执行 processor"),
                criterion("R016", True, "无法验证跨文件一致性"),
                criterion("R017", True, "无法验证工作簿结构"),
            ]
        )
        return build_result(results, list(paths.values()))

    checks, context = inspect_business_outputs(artifact_dir, input_dir)
    unpaid_bytes = paths["未付统计.csv"].read_bytes() == (task_root / "adaptation" / "upstream-unpaid-golden.csv").read_bytes()
    unreceived_bytes = paths["未收统计.csv"].read_bytes() == (task_root / "adaptation" / "upstream-unreceived-golden.csv").read_bytes()
    r002 = checks["unpaid"] and unpaid_bytes
    r003 = checks["unreceived"] and unreceived_bytes
    results.append(criterion("R002", r002, f"内容={checks['unpaid']}，逐字节={unpaid_bytes}，编码={context['unpaid_encoding']}"))
    results.append(criterion("R003", r003, f"内容={checks['unreceived']}，逐字节={unreceived_bytes}，编码={context['unreceived_encoding']}"))
    canonical_truth = (
        context["counts"] == [74, 59, 58]
        and len(context["truth"]["unpaid"]) == 31
        and sum((r.closing_balance or Decimal("0")) for r in context["truth"]["unpaid"]) == Decimal("3984606.46")
        and len(context["truth"]["unreceived"]) == 2
        and sum((r.closing_balance or Decimal("0")) for r in context["truth"]["unreceived"]) == Decimal("4992891.47")
        and not any(item["is_zombie"] for item in context["truth"]["audit"])
    )
    r004 = checks["unpaid"] and checks["unreceived"] and checks["zombie"] and canonical_truth
    results.append(criterion("R004", r004, f"counts={context['counts']}，unpaid={len(context['truth']['unpaid'])}，unreceived={len(context['truth']['unreceived'])}，zombie={checks['zombie']}"))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-forte-replay-") as directory:
        base = Path(directory)
        fresh_dir = base / "fresh"
        fresh_ran, fresh_log = run_processor(paths["reconcile_workbooks.py"], input_dir, fresh_dir)
        fresh_ok = False
        if fresh_ran:
            fresh_checks, fresh_context = inspect_business_outputs(fresh_dir, input_dir)
            fresh_workbook, fresh_native, _ = inspect_workbook(fresh_dir / "往来款跨期整合.xlsx", fresh_context)
            fresh_ok = all(fresh_checks.values()) and fresh_workbook and fresh_native
        results.append(criterion("R005", fresh_ran and fresh_ok, f"run={fresh_ran}，outputs={fresh_ok}，log={fresh_log}"))

        mutated_inputs = base / "mutated-inputs"
        build_mutation(input_dir, mutated_inputs)
        mutation_dir = base / "mutation-output"
        mutation_ran, mutation_log = run_processor(paths["reconcile_workbooks.py"], mutated_inputs, mutation_dir)
        mutation_ok = False
        mutation_evidence = ""
        if mutation_ran:
            mutation_checks, mutation_context = inspect_business_outputs(mutation_dir, mutated_inputs)
            mutation_workbook, mutation_native, mutation_book_evidence = inspect_workbook(
                mutation_dir / "往来款跨期整合.xlsx", mutation_context
            )
            mutation_truth = mutation_context["truth"]
            zombies = [item for item in mutation_truth["audit"] if item["is_zombie"]]
            mutation_ok = (
                all(mutation_checks.values())
                and mutation_workbook
                and mutation_native
                and len(mutation_truth["unpaid"]) == 32
                and sum((r.closing_balance or Decimal("0")) for r in mutation_truth["unpaid"]) == Decimal("3996952.13")
                and len(mutation_truth["unreceived"]) == 2
                and sum((r.closing_balance or Decimal("0")) for r in mutation_truth["unreceived"]) == Decimal("4792891.47")
                and len(zombies) == 1
                and zombies[0]["counterparty"] == "【绵阳长城发展融资担保有限公司】"
            )
            mutation_evidence = f"unpaid={len(mutation_truth['unpaid'])}，unreceived={len(mutation_truth['unreceived'])}，zombies={len(zombies)}，book={mutation_book_evidence}"
        results.append(criterion("R006", mutation_ran and mutation_ok, f"run={mutation_ran}，{mutation_evidence}，log={mutation_log}"))

    workbook_ok, native_ok, workbook_evidence = inspect_workbook(paths["往来款跨期整合.xlsx"], context)
    results.append(criterion("R007", workbook_ok, workbook_evidence))
    results.append(criterion("R008", checks["sort_and_headers"], f"未付={checks['unpaid']}，未收={checks['unreceived']}"))
    source_ok, source_evidence = inspect_source_package(task_root)
    results.append(criterion("R009", source_ok, source_evidence))
    render_ok, render_evidence = render_workbook(paths["往来款跨期整合.xlsx"]) if render else (True, "测试模式跳过渲染")
    results.append(criterion("R010", render_ok, render_evidence))
    results.append(criterion("R011", native_ok, workbook_evidence))

    results.extend(
        [
            criterion("R012", False, "交付完整可解析，未触发"),
            criterion("R013", not (r002 and r003 and checks["sort_and_headers"]), f"R002={r002}，R003={r003}"),
            criterion("R014", not checks["zombie"], f"zombie_check={checks['zombie']}，bullets={context['bullet_lines']}"),
            criterion("R015", not (fresh_ran and fresh_ok and mutation_ran and mutation_ok), f"fresh={fresh_ok}，mutation={mutation_ok}"),
            criterion("R016", not workbook_ok, workbook_evidence),
            criterion("R017", not native_ok or len(context["records"]) != 191, workbook_evidence),
        ]
    )
    return build_result(results, list(paths.values()))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", required=True, type=Path)
    parser.add_argument("--skip-render", action="store_true")
    parser.add_argument("--json-output", type=Path)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir.resolve(), render=not args.skip_render)
    payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    result_path = args.result or args.json_output
    if result_path:
        result_path.parent.mkdir(parents=True, exist_ok=True)
        result_path.write_text(payload, encoding="utf-8")
    if args.reward_file:
        args.reward_file.parent.mkdir(parents=True, exist_ok=True)
        args.reward_file.write_text(f"{result['reward']:.6f}\n", encoding="utf-8")
    print(payload, end="")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
