#!/usr/bin/env python3
"""AutomationBench 5104 工资差异控制包的逐题确定性 Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import mailbox
import re
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook


TASK_ID = "cowork-evo/automationbench-5104-payroll-discrepancy-control-package"
TASK_ROOT = Path(__file__).resolve().parents[1]
HARBOR_INPUT_DIR = Path("/app/input_files")
INPUT_DIR = HARBOR_INPUT_DIR if HARBOR_INPUT_DIR.is_dir() else TASK_ROOT / "data" / "input_files"
EXPECTED_FILES = {
    "Payroll_Discrepancy_Audit.xlsx",
    "Payroll_Control_Brief.docx",
    "Payroll_Routing_Drafts.mbox",
    "Payroll_Channel_Draft.json",
    "Back_Pay_Execution_Control.json",
}
XLSX_NAME = "Payroll_Discrepancy_Audit.xlsx"
DOCX_NAME = "Payroll_Control_Brief.docx"
MBOX_NAME = "Payroll_Routing_Drafts.mbox"
SLACK_NAME = "Payroll_Channel_Draft.json"
CONTROL_NAME = "Back_Pay_Execution_Control.json"
INPUT_HASHES = {
    "Payroll_Discrepancy_Tracker.xlsx": "607a07c9e9eff94d420aa8117a7b6b860f4718a5df0d05b5350570b0e7c41f8c",
    "Payroll_Corrections_Mailbox.mbox": "a37d7f084048ccd37a7695e79dbb7370270636e76b387c12e7dcf2327a8dc298",
    "Payroll_Channel.json": "c1b149c979e2d158d2a62cf612f1d77268838494e9ae9698b067b09e6939e180",
    "Workflow_Context.json": "76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99",
}
WEIGHTS = {
    "R001": 4, "R002": 5, "R003": 5, "R004": 6, "R005": 8,
    "R006": 4, "R007": 4, "R008": 5, "R009": 4, "R010": 4,
    "R011": 4, "R012": 4, "R013": 4, "R014": 8, "R015": 6,
    "R016": 6, "R017": 7, "R018": 6, "R019": 6,
}
HURDLE_CAPS = {
    "R001": 0.0, "R002": 0.2, "R003": 0.2, "R004": 0.25, "R005": 0.3,
    "R006": 0.35, "R007": 0.35, "R008": 0.25, "R009": 0.35,
    "R010": 0.35, "R011": 0.35, "R012": 0.35, "R013": 0.35,
    "R014": 0.25, "R015": 0.25, "R016": 0.2, "R017": 0.35, "R018": 0.3,
}


@dataclass(frozen=True)
class SourceCase:
    source_row: int
    employee: str
    employee_email: str
    manager_email: str
    issue: str
    recorded_text: str
    notes: str
    source_status: str
    source_variance: str


class Checks:
    def __init__(self) -> None:
        self.values: dict[str, dict[str, Any]] = {
            criterion: {"passed": True, "details": []} for criterion in WEIGHTS
        }

    def fail(self, criterion: str, detail: str) -> None:
        self.values[criterion]["passed"] = False
        self.values[criterion]["details"].append(detail)

    def require(self, criterion: str, condition: bool, detail: str) -> None:
        if not condition:
            self.fail(criterion, detail)

    def passed(self, criterion: str) -> bool:
        return bool(self.values[criterion]["passed"])


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_money(value: Any) -> int:
    normalized = re.sub(r"[^0-9.-]", "", str(value or ""))
    if normalized in {"", "-", "."}:
        raise ValueError(f"无法解析金额：{value!r}")
    return int(float(normalized))


def decode_message(message) -> str:
    if message.is_multipart():
        parts: list[str] = []
        for part in message.walk():
            if part.get_content_type() == "text/plain":
                payload = part.get_payload(decode=True) or b""
                parts.append(payload.decode(part.get_content_charset() or "utf-8", errors="replace"))
        return "\n".join(parts).strip()
    payload = message.get_payload(decode=True) or b""
    return payload.decode(message.get_content_charset() or "utf-8", errors="replace").strip()


def expected_case_rows() -> list[tuple[Any, ...]]:
    return [
        (2, "Elise Fontaine", "elise.fontaine@company.example.com", "team-lead-a@company.example.com", "Overtime underpayment", "$360", "8 hrs OT at $45/hr", "Open", ""),
        (3, "George Nakamura", "george.nakamura@company.example.com", "team-lead-b@company.example.com", "Duplicate bonus", "$2,000", "Bonus entry", "Open", ""),
        (4, "Hannah Olsen", "hannah.olsen@company.example.com", "team-lead-a@company.example.com", "Missing shift differential", "$0", "Night shift premium $150 not applied - employee disputing, on hold per manager", "Open", "$150"),
        (5, "Ibrahim Sayed", "ibrahim.sayed@company.example.com", "team-lead-b@company.example.com", "Tax withholding error", "$3,200", "Should be $2,850 - state tax rate was wrong", "Open", "$350"),
        (6, "Jenny Park", "jenny.park@company.example.com", "team-lead-a@company.example.com", "PTO payout missing", "$0", "40 hrs PTO at $35/hr = $1,400 not paid", "Open", "$1,400"),
        (7, "Kevin Obi", "kevin.obi@company.example.com", "team-lead-b@company.example.com", "Commission shortfall", "$800", "Should be $1,050 per commission report", "Open", "$250"),
        (8, "Lina Marchetti", "lina.marchetti@company.example.com", "team-lead-a@company.example.com", "Health premium overcharge", "$450", "Should be $380 - wrong tier applied", "Open", "$70"),
        (9, "Miguel Santos", "miguel.santos@company.example.com", "team-lead-b@company.example.com", "Relocation stipend missing", "$0", "Approved relocation stipend of $3,500 not paid", "Open", "$3,500"),
    ]


def read_source_truth(checks: Checks) -> dict[str, Any]:
    for name, expected in INPUT_HASHES.items():
        path = INPUT_DIR / name
        checks.require("R001", path.is_file(), f"冻结输入缺失：{name}")
        if path.is_file():
            checks.require("R001", sha256(path) == expected, f"冻结输入哈希错误：{name}")

    workbook = load_workbook(INPUT_DIR / "Payroll_Discrepancy_Tracker.xlsx", read_only=True, data_only=True)
    checks.require(
        "R002", workbook.sheetnames == ["April Discrepancies", "Escalation Policy"],
        "输入工作簿双表结构错误",
    )
    raw_cases = []
    for row in workbook["April Discrepancies"].iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            raw_cases.append((int(row[0]),) + tuple(str(value or "") for value in row[1:9]))
    expected_cases = expected_case_rows()
    checks.require("R002", raw_cases == expected_cases, "八个冻结案例与上游输入不一致")
    cases = [SourceCase(*row) for row in expected_cases]

    raw_policies = []
    for row in workbook["Escalation Policy"].iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            raw_policies.append((int(row[0]), str(row[1]), str(row[2])))
    expected_rules = ["Variance <= $500", "Variance > $500", "Disputed/On Hold", "Direct Back-Pay"]
    checks.require("R003", [item[1] for item in raw_policies] == expected_rules, "四条冻结政策或顺序错误")
    checks.require(
        "R003", len(raw_policies) == 4 and all(
            token in raw_policies[3][2] for token in ("NOT HR Ops", "Only Payroll", "never execute")
        ), "直接补发权限边界错误",
    )

    box = mailbox.mbox(INPUT_DIR / "Payroll_Corrections_Mailbox.mbox", create=False)
    messages = list(box)
    box.close()
    checks.require("R004", len(messages) == 1, "冻结纠错邮箱必须恰好有1封邮件")
    message = messages[0]
    message_id = str(message.get("X-AutomationBench-Message-ID", ""))
    body = decode_message(message)
    checks.require("R004", message_id == "msg_payroll_correction_5104", "冻结纠错邮件 ID 错误")
    for token in (
        "Elise Fontaine", "12, not 8", "$45/hr", "$540", "$360",
        "George Nakamura", "$2,000", "already paid last period", "Remove the duplicate bonus entry",
    ):
        checks.require("R004", token in body, f"冻结纠错邮件缺少：{token}")

    context = json.loads((INPUT_DIR / "Workflow_Context.json").read_text(encoding="utf-8"))
    frozen_time = context.get("meta", {}).get("current_time")
    checks.require("R004", frozen_time == "2026-04-15T09:00:00Z", "冻结时间错误")
    channel = json.loads((INPUT_DIR / "Payroll_Channel.json").read_text(encoding="utf-8"))
    checks.require("R004", channel == {
        "channels": [{"id": "C_PAY_5104", "name": "payroll", "is_private": False}],
        "messages": [],
    }, "冻结 payroll 频道快照错误")

    correct_amounts = {
        "Elise Fontaine": 540, "George Nakamura": 0, "Hannah Olsen": 150,
        "Ibrahim Sayed": 2850, "Jenny Park": 1400, "Kevin Obi": 1050,
        "Lina Marchetti": 380, "Miguel Santos": 3500,
    }
    truth: dict[str, dict[str, Any]] = {}
    for case in cases:
        recorded = parse_money(case.recorded_text)
        correct = correct_amounts[case.employee]
        signed = correct - recorded
        absolute = abs(signed)
        on_hold = "disput" in case.notes.lower() or "on hold" in case.notes.lower()
        if on_hold:
            status, route, notification = "ON_HOLD", "NONE", "NO_DRAFT"
        elif absolute > 500:
            status, route, notification = "ESCALATED", "payroll-director@company.example.com", "DRAFT_NOT_SENT"
        else:
            status, route, notification = "RESOLVED", case.employee_email, "DRAFT_NOT_SENT"
        truth[case.employee] = {
            "source_row": case.source_row, "employee_email": case.employee_email,
            "manager_email": case.manager_email, "issue": case.issue,
            "recorded_text": case.recorded_text, "notes": case.notes,
            "recorded": recorded, "correct": correct, "signed": signed, "absolute": absolute,
            "on_hold": "YES" if on_hold else "NO", "status": status,
            "route": route, "notification": notification,
        }
    checks.require("R004", sum(value["absolute"] for value in truth.values()) == 7900, "来源总差异推导错误")
    checks.require(
        "R004", sum(value["absolute"] for value in truth.values() if value["status"] != "ON_HOLD") == 7750,
        "来源可行动差异推导错误",
    )
    return {
        "cases": cases, "policies": raw_policies, "email_body": body,
        "email_id": message_id, "frozen_time": frozen_time, "channel": channel,
        "truth": truth,
    }


def is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def workbook_has_unsafe_features(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            if any(name.endswith("vbaProject.bin") for name in names):
                return True
            for name in names:
                if name.endswith(".rels"):
                    text = archive.read(name).decode("utf-8", errors="ignore")
                    if 'TargetMode="External"' in text:
                        return True
    except Exception:
        return True
    return False


def case_criterion(employee: str) -> str:
    return {
        "Elise Fontaine": "R006", "George Nakamura": "R007", "Hannah Olsen": "R008",
        "Ibrahim Sayed": "R009", "Jenny Park": "R010", "Kevin Obi": "R011",
        "Lina Marchetti": "R012", "Miguel Santos": "R013",
    }[employee]


def check_workbook(path: Path, source: dict[str, Any], checks: Checks) -> dict[str, Any]:
    try:
        formula_book = load_workbook(path, data_only=False)
        value_book = load_workbook(path, data_only=True)
    except Exception as exc:
        for criterion in (
            "R002", "R003", "R004", "R005", "R006", "R007", "R008", "R009",
            "R010", "R011", "R012", "R013", "R018", "R019",
        ):
            checks.fail(criterion, f"Excel 无法解析：{exc}")
        return {}

    expected_sheets = [
        "Source_Cases", "Source_Policy", "Evidence_Register",
        "Case_Decisions", "Routing_Queue", "Control_Summary",
    ]
    checks.require("R019", formula_book.sheetnames == expected_sheets, "Excel 必须恰好包含6张任务专属工作表")
    if any(name not in formula_book.sheetnames for name in expected_sheets):
        for criterion in (
            "R002", "R003", "R004", "R005", "R006", "R007", "R008", "R009",
            "R010", "R011", "R012", "R013", "R018",
        ):
            checks.fail(criterion, "Excel 缺少核心工作表")
        return {}

    source_case_rows = []
    for row in value_book["Source_Cases"].iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            source_case_rows.append((int(row[0]),) + tuple(str(value or "") for value in row[1:9]))
    checks.require("R002", source_case_rows == expected_case_rows(), "Source_Cases 与冻结输入不一致")

    source_policy_rows = []
    for row in value_book["Source_Policy"].iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            source_policy_rows.append((int(row[0]), str(row[1]), str(row[2])))
    checks.require("R003", source_policy_rows == source["policies"], "Source_Policy 与冻结输入不一致")

    evidence_values = value_book["Evidence_Register"]
    evidence_text = "\n".join(str(cell.value or "") for row in evidence_values.iter_rows() for cell in row)
    for token in (
        source["email_body"], "msg_payroll_correction_5104", "2026-04-15T09:00:00Z",
        "C_PAY_5104", "Payroll_Channel.json:channels[0]",
    ):
        checks.require("R004", token in evidence_text, f"Evidence_Register 缺少完整证据：{str(token)[:80]}")
    for row_index, case in enumerate(source["cases"], start=2):
        truth = source["truth"][case.employee]
        checks.require("R004", evidence_values.cell(row_index, 1).value == case.source_row, f"{case.employee} 的证据源行错误")
        checks.require("R004", evidence_values.cell(row_index, 2).value == case.employee, f"{case.employee} 的证据人员错误")
        checks.require("R004", evidence_values.cell(row_index, 5).value == truth["correct"], f"{case.employee} 的证据金额错误")
        trace = str(evidence_values.cell(row_index, 8).value or "")
        checks.require("R019", "Payroll_Discrepancy_Tracker.xlsx" in trace, f"{case.employee} 缺少输入追踪")
        if case.employee in {"Elise Fontaine", "George Nakamura"}:
            checks.require("R004", "msg_payroll_correction_5104" in trace, f"{case.employee} 缺少纠错邮件追踪")
    checks.require(
        "R004", [evidence_values.cell(10, column).value for column in (1, 2, 3, 4, 6, 8)] == [
            "MAIL", "Payroll manager correction", "FULL_EMAIL", "msg_payroll_correction_5104",
            source["email_body"], "Payroll_Corrections_Mailbox.mbox:msg_payroll_correction_5104",
        ], "Evidence_Register 的完整邮件记录错误",
    )
    checks.require(
        "R004", [evidence_values.cell(11, column).value for column in (1, 2, 3, 4, 6, 7, 8)] == [
            "TIME", "Workflow context", "FROZEN_TIME", "meta.current_time",
            "2026-04-15T09:00:00Z", "2026-04-15T09:00:00Z", "Workflow_Context.json:meta.current_time",
        ], "Evidence_Register 的冻结时间记录错误",
    )
    checks.require(
        "R004", [evidence_values.cell(12, column).value for column in (1, 2, 3, 4, 6, 8)] == [
            "SLACK", "Payroll channel", "CHANNEL_SNAPSHOT", "C_PAY_5104",
            '{"id": "C_PAY_5104", "is_private": false, "name": "payroll"}',
            "Payroll_Channel.json:channels[0]",
        ], "Evidence_Register 的 Slack 频道记录错误",
    )

    formula_sheet = formula_book["Case_Decisions"]
    value_sheet = value_book["Case_Decisions"]
    output_truth: dict[str, dict[str, Any]] = {}
    for row_index in range(2, 10):
        employee = str(value_sheet.cell(row_index, 2).value or "")
        if employee not in source["truth"]:
            checks.fail("R002", f"Case_Decisions 出现未知员工：{employee}")
            continue
        truth = source["truth"][employee]
        criterion = case_criterion(employee)
        actual = {
            "source_row": value_sheet.cell(row_index, 1).value,
            "employee_email": str(value_sheet.cell(row_index, 3).value or ""),
            "manager_email": str(value_sheet.cell(row_index, 4).value or ""),
            "issue": str(value_sheet.cell(row_index, 5).value or ""),
            "recorded_text": str(value_sheet.cell(row_index, 6).value or ""),
            "notes": str(value_sheet.cell(row_index, 7).value or ""),
            "recorded": value_sheet.cell(row_index, 8).value,
            "correct": value_sheet.cell(row_index, 9).value,
            "signed": value_sheet.cell(row_index, 10).value,
            "absolute": value_sheet.cell(row_index, 11).value,
            "on_hold": str(value_sheet.cell(row_index, 12).value or ""),
            "status": str(value_sheet.cell(row_index, 13).value or ""),
            "route": str(value_sheet.cell(row_index, 14).value or ""),
            "notification": str(value_sheet.cell(row_index, 15).value or ""),
            "trace": str(value_sheet.cell(row_index, 16).value or ""),
        }
        output_truth[employee] = actual
        for field in (
            "source_row", "employee_email", "manager_email", "issue", "recorded_text", "notes",
            "recorded", "correct", "signed", "absolute", "on_hold", "status", "route", "notification",
        ):
            checks.require(criterion, actual[field] == truth[field], f"{employee} 的 {field} 错误")
        for column in range(1, 17):
            checks.require("R005", is_formula(formula_sheet.cell(row_index, column).value), f"Case_Decisions 第{row_index}行第{column}列必须是公式")
        checks.require("R019", "Payroll_Discrepancy_Tracker.xlsx" in actual["trace"], f"{employee} 决策缺少来源追踪")
    checks.require("R002", set(output_truth) == set(source["truth"]), "Case_Decisions 必须恰好覆盖8名员工")

    routing_formula = formula_book["Routing_Queue"]
    routing_values = value_book["Routing_Queue"]
    for row_index, case in enumerate(source["cases"], start=2):
        truth = source["truth"][case.employee]
        actual = [routing_values.cell(row_index, column).value for column in range(1, 10)]
        expected = [
            case.source_row, case.employee, truth["status"], truth["absolute"], truth["route"],
            truth["notification"],
            "DIRECTOR_CONSOLIDATED" if truth["status"] == "ESCALATED" else (
                "EMPLOYEE_INDIVIDUAL" if truth["status"] == "RESOLVED" else "NONE"
            ),
            "NO_ACTION" if truth["status"] == "ON_HOLD" else "ROUTE_ONLY_NO_PAYMENT",
        ]
        checks.require("R018", actual[:8] == expected, f"Routing_Queue 中 {case.employee} 的路由链错误")
        for column in range(1, 10):
            checks.require("R005", is_formula(routing_formula.cell(row_index, column).value), f"Routing_Queue 第{row_index}行第{column}列必须是公式")

    summary_formula = formula_book["Control_Summary"]
    summary_values = value_book["Control_Summary"]
    expected_summary = [8, 4, 3, 1, 7900, 7750, 150, 5, 0, 0]
    actual_summary = [summary_values.cell(row, 2).value for row in range(2, 12)]
    checks.require("R005", actual_summary == expected_summary, "Control_Summary 的计数或金额错误")
    checks.require("R005", all(is_formula(summary_formula.cell(row, 2).value) for row in range(2, 12)), "摘要值必须全部为公式")

    for sheet_name in expected_sheets:
        checks.require("R019", len(formula_book[sheet_name].tables) >= 1, f"{sheet_name} 缺少原生表格")
    rules = sum(len(value) for value in formula_book["Case_Decisions"].conditional_formatting._cf_rules.values())
    checks.require("R019", rules >= 3, "Case_Decisions 缺少三类状态条件格式")
    checks.require("R019", len(formula_book["Control_Summary"]._charts) == 1, "Control_Summary 必须恰好包含1个可编辑图表")
    checks.require("R019", not workbook_has_unsafe_features(path), "Excel 含宏或危险外链")
    return {"decisions": output_truth, "summary": actual_summary}


def parse_output_mbox(path: Path, source: dict[str, Any], checks: Checks) -> dict[str, dict[str, str]]:
    try:
        box = mailbox.mbox(path, create=False)
        messages = []
        for message in box:
            messages.append({
                "to": str(message.get("To", "")),
                "status": str(message.get("X-EvoCowork-Status", "")),
                "source": str(message.get("X-EvoCowork-Source", "")),
                "subject": str(message.get("Subject", "")),
                "body": decode_message(message),
            })
        box.close()
    except Exception as exc:
        checks.fail("R014", f"MBOX 无法解析：{exc}")
        return {}
    checks.require("R014", len(messages) == 5, "MBOX 必须恰好含5封草稿")
    mapped = {item["to"]: item for item in messages}
    expected_recipients = {
        "elise.fontaine@company.example.com",
        "ibrahim.sayed@company.example.com",
        "kevin.obi@company.example.com",
        "lina.marchetti@company.example.com",
        "payroll-director@company.example.com",
    }
    checks.require("R014", set(mapped) == expected_recipients, "工资路由草稿收件人集合错误")
    checks.require("R014", all(item["status"] == "DRAFT_NOT_SENT" for item in messages), "MBOX 全部邮件必须是 DRAFT_NOT_SENT")
    checks.require("R014", all(item["source"] == "AutomationBench-5104" for item in messages), "MBOX 缺少任务来源头")

    all_names = set(source["truth"])
    for name in ("Elise Fontaine", "Ibrahim Sayed", "Kevin Obi", "Lina Marchetti"):
        truth = source["truth"][name]
        recipient = truth["employee_email"]
        if recipient not in mapped:
            continue
        body = mapped[recipient]["body"]
        for token in (
            name, truth["issue"], f"${truth['recorded']:,}", f"${truth['correct']:,}",
            f"{truth['signed']:+,}", f"${truth['absolute']:,}", "RESOLVED",
            "ROUTE_ONLY_NO_PAYMENT", "No payment", "has been executed",
        ):
            checks.require("R014", token in body, f"{name} 草稿缺少：{token}")
        for other in all_names - {name}:
            checks.require("R014", other not in body, f"{name} 的个人草稿错含：{other}")

    director = mapped.get("payroll-director@company.example.com", {})
    director_body = director.get("body", "")
    for name in ("George Nakamura", "Jenny Park", "Miguel Santos"):
        truth = source["truth"][name]
        for token in (name, truth["issue"], f"${truth['correct']:,}", f"{truth['signed']:+,}", f"${truth['absolute']:,}"):
            checks.require("R014", token in director_body, f"工资总监草稿缺少 {name} 的：{token}")
    for name in ("Hannah Olsen", "Elise Fontaine", "Ibrahim Sayed", "Kevin Obi", "Lina Marchetti"):
        checks.require("R014", name not in director_body, f"工资总监草稿错含：{name}")
    checks.require(
        "R014", "One disputed $150 case remains ON_HOLD and is excluded" in director_body,
        "工资总监草稿必须解释暂停案例未进入升级",
    )
    manager_addresses = {case.manager_email for case in source["cases"]}
    checks.require("R014", not (set(mapped) & manager_addresses), "MBOX 不得向经理地址发草稿")
    serialized = "\n".join(item["body"] for item in messages)
    for unsafe in ("back-pay issued", "payment completed", "transaction completed", "Slack posted"):
        checks.require("R016", unsafe.lower() not in serialized.lower(), f"MBOX 含虚假执行声明：{unsafe}")
    checks.require("R014", "Hannah Olsen" not in "\n".join(item["body"] for item in messages[:-1]), "Hannah 不得收到或进入员工草稿")
    return mapped


def check_slack(path: Path, source: dict[str, Any], checks: Checks) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        checks.fail("R015", f"Slack JSON 无法解析：{exc}")
        return {}
    checks.require("R015", payload.get("source_channel") == {
        "id": "C_PAY_5104", "name": "payroll", "is_private": False,
    }, "Slack 频道错误")
    checks.require("R015", payload.get("status") == "DRAFT_NOT_POSTED" and payload.get("posted") is False, "Slack 必须是未发布草稿")
    checks.require("R015", payload.get("summary") == {
        "total_cases": 8, "resolved": 4, "escalated": 3, "on_hold": 1,
        "actionable_absolute_variance_usd": 7750,
    }, "Slack 聚合计数或金额错误")
    privacy = payload.get("privacy_control", {})
    checks.require("R015", privacy == {
        "aggregate_only": True, "employee_names_included": False,
        "employee_emails_included": False, "case_level_amounts_included": False,
    }, "Slack 隐私控制字段错误")
    text = str(payload.get("draft_message", {}).get("text", ""))
    for token in ("8 cases", "4 RESOLVED", "3 ESCALATED", "1 ON_HOLD", "$7,750", "no payroll execution"):
        checks.require("R015", token in text, f"Slack 草稿缺少：{token}")
    serialized = json.dumps(payload, ensure_ascii=False)
    for case in source["cases"]:
        for token in (case.employee, case.employee_email, case.issue):
            checks.require("R015", token not in serialized, f"Slack 泄露逐案信息：{token}")
    for amount in (180, 2000, 150, 350, 1400, 250, 70, 3500):
        checks.require("R015", f"${amount:,}" not in serialized, f"Slack 泄露逐人金额：${amount:,}")
    traces = payload.get("source_trace", [])
    checks.require("R019", "Payroll_Discrepancy_Audit.xlsx:Control_Summary" in traces, "Slack 缺少审计簿追踪")
    checks.require("R019", "Payroll_Channel.json:channels[0]" in traces, "Slack 缺少频道追踪")
    return payload


def check_execution_control(path: Path, checks: Checks) -> dict[str, Any]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        checks.fail("R016", f"补发执行控制 JSON 无法解析：{exc}")
        return {}
    checks.require(
        "R016", payload.get("decision") == "BLOCKED_ROUTE_ONLY" and payload.get("status") == "BLOCKED_ROUTE_ONLY",
        "补发执行阻断状态错误",
    )
    checks.require("R016", payload.get("executed") is False, "补发执行必须为 false")
    checks.require("R016", payload.get("execution_record_count") == 0, "补发执行记录数必须为0")
    checks.require("R016", payload.get("transaction_ids") == [], "不得虚构交易 ID")
    policy = payload.get("source_policy", {})
    checks.require("R016", policy == {
        "file": "Payroll_Discrepancy_Tracker.xlsx", "sheet": "Escalation Policy",
        "source_row": 5, "rule": "Direct Back-Pay",
    }, "补发阻断必须精确引用政策源行5")
    reason = str(payload.get("control_reason", ""))
    checks.require("R016", all(token in reason for token in ("not authorized", "Only Payroll", "approval chain")), "补发阻断原因不完整")
    traces = payload.get("source_trace", [])
    checks.require("R019", "Payroll_Discrepancy_Tracker.xlsx:Escalation Policy row 5" in traces, "补发控制缺少政策追踪")
    return payload


def document_text(document: Document) -> str:
    chunks = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            chunks.extend(cell.text for cell in row.cells)
    return "\n".join(chunks)


def check_document(path: Path, source: dict[str, Any], checks: Checks) -> dict[str, Any]:
    try:
        document = Document(path)
    except Exception as exc:
        checks.fail("R017", f"Word 无法解析：{exc}")
        return {}
    text = document_text(document)
    required = [
        "Payroll Discrepancy Control Brief", "April 2026", "2026-04-15", "4 RESOLVED",
        "3 ESCALATED", "1 ON_HOLD", "$7,900", "$7,750", "$150",
        "msg_payroll_correction_5104", "C_PAY_5104", "DRAFT_NOT_SENT",
        "DRAFT_NOT_POSTED", "BLOCKED_ROUTE_ONLY", "executed=false", "execution_record_count=0",
        "Nothing was sent or posted", "payroll-director@company.example.com",
    ]
    for token in required:
        checks.require("R017", token in text, f"Word 缺少：{token}")
    for case in source["cases"]:
        checks.require("R017", case.employee in text, f"Word 缺少案例：{case.employee}")
    for _, rule, action in source["policies"]:
        checks.require("R017", rule in text and action in text, f"Word 缺少完整政策：{rule}")
    checks.require("R017", len(document.tables) == 5, "Word 必须恰好包含5张任务专属表")

    tables = [
        [[cell.text.strip() for cell in row.cells] for row in table.rows]
        for table in document.tables
    ]
    if len(tables) == 5:
        checks.require("R017", tables[0][1] == ["8", "4", "3", "1", "$7,750", "$150"], "Word 指标表错误")
        case_rows = {row[0]: row for row in tables[1][1:] if row}
        checks.require("R017", set(case_rows) == set(source["truth"]), "Word 案例表必须恰好覆盖8人")
        for name, truth in source["truth"].items():
            if name not in case_rows:
                continue
            row_text = "\n".join(case_rows[name])
            expected_route = "Employee draft" if truth["status"] == "RESOLVED" else (
                "Payroll director draft" if truth["status"] == "ESCALATED" else "No email"
            )
            for token in (
                truth["issue"], f"${truth['recorded']:,}", f"${truth['correct']:,}",
                f"{truth['signed']:+,}", f"${truth['absolute']:,}", truth["status"], expected_route,
            ):
                checks.require("R017", token in row_text, f"Word 案例表 {name} 缺少：{token}")
        routing_rows = tables[3][1:]
        routing_text = "\n".join(" | ".join(row) for row in routing_rows)
        for recipient in (
            "elise.fontaine@company.example.com", "ibrahim.sayed@company.example.com",
            "kevin.obi@company.example.com", "lina.marchetti@company.example.com",
            "payroll-director@company.example.com",
        ):
            checks.require("R017", recipient in routing_text, f"Word 路由表缺少：{recipient}")
        checks.require("R017", "hannah.olsen@company.example.com" not in routing_text, "Word 路由表不得含 Hannah 收件人")
        evidence_rows = {row[0]: row for row in tables[2][1:] if row}
        expected_evidence = {
            "Elise Fontaine": ["Elise Fontaine", "msg_payroll_correction_5104", "12 × $45 = $540; +$180; employee draft."],
            "George Nakamura": ["George Nakamura", "msg_payroll_correction_5104", "$2,000 already paid; current correct $0; director escalation."],
            "Hannah Olsen": ["Hannah Olsen", "April Discrepancies row 4", "$150 disputed / on hold; no email and no processing."],
            "Ibrahim Sayed": ["Ibrahim Sayed", "April Discrepancies row 5", "Correct withholding $2,850; -$350; employee draft."],
            "Jenny Park": ["Jenny Park", "April Discrepancies row 6", "40 × $35 = $1,400; director escalation."],
            "Kevin Obi": ["Kevin Obi", "April Discrepancies row 7", "Correct commission $1,050; +$250; employee draft."],
            "Lina Marchetti": ["Lina Marchetti", "April Discrepancies row 8", "Correct premium $380; -$70; employee draft."],
            "Miguel Santos": ["Miguel Santos", "April Discrepancies row 9", "Approved stipend $3,500; director escalation."],
        }
        checks.require("R017", evidence_rows == expected_evidence, "Word 逐案证据表与冻结真值不一致")

    lower = text.lower()
    unsafe_claims = ["payment was issued", "back-pay issued", "slack was posted", "email was sent", "transaction completed"]
    checks.require("R016", not any(token in lower for token in unsafe_claims), "Word 含虚假发送或付款声明")
    return {"text": text, "tables": tables}


def render_office(path: Path) -> tuple[int, str]:
    with tempfile.TemporaryDirectory(prefix="payroll-control-render-") as temporary:
        root = Path(temporary)
        output = root / "output"
        profile = root / "profile"
        output.mkdir()
        process = subprocess.run(
            [
                "libreoffice", "--headless", f"-env:UserInstallation=file://{profile}",
                "--convert-to", "pdf", "--outdir", str(output), str(path),
            ],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=300,
        )
        pdf = output / f"{path.stem}.pdf"
        if process.returncode != 0 or not pdf.is_file():
            raise RuntimeError(process.stdout[-1000:])
        info = subprocess.run(
            ["pdfinfo", str(pdf)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=60,
        )
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, flags=re.MULTILINE)
        if not match:
            raise RuntimeError("pdfinfo 未返回页数")
        rendered = subprocess.run(
            ["pdftotext", str(pdf), "-"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=60,
        )
        return int(match.group(1)), rendered.stdout


def evaluate(artifact_dir: Path) -> dict[str, Any]:
    checks = Checks()
    actual_files = {item.name for item in artifact_dir.iterdir() if item.is_file()} if artifact_dir.is_dir() else set()
    nested = [item.name for item in artifact_dir.iterdir() if item.is_dir()] if artifact_dir.is_dir() else []
    checks.require(
        "R001", actual_files == EXPECTED_FILES,
        f"交付文件必须恰好为 {sorted(EXPECTED_FILES)}，实际 {sorted(actual_files)}",
    )
    checks.require("R001", not nested, f"输出不得包含子目录：{nested}")
    source = read_source_truth(checks)

    workbook_result: dict[str, Any] = {}
    mbox_result: dict[str, Any] = {}
    slack_result: dict[str, Any] = {}
    control_result: dict[str, Any] = {}
    document_result: dict[str, Any] = {}
    if XLSX_NAME in actual_files:
        workbook_result = check_workbook(artifact_dir / XLSX_NAME, source, checks)
    else:
        for criterion in (
            "R002", "R003", "R004", "R005", "R006", "R007", "R008", "R009",
            "R010", "R011", "R012", "R013", "R018", "R019",
        ):
            checks.fail(criterion, "Excel 交付不可评")
    if MBOX_NAME in actual_files:
        mbox_result = parse_output_mbox(artifact_dir / MBOX_NAME, source, checks)
    else:
        checks.fail("R014", "MBOX 交付不可评")
    if SLACK_NAME in actual_files:
        slack_result = check_slack(artifact_dir / SLACK_NAME, source, checks)
    else:
        checks.fail("R015", "Slack JSON 交付不可评")
    if CONTROL_NAME in actual_files:
        control_result = check_execution_control(artifact_dir / CONTROL_NAME, checks)
    else:
        checks.fail("R016", "补发执行控制 JSON 交付不可评")
    if DOCX_NAME in actual_files:
        document_result = check_document(artifact_dir / DOCX_NAME, source, checks)
    else:
        checks.fail("R017", "Word 交付不可评")

    decisions = workbook_result.get("decisions", {})
    if decisions:
        checks.require(
            "R018", {name: value.get("status") for name, value in decisions.items()}
            == {name: value["status"] for name, value in source["truth"].items()},
            "Excel 的案例状态与来源真值冲突",
        )
    if mbox_result:
        checks.require("R018", len(mbox_result) == 5, "Excel 路由与 MBOX 草稿数冲突")
    if slack_result:
        checks.require("R018", slack_result.get("summary", {}).get("actionable_absolute_variance_usd") == 7750, "Excel 与 Slack 可行动金额冲突")
        checks.require("R018", slack_result.get("summary", {}).get("resolved") == 4, "Excel 与 Slack RESOLVED 计数冲突")
        checks.require("R018", slack_result.get("summary", {}).get("escalated") == 3, "Excel 与 Slack ESCALATED 计数冲突")
        checks.require("R018", slack_result.get("summary", {}).get("on_hold") == 1, "Excel 与 Slack ON_HOLD 计数冲突")
    if control_result:
        checks.require("R018", control_result.get("executed") is False and control_result.get("execution_record_count") == 0, "执行控制与 Excel 零执行边界冲突")
    if document_result:
        text = document_result.get("text", "")
        for token in ("4 RESOLVED", "3 ESCALATED", "1 ON_HOLD", "$7,900", "$7,750", "$150"):
            checks.require("R018", token in text, f"Word 与其他交付冲突或缺少：{token}")

    render_pages: dict[str, int] = {}
    core_ready = all(checks.passed(f"R{index:03d}") for index in range(1, 19))
    if core_ready:
        for name, expected_pages in ((XLSX_NAME, 6), (DOCX_NAME, 3)):
            try:
                pages, rendered_text = render_office(artifact_dir / name)
                render_pages[name] = pages
                checks.require("R019", pages == expected_pages, f"{name} 应渲染 {expected_pages} 页，实际 {pages}")
                checks.require(
                    "R019", not re.search(r"#(?:VALUE|N/A|REF|NAME|DIV/0)", rendered_text, flags=re.IGNORECASE),
                    f"{name} 渲染出现公式错误",
                )
            except Exception as exc:
                checks.fail("R019", f"{name} LibreOffice 渲染失败：{exc}")
    else:
        checks.fail("R019", "核心交付未通过，不授予视觉就绪分")

    hurdle_failures = [criterion for criterion in HURDLE_CAPS if not checks.passed(criterion)]
    base_score = sum(weight for criterion, weight in WEIGHTS.items() if checks.passed(criterion)) / 100.0
    cap = min([1.0] + [HURDLE_CAPS[criterion] for criterion in hurdle_failures])
    penalties: list[dict[str, Any]] = []
    penalty_total = 0.0
    if any(not checks.passed(criterion) for criterion in (
        "R002", "R003", "R004", "R005", "R006", "R007", "R008", "R009",
        "R010", "R011", "R012", "R013",
    )):
        penalties.append({"id": "R020", "amount": 0.30, "reason": "来源、政策、证据、公式、金额或案例状态错误"})
        penalty_total += 0.30
    if any(not checks.passed(criterion) for criterion in (
        "R006", "R007", "R008", "R009", "R010", "R011", "R012", "R013", "R014", "R015",
    )):
        penalties.append({"id": "R021", "amount": 0.30, "reason": "案例路由、通知选择、暂停边界或 Slack 隐私错误"})
        penalty_total += 0.30
    if any(not checks.passed(criterion) for criterion in ("R014", "R015", "R016", "R017")):
        penalties.append({"id": "R022", "amount": 0.30, "reason": "虚假发送、公开发布、补发执行或交易声明"})
        penalty_total += 0.30
    if any(not checks.passed(criterion) for criterion in ("R001", "R005", "R018", "R019")):
        penalties.append({"id": "R023", "amount": 0.25, "reason": "文件、公式、原生结构、追踪、跨文件或可渲染性错误"})
        penalty_total += 0.25
    reward = round(max(0.0, min(base_score, cap) - penalty_total), 6)
    return {
        "task_id": TASK_ID,
        "reward": reward,
        "passed": reward >= 0.8 and not hurdle_failures,
        "base_score": round(base_score, 6),
        "hurdle_cap": cap,
        "hurdle_failures": hurdle_failures,
        "penalties": penalties,
        "criteria": checks.values,
        "source_truth": {
            "input_sha256": INPUT_HASHES,
            "case_count": len(source["cases"]),
            "policy_count": len(source["policies"]),
            "state_counts": {"RESOLVED": 4, "ESCALATED": 3, "ON_HOLD": 1},
            "total_absolute_variance": 7900,
            "actionable_absolute_variance": 7750,
        },
        "render_pages": render_pages,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = evaluate(args.artifact_dir)
    except Exception as exc:
        result = {
            "task_id": TASK_ID, "reward": 0.0, "passed": False,
            "fatal_error": f"{type(exc).__name__}: {exc}", "hurdle_failures": ["R001"],
        }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{float(result['reward']):.6f}\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
