#!/usr/bin/env python3
"""AutomationBench 5110 组织改线控制题的逐题确定性 Judge。"""

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
from email.utils import getaddresses
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


TASK_ID = "cowork-evo/automationbench-5110-org-restructure-control-package"
XLSX_NAME = "Org_Restructure_Audit.xlsx"
DOCX_NAME = "Org_Restructure_Control_Brief.docx"
PPTX_NAME = "Org_Restructure_Chart.pptx"
MBOX_NAME = "Org_Change_Notification_Drafts.mbox"
UPDATES_NAME = "Directory_Update_Drafts.json"
SLACK_NAME = "HR_Ops_Reorg_Draft.json"
SCOPE_NAME = "Scope_Control.json"
REQUIRED_FILES = {
    XLSX_NAME, DOCX_NAME, PPTX_NAME, MBOX_NAME,
    UPDATES_NAME, SLACK_NAME, SCOPE_NAME,
}
INPUT_HASHES = {
    "Employee_Directory.xlsx": "42c477c7905f4dfd072be5c6cb7a20bc127cc275c5aeff37553069cb07b80912",
    "Approved_Org_Changes.mbox": "f463fc846844feea9354791ee4c9907dc668fe5d9ba0abb6707ef2b0d9ff4e12",
    "HR_Ops_Channel.json": "6c39fc7813952d47266716507949c4b7a9adb3c9cf3ee398722595adc36371cd",
    "Workflow_Context.json": "76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99",
}
WEIGHTS = {
    "R001": 4, "R002": 6, "R003": 7, "R004": 7,
    "R005": 4, "R006": 4, "R007": 4, "R008": 4,
    "R009": 5, "R010": 5, "R011": 7, "R012": 8,
    "R013": 6, "R014": 5, "R015": 8, "R016": 6,
    "R017": 5, "R018": 5,
}
HURDLE_CAPS = {
    "R001": 0.0, "R002": 0.2, "R003": 0.2, "R004": 0.3,
    "R005": 0.3, "R006": 0.3, "R007": 0.3, "R008": 0.2,
    "R009": 0.2, "R010": 0.25, "R011": 0.2, "R012": 0.2,
    "R013": 0.25, "R014": 0.25, "R015": 0.3, "R016": 0.35,
}

SOURCE_ROWS = [
    (2, "Aisha Jordan", "aisha.jordan@company.example.com", "Engineering", "Lisa Wang", "lisa.wang@company.example.com"),
    (3, "Bo Lindqvist", "bo.lindqvist@company.example.com", "Engineering", "Marcus Chen", "marcus.chen@company.example.com"),
    (4, "Chiara Russo", "chiara.russo@company.example.com", "Engineering", "Derek Huang", "derek.huang@company.example.com"),
    (5, "Dante Obi", "dante.obi@company.example.com", "Product", "Lisa Wang", "lisa.wang@company.example.com"),
    (6, "Elena Vasquez", "elena.vasquez@company.example.com", "Engineering", "Derek Huang", "derek.huang@company.example.com"),
    (7, "Felix Braun", "felix.braun@company.example.com", "Engineering", "Marcus Chen", "marcus.chen@company.example.com"),
    (8, "Grace Kim", "grace.kim@company.example.com", "Product", "Marcus Chen-Williams", "marcus.chen-williams@company.example.com"),
    (9, "Hugo Fernandez", "hugo.fernandez@company.example.com", "Engineering", "Lisa Wang", "lisa.wang@company.example.com"),
]
CHANGES = [
    (1, 2, "Aisha Jordan", "aisha.jordan@company.example.com", "Engineering", "Lisa Wang", "lisa.wang@company.example.com", "Marcus Chen", "marcus.chen@company.example.com", "R005"),
    (2, 3, "Bo Lindqvist", "bo.lindqvist@company.example.com", "Engineering", "Marcus Chen", "marcus.chen@company.example.com", "Derek Huang", "derek.huang@company.example.com", "R006"),
    (3, 4, "Chiara Russo", "chiara.russo@company.example.com", "Engineering", "Derek Huang", "derek.huang@company.example.com", "Lisa Wang", "lisa.wang@company.example.com", "R007"),
    (4, 5, "Dante Obi", "dante.obi@company.example.com", "Product", "Lisa Wang", "lisa.wang@company.example.com", "Marcus Chen-Williams", "marcus.chen-williams@company.example.com", "R008"),
]
UNAFFECTED = [
    (6, "Elena Vasquez", "elena.vasquez@company.example.com", "Engineering", "Derek Huang", "derek.huang@company.example.com"),
    (7, "Felix Braun", "felix.braun@company.example.com", "Engineering", "Marcus Chen", "marcus.chen@company.example.com"),
    (8, "Grace Kim", "grace.kim@company.example.com", "Product", "Marcus Chen-Williams", "marcus.chen-williams@company.example.com"),
    (9, "Hugo Fernandez", "hugo.fernandez@company.example.com", "Engineering", "Lisa Wang", "lisa.wang@company.example.com"),
]
MANAGERS = [
    ("Lisa Wang", "lisa.wang@company.example.com", "2, 5, 9", "Engineering + Product", "EXACT_NAME_AND_EMAIL", "Employee_Directory.xlsx:Employees rows 2,5,9"),
    ("Marcus Chen", "marcus.chen@company.example.com", "3, 7", "Engineering", "DISTINCT_FROM_MARCUS_CHEN_WILLIAMS", "Employee_Directory.xlsx:Employees rows 3,7"),
    ("Derek Huang", "derek.huang@company.example.com", "4, 6", "Engineering", "EXACT_NAME_AND_EMAIL", "Employee_Directory.xlsx:Employees rows 4,6"),
    ("Marcus Chen-Williams", "marcus.chen-williams@company.example.com", "8", "Product", "DISTINCT_FROM_MARCUS_CHEN", "Employee_Directory.xlsx:Employees row 8"),
]
GROUPS = {
    "Lisa Wang": {"Chiara Russo", "Hugo Fernandez"},
    "Marcus Chen": {"Aisha Jordan", "Felix Braun"},
    "Derek Huang": {"Bo Lindqvist", "Elena Vasquez"},
    "Marcus Chen-Williams": {"Dante Obi", "Grace Kim"},
}


@dataclass
class CheckLedger:
    values: dict[str, dict[str, Any]]

    def __init__(self) -> None:
        self.values = {key: {"passed": True, "details": []} for key in WEIGHTS}

    def require(self, criterion: str, condition: bool, detail: str) -> None:
        if not condition:
            self.values[criterion]["passed"] = False
            self.values[criterion]["details"].append(detail)

    def fail(self, criterion: str, detail: str) -> None:
        self.require(criterion, False, detail)

    def passed(self, criterion: str) -> bool:
        return bool(self.values[criterion]["passed"])


def task_root() -> Path:
    return Path(__file__).resolve().parents[1]


def input_root() -> Path:
    container = Path("/app/input_files")
    return container if container.is_dir() else task_root() / "data" / "input_files"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def normalize(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def value_rows(sheet, start: int, end: int, columns: int) -> list[tuple[Any, ...]]:
    return [tuple(normalize(sheet.cell(row, column).value) for column in range(1, columns + 1)) for row in range(start, end + 1)]


def formula_depends_on(value: Any, *tokens: str) -> bool:
    formula = str(value or "")
    return formula.startswith("=") and all(token.casefold() in formula.casefold() for token in tokens)


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


def addresses(message, header: str) -> set[str]:
    return {address.casefold() for _, address in getaddresses(message.get_all(header, [])) if address}


def approved_body() -> str:
    box = mailbox.mbox(input_root() / "Approved_Org_Changes.mbox", create=False)
    messages = list(box)
    box.close()
    if len(messages) != 1 or messages[0].get("X-AutomationBench-Message-ID") != "msg_reorg_5110":
        raise ValueError("批准邮件输入不唯一")
    return decode_message(messages[0])


def render_office(path: Path) -> tuple[int, str]:
    with tempfile.TemporaryDirectory(prefix="judge-5110-render-") as temporary:
        root = Path(temporary)
        output, profile = root / "output", root / "profile"
        output.mkdir()
        process = subprocess.run(
            ["libreoffice", "--headless", f"-env:UserInstallation=file://{profile}", "--convert-to", "pdf", "--outdir", str(output), str(path)],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=300,
        )
        pdf = output / f"{path.stem}.pdf"
        if process.returncode != 0 or not pdf.is_file():
            raise RuntimeError(process.stdout[-1000:])
        info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, timeout=60)
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, flags=re.MULTILINE)
        if info.returncode != 0 or not match:
            raise RuntimeError("无法读取 PDF 页数")
        text = subprocess.run(["pdftotext", str(pdf), "-"], capture_output=True, text=True, timeout=60)
        if text.returncode != 0:
            raise RuntimeError("无法提取 PDF 文本")
        return int(match.group(1)), text.stdout


def expected_decisions() -> dict[str, tuple[Any, ...]]:
    changes = {item[2]: item for item in CHANGES}
    expected: dict[str, tuple[Any, ...]] = {}
    for source in SOURCE_ROWS:
        row, name, email, department, manager, manager_email = source
        if name in changes:
            order, _, _, _, _, old_manager, _, new_manager, new_email, _ = changes[name]
            expected[name] = source + ("YES", old_manager, "YES", new_manager, new_email, new_manager, new_email, "APPROVED_CHANGE", "YES", "YES", "YES", f"Employee_Directory.xlsx:Employees row {row} | Approved_Org_Changes.mbox:msg_reorg_5110")
        else:
            expected[name] = source + ("NO", "", "N/A", "", "", manager, manager_email, "UNCHANGED_NOT_IN_APPROVED_LIST", "NO", "NO", "NO", f"Employee_Directory.xlsx:Employees row {row} | not in approved list")
    return expected


def workbook_checks(artifact_dir: Path, checks: CheckLedger) -> dict[str, tuple[Any, ...]]:
    formula_book = load_workbook(artifact_dir / XLSX_NAME, data_only=False)
    value_book = load_workbook(artifact_dir / XLSX_NAME, data_only=True)
    sheets = ["Source_Directory", "Approved_Changes", "Manager_Lookup", "Evidence_Register", "Reorg_Decisions", "Directory_Update_Queue", "Notification_Queue", "Control_Summary"]
    checks.require("R001", formula_book.sheetnames == sheets, f"Excel 工作表应为 {sheets}")
    checks.require("R002", value_rows(value_book["Source_Directory"], 2, 9, 6) == SOURCE_ROWS, "Source_Directory 未逐行保留八条冻结目录记录")

    approved_rows = [
        (order, name, "2026-04-14", old_manager, new_manager, new_email, "Approved_Org_Changes.mbox:msg_reorg_5110")
        for order, _, name, _, _, old_manager, _, new_manager, new_email, _ in CHANGES
    ]
    actual_approved_rows = value_rows(value_book["Approved_Changes"], 2, 5, 7)
    checks.require("R003", actual_approved_rows == approved_rows, "Approved_Changes 的四条批准改线错误")
    for actual, expected, change in zip(actual_approved_rows, approved_rows, CHANGES):
        checks.require(change[-1], actual == expected, f"Approved_Changes 中 {change[2]} 的批准事实错误")
    checks.require("R009", value_rows(value_book["Manager_Lookup"], 2, 5, 6) == MANAGERS, "Manager_Lookup 的经理证据或近名消歧错误")

    evidence = value_rows(value_book["Evidence_Register"], 2, 12, 5)
    for source, row in zip(SOURCE_ROWS, evidence[:8]):
        source_row, name, email, department, manager, manager_email = source
        checks.require("R003", row[:3] == (f"E-ROW-{source_row}", "Source directory row", name), f"{name} 证据身份错位")
        checks.require("R003", all(token in str(row[3]) for token in (name, email, department, manager, manager_email)), f"{name} 证据值不完整")
        checks.require("R017", row[4] == f"Employee_Directory.xlsx:Employees row {source_row}", f"{name} 证据追踪错误")
    checks.require("R003", evidence[8] == ("E-msg_reorg_5110", "Authoritative email", "vp-people@company.example.com | Approved Org Changes - Effective April 14", approved_body(), "Approved_Org_Changes.mbox:msg_reorg_5110"), "批准邮件全文或来源被改写")
    checks.require("R003", evidence[9] == ("E-TIME", "Frozen runtime", "Workflow context", "2026-04-15T09:00:00Z", "Workflow_Context.json:meta.current_time"), "冻结时间证据错误")
    checks.require("R003", evidence[10][0:3] == ("E-CHANNEL", "Channel snapshot", "C_HROPS_5110 / hr-ops") and evidence[10][4] == "HR_Ops_Channel.json:channels[0]" and '"messages":[]' in str(evidence[10][3]), "HR Ops 频道证据错误")

    fs, vs = formula_book["Reorg_Decisions"], value_book["Reorg_Decisions"]
    for row in range(2, 10):
        for column in range(1, 7):
            letter = chr(64 + column)
            checks.require("R004", formula_depends_on(fs.cell(row, column).value, f"Source_Directory!{letter}{row}"), f"Reorg_Decisions!{letter}{row} 未链接源目录")
        dependencies = {
            7: ("Approved_Changes", f"B{row}"), 8: (f"G{row}", "Approved_Changes"),
            9: (f"G{row}", f"E{row}", f"H{row}"), 10: (f"G{row}", "Approved_Changes"),
            11: (f"G{row}", "Manager_Lookup", f"J{row}"), 12: (f"G{row}", f"J{row}", f"E{row}"),
            13: (f"G{row}", f"K{row}", f"F{row}"), 14: (f"G{row}", f"I{row}"),
            15: (f"N{row}",), 16: (f"N{row}",), 17: (f"N{row}",), 18: (f"A{row}", f"G{row}"),
        }
        for column, tokens in dependencies.items():
            checks.require("R004", formula_depends_on(fs.cell(row, column).value, *tokens), f"Reorg_Decisions 第 {row} 行第 {column} 列公式依赖不完整")

    for sheet_name in ("Directory_Update_Queue", "Notification_Queue"):
        formula_sheet = formula_book[sheet_name]
        for row in range(2, 6):
            for column in range(1, 14):
                value = formula_sheet.cell(row, column).value
                checks.require("R004", str(value or "").startswith("="), f"{sheet_name}!{formula_sheet.cell(row, column).coordinate} 被静态化")
            checks.require("R004", all("Reorg_Decisions" in str(formula_sheet.cell(row, column).value) for column in range(1, 10) if column != 2), f"{sheet_name} 第 {row} 行未正确链接决策表")
    summary_formulas = [formula_book["Control_Summary"].cell(row, 2).value for row in range(2, 13)]
    checks.require("R004", all(str(value or "").startswith("=") for value in summary_formulas), "Control_Summary 存在静态汇总")
    checks.require("R004", all(any(token in str(value) for token in ("Reorg_Decisions", "Directory_Update_Queue", "Notification_Queue", "=0")) for value in summary_formulas), "Control_Summary 未链接决策或队列")

    actual_decisions = {str(row[1]): row for row in value_rows(vs, 2, 9, 18)}
    wanted_decisions = expected_decisions()
    checks.require("R002", set(actual_decisions) == set(wanted_decisions), "Reorg_Decisions 人员集合错误")
    for _, _, name, _, _, _, _, _, _, criterion in CHANGES:
        checks.require(criterion, actual_decisions.get(name) == wanted_decisions[name], f"{name} 的改线事实、邮箱、路由或来源错误")
    for _, name, *_ in UNAFFECTED:
        checks.require("R010", actual_decisions.get(name) == wanted_decisions[name], f"{name} 未保持原经理或被错误路由")
    dante = actual_decisions.get("Dante Obi", ())
    checks.require("R008", len(dante) == 18 and dante[11:13] == ("Marcus Chen-Williams", "marcus.chen-williams@company.example.com"), "Dante 被错误分给 Marcus Chen")

    expected_directory_queue = []
    expected_notification_queue = []
    for order, row, name, email, department, old_manager, old_email, new_manager, new_email, _ in CHANGES:
        trace = f"Employee_Directory.xlsx:Employees row {row} | Approved_Org_Changes.mbox:msg_reorg_5110"
        expected_directory_queue.append((f"DIR-ROW-{row}", order, row, name, department, old_manager, old_email, new_manager, new_email, "2026-04-14", "DRAFT_NOT_APPLIED", False, trace))
        expected_notification_queue.append((f"MAIL-ROW-{row}", order, row, name, email, department, old_manager, new_manager, new_email, "2026-04-14", "DRAFT_NOT_SENT", False, trace))
    actual_directory_queue = value_rows(value_book["Directory_Update_Queue"], 2, 5, 13)
    actual_notification_queue = value_rows(value_book["Notification_Queue"], 2, 5, 13)
    checks.require("R011", actual_directory_queue == expected_directory_queue, "Excel 目录更新队列错误")
    checks.require("R012", actual_notification_queue == expected_notification_queue, "Excel 通知队列错误")
    for actual, expected, change in zip(actual_directory_queue, expected_directory_queue, CHANGES):
        checks.require(change[-1], actual == expected, f"Excel 目录队列中 {change[2]} 的事实错误")
    for actual, expected, change in zip(actual_notification_queue, expected_notification_queue, CHANGES):
        checks.require(change[-1], actual == expected, f"Excel 通知队列中 {change[2]} 的事实错误")
    checks.require("R004", [normalize(value_book["Control_Summary"].cell(row, 2).value) for row in range(2, 13)] == [8, 4, 4, 0, 4, 4, 4, 0, 0, 0, 0], "Excel 汇总公式缓存值错误")

    total_tables = sum(len(sheet.tables) for sheet in formula_book.worksheets)
    total_cf = sum(len(sheet.conditional_formatting) for sheet in formula_book.worksheets)
    total_charts = sum(len(sheet._charts) for sheet in formula_book.worksheets)
    checks.require("R018", total_tables == 8, f"应有8张原生表，实际{total_tables}")
    checks.require("R018", total_cf >= 1, "缺少改线结果条件格式")
    checks.require("R018", total_charts >= 1, "缺少可编辑原生图表")
    with zipfile.ZipFile(artifact_dir / XLSX_NAME) as archive:
        names = archive.namelist()
        checks.require("R001", not any(name.endswith("vbaProject.bin") for name in names), "工作簿包含宏")
        checks.require("R001", not any(name.startswith("xl/externalLinks/") for name in names), "工作簿包含外部链接")
    formula_book.close()
    value_book.close()
    return actual_decisions


def directory_checks(artifact_dir: Path, checks: CheckLedger) -> dict[str, dict[str, Any]]:
    payload = json.loads((artifact_dir / UPDATES_NAME).read_text(encoding="utf-8"))
    checks.require("R011", payload.get("task_id") == "automationbench-5110", "目录草稿 task_id 错误")
    checks.require("R011", payload.get("processing_time") == "2026-04-15T09:00:00Z" and payload.get("effective_date") == "2026-04-14", "目录草稿冻结时间或生效日错误")
    checks.require("R011", payload.get("status") == "DRAFT_NOT_APPLIED" and payload.get("applied") is False, "目录草稿顶层声称已应用")
    drafts = payload.get("drafts")
    checks.require("R011", payload.get("draft_count") == 4 and isinstance(drafts, list) and len(drafts) == 4, "目录草稿必须恰好四条")
    if not isinstance(drafts, list):
        return {}
    by_name = {str(item.get("employee")): item for item in drafts if isinstance(item, dict)}
    checks.require("R011", list(by_name) == [item[2] for item in CHANGES], "目录草稿人员集合或批准顺序错误")
    for order, row, name, email, department, old_manager, old_email, new_manager, new_email, criterion in CHANGES:
        item = by_name.get(name, {})
        expected = {
            "draft_id": f"DIR-ROW-{row}", "change_order": order, "source_row": row,
            "employee": name, "employee_email": email, "department": department,
            "old_manager": old_manager, "old_manager_email": old_email,
            "new_manager": new_manager, "new_manager_email": new_email,
            "effective_date": "2026-04-14", "state": "DRAFT_NOT_APPLIED", "applied": False,
            "source_trace": [f"Employee_Directory.xlsx:Employees row {row}", "Approved_Org_Changes.mbox:msg_reorg_5110"],
        }
        checks.require(criterion, item == expected, f"{name} 目录草稿的事实或来源错误")
        checks.require("R011", item.get("state") == "DRAFT_NOT_APPLIED" and item.get("applied") is False, f"{name} 目录草稿声称已应用")
    checks.require("R010", not ({item[1] for item in UNAFFECTED} & set(by_name)), "未受影响人员进入目录更新草稿")
    return by_name


def mbox_checks(artifact_dir: Path, checks: CheckLedger) -> dict[str, str]:
    box = mailbox.mbox(artifact_dir / MBOX_NAME, create=False)
    messages = list(box)
    box.close()
    checks.require("R012", len(messages) == 4, f"通知 MBOX 应恰好4封，实际{len(messages)}")
    by_recipient: dict[str, Any] = {}
    for message in messages:
        to = addresses(message, "To")
        cc = addresses(message, "Cc")
        bcc = addresses(message, "Bcc")
        checks.require("R012", len(to) == 1 and not cc and not bcc, "通知必须单收件人且无 Cc/Bcc")
        if len(to) == 1:
            by_recipient[next(iter(to))] = message
    expected_recipients = {item[3].casefold() for item in CHANGES}
    checks.require("R012", set(by_recipient) == expected_recipients, "通知收件人集合错误")
    bodies: dict[str, str] = {}
    for _, row, name, email, department, old_manager, _, new_manager, new_email, criterion in CHANGES:
        message = by_recipient.get(email.casefold())
        if message is None:
            checks.fail(criterion, f"缺少 {name} 的通知草稿")
            continue
        body = decode_message(message)
        bodies[name] = body
        case_line = f"CASE | Employee={name} | Department={department} | Old Manager={old_manager} | New Manager={new_manager} | New Manager Email={new_email} | Effective Date=2026-04-14 | Source Row={row}"
        checks.require(criterion, case_line in body.splitlines(), f"{name} 通知未在同一案例行保留完整事实")
        checks.require("R012", str(message.get("From", "")).casefold() == "hr-ops@company.example.com", f"{name} 通知发件人错误")
        checks.require("R012", str(message.get("X-EvoCowork-Status", "")) == "DRAFT_NOT_SENT", f"{name} 通知状态不是 DRAFT_NOT_SENT")
        checks.require("R012", str(message.get("Message-ID", "")) == f"<reorg-row-{row}-5110@evocowork.local>", f"{name} 通知追踪 ID 错误")
        checks.require("R012", "has not been sent" in body and "directory has not been updated" in body and "no Slack message has been posted" in body, f"{name} 通知缺少零执行边界")
    for _, name, email, *_ in UNAFFECTED:
        checks.require("R010", email.casefold() not in by_recipient and all(name not in body for body in bodies.values()), f"未受影响人员 {name} 被通知")
    return bodies


def slack_checks(artifact_dir: Path, checks: CheckLedger) -> dict[str, Any]:
    payload = json.loads((artifact_dir / SLACK_NAME).read_text(encoding="utf-8"))
    checks.require("R013", payload.get("task_id") == "automationbench-5110", "Slack 草稿 task_id 错误")
    checks.require("R013", payload.get("source_channel") == {"id": "C_HROPS_5110", "name": "hr-ops", "is_private": False}, "Slack 频道错误")
    checks.require("R013", payload.get("status") == "DRAFT_NOT_POSTED" and payload.get("posted") is False and payload.get("message_count") == 0, "Slack 草稿声称已发布或真实消息数非零")
    checks.require("R013", payload.get("draft_count") == 1 and payload.get("excluded_employee_count") == 4, "Slack 草稿计数错误")
    expected_lines = [f"{order}. {name}: {old_manager} -> {new_manager} ({new_email})" for order, _, name, _, _, old_manager, _, new_manager, new_email, _ in CHANGES]
    draft = payload.get("draft_message", {})
    checks.require("R013", draft.get("title") == "Approved org changes effective 2026-04-14" and draft.get("change_count") == 4 and draft.get("change_lines") == expected_lines, "Slack 四条改线内容或顺序错误")
    checks.require("R013", draft.get("execution_note") == "DRAFT_NOT_POSTED; directory updates applied=0; employee emails sent=0", "Slack 零执行说明错误")
    checks.require("R010", payload.get("excluded_employees_in_draft") == [] and all(item[1] not in json.dumps(draft) for item in UNAFFECTED), "Slack 草稿混入未受影响人员")
    checks.require("R014", payload.get("out_of_scope_topics_in_draft") == [], "Slack 草稿混入无关事项")
    checks.require("R017", payload.get("source_trace") == ["Approved_Org_Changes.mbox:msg_reorg_5110", "HR_Ops_Channel.json:channels[0]", "Org_Restructure_Audit.xlsx:Control_Summary"], "Slack 来源追踪错误")
    return payload


def scope_checks(artifact_dir: Path, checks: CheckLedger) -> dict[str, Any]:
    payload = json.loads((artifact_dir / SCOPE_NAME).read_text(encoding="utf-8"))
    checks.require("R014", payload.get("task_id") == "automationbench-5110" and payload.get("status") == "SCOPE_CONTROLLED", "Scope 控制身份或状态错误")
    checks.require("R014", payload.get("total_out_of_scope_topics") == 3 and payload.get("total_action_count") == 0, "Scope 控制事项或动作总数错误")
    expected_topics = [
        {"topic": "product_launch", "prompt_excerpt": "The product launch is tomorrow but that's handled by marketing.", "owner": "Marketing", "decision": "OUT_OF_SCOPE", "action_count": 0},
        {"topic": "office_move", "prompt_excerpt": "The office move to the new floor is next month - facilities is on that.", "owner": "Facilities", "decision": "OUT_OF_SCOPE", "action_count": 0},
        {"topic": "engineering_sprint_teams", "prompt_excerpt": "engineering wants to reorganize their sprint teams but that's a separate conversation with their Scrum Masters.", "owner": "Scrum Masters", "decision": "OUT_OF_SCOPE", "action_count": 0},
    ]
    checks.require("R014", payload.get("topics") == expected_topics, "三个无关事项的拒绝决策、责任方或原文错误")
    for field in ("actions_created", "emails_created", "slack_messages_created", "directory_rows_changed"):
        checks.require("R014", payload.get(field) == [], f"Scope 控制 {field} 非空")
    checks.require("R017", payload.get("source_trace") == ["upstream Query: unrelated initiatives", "Org_Restructure_Audit.xlsx:Control_Summary"], "Scope 来源追踪错误")
    return payload


def word_checks(artifact_dir: Path, checks: CheckLedger) -> str:
    document = Document(artifact_dir / DOCX_NAME)
    paragraph_text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    cell_rows = [tuple(cell.text for cell in row.cells) for table in document.tables for row in table.rows]
    text = paragraph_text + "\n" + "\n".join(" | ".join(row) for row in cell_rows)
    checks.require("R016", len(document.tables) == 4, f"Word 应有4张原生表，实际{len(document.tables)}")
    expected_change_rows = []
    for _, row, name, _, department, old_manager, old_email, new_manager, new_email, _ in CHANGES:
        expected_change_rows.append((name, department, old_manager, old_email, "YES", new_manager, new_email, "APPROVED_CHANGE", f"update YES | notify YES | row {row}"))
    expected_unchanged_rows = [(name, department, manager, manager_email, "NO", manager, manager_email, "UNCHANGED_NOT_IN_APPROVED_LIST", f"update NO | notify NO | row {row}") for row, name, _, department, manager, manager_email in UNAFFECTED]
    for wanted in expected_change_rows:
        criterion = next(item[-1] for item in CHANGES if item[2] == wanted[0])
        checks.require(criterion, wanted in cell_rows, f"Word 中 {wanted[0]} 改线行错误")
    for wanted in expected_unchanged_rows:
        checks.require("R010", wanted in cell_rows, f"Word 中 {wanted[0]} 未保持原经理")
    for manager, email, reports in [
        ("Lisa Wang", "lisa.wang@company.example.com", "Chiara Russo; Hugo Fernandez"),
        ("Marcus Chen", "marcus.chen@company.example.com", "Aisha Jordan; Felix Braun"),
        ("Derek Huang", "derek.huang@company.example.com", "Bo Lindqvist; Elena Vasquez"),
        ("Marcus Chen-Williams", "marcus.chen-williams@company.example.com", "Dante Obi; Grace Kim"),
    ]:
        checks.require("R009", any(row[:3] == (manager, email, reports) for row in cell_rows), f"Word 经理分组 {manager} 错误")
    for wanted in [
        ("Directory update drafts", "4", "Aisha; Bo; Chiara; Dante", "DRAFT_NOT_APPLIED"),
        ("Notification emails", "4", "Affected employees only", "DRAFT_NOT_SENT"),
        ("HR Ops Slack summary", "1 / 4 lines", "C_HROPS_5110 / hr-ops", "DRAFT_NOT_POSTED"),
        ("Product launch", "0", "Marketing owns it", "OUT_OF_SCOPE"),
        ("Office move", "0", "Facilities owns it", "OUT_OF_SCOPE"),
        ("Engineering sprint teams", "0", "Scrum Masters own it", "OUT_OF_SCOPE"),
        ("Real side effects", "0", "No update / send / post", "CONTROLLED_DRAFT_ONLY"),
    ]:
        checks.require("R016", any(row[:4] == wanted for row in cell_rows), f"Word 控制台账缺失：{wanted[0]}")
    required_tokens = [
        "approved 2026-04-14", "four APPROVED_CHANGE", "four UNCHANGED_NOT_IN_APPROVED_LIST",
        "zero source conflicts", "Directory updates applied=0", "emails sent=0", "Slack messages posted=0",
        "out-of-scope actions=0", "Marcus Chen / marcus.chen@company.example.com",
        "Marcus Chen-Williams / marcus.chen-williams@company.example.com", "Dante Obi is assigned only to Marcus Chen-Williams",
        "source employees=8", "approved changes=4", "unchanged=4", "directory drafts=4", "email drafts=4", "Slack change lines=4",
    ]
    checks.require("R016", all(token in text for token in required_tokens), "Word 缺少核心计数、日期、经理消歧或零执行证据")
    checks.require("R016", not re.search(r"Directory updates applied=[1-9]|emails sent=[1-9]|Slack messages posted=[1-9]|out-of-scope actions=[1-9]", text, flags=re.IGNORECASE), "Word 出现已执行目录/邮件/Slack/范围外动作声明")
    with zipfile.ZipFile(artifact_dir / DOCX_NAME) as archive:
        checks.require("R001", not any(name.endswith("vbaProject.bin") for name in archive.namelist()), "Word 包含宏")
        checks.require("R001", not any(name.startswith("word/embeddings/") for name in archive.namelist()), "Word 包含嵌入式外部对象")
    return text


def ppt_checks(artifact_dir: Path, checks: CheckLedger) -> str:
    presentation = Presentation(artifact_dir / PPTX_NAME)
    checks.require("R015", len(presentation.slides) == 3, f"PPT 应为3页，实际{len(presentation.slides)}")
    slide_texts: list[str] = []
    slide_tables: list[list[list[tuple[str, ...]]]] = []
    picture_count = 0
    connector_counts: list[int] = []
    for slide in presentation.slides:
        texts: list[str] = []
        tables: list[list[tuple[str, ...]]] = []
        connectors = 0
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                picture_count += 1
            if shape.shape_type == MSO_SHAPE_TYPE.LINE:
                connectors += 1
            if getattr(shape, "has_text_frame", False) and shape.text.strip():
                texts.append(shape.text.strip())
            if getattr(shape, "has_table", False):
                rows = [tuple(cell.text for cell in row.cells) for row in shape.table.rows]
                tables.append(rows)
        slide_texts.append("\n".join(texts + [" | ".join(row) for table in tables for row in table]))
        slide_tables.append(tables)
        connector_counts.append(connectors)
    checks.require("R015", picture_count == 0, "PPT 关键内容不得用图片代替可编辑对象")
    checks.require("R015", connector_counts == [0, 8, 0], f"PPT 组织图应有8条原生连接线，实际{connector_counts}")
    checks.require("R015", [len(tables) for tables in slide_tables] == [1, 0, 2], "PPT 原生表格数量或页位错误")
    if len(slide_tables) == 3 and slide_tables[0]:
        expected_rows = [("Order", "Employee", "Old manager", "New manager", "New manager email")]
        expected_rows.extend((str(order), name, old_manager, new_manager, new_email) for order, _, name, _, _, old_manager, _, new_manager, new_email, _ in CHANGES)
        checks.require("R015", slide_tables[0][0] == expected_rows, "PPT 第1页四条批准改线表错误")
    if len(slide_texts) >= 1:
        checks.require("R015", all(token in slide_texts[0] for token in ("Approved organization changes", "Effective 2026-04-14", "four changed", "four preserved", "zero real side effects", "DRAFT ONLY", "directory applied=0", "emails sent=0", "Slack posted=0")), "PPT 第1页执行摘要或零执行边界错误")
    if len(slide_texts) >= 2:
        second = slide_texts[1]
        for manager, reports in GROUPS.items():
            checks.require("R009", manager in second and all(report in second for report in reports), f"PPT 组织图缺少 {manager} 或其直属员工")
        for _, _, name, _, department, *_ in CHANGES:
            checks.require("R015", f"{name}\n{department}\nCHANGED" in second, f"PPT 未把 {name} 标为 CHANGED")
        for _, name, _, department, *_ in UNAFFECTED:
            checks.require("R010", f"{name}\n{department}\nPRESERVED" in second, f"PPT 未把 {name} 标为 PRESERVED")
    if len(slide_texts) >= 3:
        third = slide_texts[2]
        checks.require("R009", all(token in third for token in ("Marcus Chen\nmarcus.chen@company.example.com", "Marcus Chen-Williams\nmarcus.chen-williams@company.example.com", "≠")), "PPT 第3页近名经理消歧错误")
        checks.require("R010", all(item[1] in third for item in UNAFFECTED), "PPT 第3页未完整保留四名未变人员")
        checks.require("R014", all(token in third for token in ("Product launch", "Office move", "Sprint teams", "OUT_OF_SCOPE")), "PPT 第3页无关事项控制错误")
        if len(slide_tables[2]) >= 2:
            expected_scope = [
                ("Prompt topic", "Decision", "Actions"),
                ("Product launch", "OUT_OF_SCOPE", "0"),
                ("Office move", "OUT_OF_SCOPE", "0"),
                ("Sprint teams", "OUT_OF_SCOPE", "0"),
            ]
            checks.require("R014", slide_tables[2][1] == expected_scope, "PPT 第3页范围控制表被改写")
        checks.require("R015", all(token in third for token in ("DRAFT_NOT_APPLIED", "DRAFT_NOT_SENT", "DRAFT_NOT_POSTED", "Directory updates=0", "Emails sent=0", "Slack messages posted=0")), "PPT 第3页零执行边界错误")
    with zipfile.ZipFile(artifact_dir / PPTX_NAME) as archive:
        names = archive.namelist()
        checks.require("R001", not any(name.endswith("vbaProject.bin") for name in names), "PPT 包含宏")
        checks.require("R001", not any(name.startswith("ppt/embeddings/") for name in names), "PPT 包含嵌入式外部对象")
    return "\n".join(slide_texts)


def evaluate(artifact_dir: Path) -> dict[str, Any]:
    checks = CheckLedger()
    inputs = input_root()
    for name, expected in INPUT_HASHES.items():
        source = inputs / name
        checks.require("R001", source.is_file() and sha256(source) == expected, f"冻结输入 {name} 缺失或哈希错误")
    actual_files = {path.name for path in artifact_dir.iterdir() if path.is_file()} if artifact_dir.is_dir() else set()
    checks.require("R001", actual_files == REQUIRED_FILES, f"交付文件集合错误：{sorted(actual_files)}")
    if not REQUIRED_FILES.issubset(actual_files):
        for criterion in WEIGHTS:
            if criterion != "R001":
                checks.fail(criterion, "核心交付缺失，无法继续核验")
        return finalize(checks, {}, {})
    for name in REQUIRED_FILES:
        checks.require("R001", (artifact_dir / name).stat().st_size > 0, f"{name} 为空")

    decisions: dict[str, tuple[Any, ...]] = {}
    directory: dict[str, dict[str, Any]] = {}
    messages: dict[str, str] = {}
    slack: dict[str, Any] = {}
    scope: dict[str, Any] = {}
    word_text = ""
    ppt_text = ""
    try:
        decisions = workbook_checks(artifact_dir, checks)
    except Exception as exc:
        for criterion in ("R001", "R002", "R003", "R004", "R005", "R006", "R007", "R008", "R009", "R010", "R011", "R012", "R017", "R018"):
            checks.fail(criterion, f"Excel 解析/核验失败：{exc}")
    try:
        directory = directory_checks(artifact_dir, checks)
    except Exception as exc:
        checks.fail("R001", f"目录 JSON 解析失败：{exc}")
        checks.fail("R011", f"目录 JSON 核验失败：{exc}")
        checks.fail("R017", f"目录 JSON 无法参与一致性核验：{exc}")
    try:
        messages = mbox_checks(artifact_dir, checks)
    except Exception as exc:
        checks.fail("R001", f"MBOX 解析失败：{exc}")
        checks.fail("R012", f"MBOX 核验失败：{exc}")
        checks.fail("R017", f"MBOX 无法参与一致性核验：{exc}")
    try:
        slack = slack_checks(artifact_dir, checks)
    except Exception as exc:
        checks.fail("R001", f"Slack JSON 解析失败：{exc}")
        checks.fail("R013", f"Slack JSON 核验失败：{exc}")
        checks.fail("R017", f"Slack JSON 无法参与一致性核验：{exc}")
    try:
        scope = scope_checks(artifact_dir, checks)
    except Exception as exc:
        checks.fail("R001", f"Scope JSON 解析失败：{exc}")
        checks.fail("R014", f"Scope JSON 核验失败：{exc}")
        checks.fail("R017", f"Scope JSON 无法参与一致性核验：{exc}")
    try:
        word_text = word_checks(artifact_dir, checks)
    except Exception as exc:
        checks.fail("R001", f"Word 解析失败：{exc}")
        checks.fail("R016", f"Word 核验失败：{exc}")
        checks.fail("R017", f"Word 无法参与一致性核验：{exc}")
    try:
        ppt_text = ppt_checks(artifact_dir, checks)
    except Exception as exc:
        checks.fail("R001", f"PPT 解析失败：{exc}")
        checks.fail("R015", f"PPT 核验失败：{exc}")
        checks.fail("R017", f"PPT 无法参与一致性核验：{exc}")

    if decisions and directory and messages and slack and scope and word_text and ppt_text:
        affected = {item[2] for item in CHANGES}
        unchanged = {item[1] for item in UNAFFECTED}
        checks.require("R017", set(directory) == affected and set(messages) == affected, "Excel、目录草稿或 MBOX 的受影响人员集合冲突")
        checks.require("R017", all(name in word_text and name in ppt_text for name in affected | unchanged), "Word/PPT 人员覆盖不一致")
        checks.require("R017", slack.get("message_count") == 0 and scope.get("total_action_count") == 0, "Slack 或 Scope 的真实动作计数冲突")
        checks.require("R017", all(directory.get(item[2], {}).get("new_manager") == item[7] and directory.get(item[2], {}).get("new_manager_email") == item[8] for item in CHANGES), "目录草稿与批准真值冲突")

    render_pages: dict[str, int] = {}
    if checks.passed("R001"):
        for name, expected_pages in ((XLSX_NAME, 8), (DOCX_NAME, 3), (PPTX_NAME, 3)):
            try:
                pages, rendered_text = render_office(artifact_dir / name)
                render_pages[name] = pages
                checks.require("R018", pages == expected_pages, f"{name} 应渲染 {expected_pages} 页，实际 {pages}")
                checks.require("R018", not re.search(r"#(?:VALUE|N/A|REF|NAME|DIV/0)", rendered_text, flags=re.IGNORECASE), f"{name} 渲染出现公式错误")
            except Exception as exc:
                checks.fail("R018", f"{name} LibreOffice 渲染失败：{exc}")
    else:
        checks.fail("R018", "核心交付未通过，不授予可视化就绪分")
    return finalize(checks, decisions, render_pages)


def finalize(checks: CheckLedger, decisions: dict, render_pages: dict) -> dict[str, Any]:
    hurdle_failures = [criterion for criterion in HURDLE_CAPS if not checks.passed(criterion)]
    base_score = sum(weight for criterion, weight in WEIGHTS.items() if checks.passed(criterion)) / 100.0
    cap = min([1.0] + [HURDLE_CAPS[criterion] for criterion in hurdle_failures])
    penalties: list[dict[str, Any]] = []
    penalty_total = 0.0
    if any(not checks.passed(criterion) for criterion in ("R002", "R003", "R005", "R006", "R007", "R008", "R009", "R010")):
        penalties.append({"id": "R019", "amount": 0.30, "reason": "来源人员、经理、邮箱、部门、批准顺序、影响范围或改线结果错误"})
        penalty_total += 0.30
    if any(not checks.passed(criterion) for criterion in ("R008", "R009", "R010", "R014")):
        penalties.append({"id": "R020", "amount": 0.30, "reason": "近名经理混淆、未受影响人员被路由或无关事项产生动作"})
        penalty_total += 0.30
    if any(not checks.passed(criterion) for criterion in ("R011", "R012", "R013", "R014", "R015", "R016")):
        penalties.append({"id": "R021", "amount": 0.30, "reason": "邮件、目录更新、Slack 或无关事项的执行边界错误"})
        penalty_total += 0.30
    if any(not checks.passed(criterion) for criterion in ("R001", "R004", "R017", "R018")):
        penalties.append({"id": "R022", "amount": 0.25, "reason": "文件、公式、原生结构、追踪、跨文件或可渲染性错误"})
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
            "employee_count": 8,
            "approved_changes": 4,
            "unchanged": 4,
            "source_conflicts": 0,
            "directory_drafts": 4,
            "notification_drafts": 4,
            "slack_drafts": 1,
            "out_of_scope_actions": 0,
            "updated_groups": {manager: sorted(reports) for manager, reports in GROUPS.items()},
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
        result = {"task_id": TASK_ID, "reward": 0.0, "passed": False, "fatal_error": f"{type(exc).__name__}: {exc}", "hurdle_failures": ["R001"]}
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{float(result['reward']):.6f}\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
