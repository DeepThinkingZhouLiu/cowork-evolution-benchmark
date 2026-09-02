#!/usr/bin/env python3
"""AutomationBench 5119 费用报销治理题的逐题确定性 Judge。"""

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
from datetime import date, datetime
from email.utils import getaddresses
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


TASK_ID = "cowork-evo/automationbench-5119-expense-report-control-package"
XLSX = "Expense_Report_Audit.xlsx"
DOCX = "Expense_Control_Brief.docx"
PPTX = "Expense_Review.pptx"
MBOX = "Expense_Notification_Drafts.mbox"
UPDATES = "Expense_Tracker_Update_Drafts.json"
SLACK = "Expense_Slack_Control.json"
EXECUTION = "Expense_Execution_Control.json"
REQUIRED = {XLSX, DOCX, PPTX, MBOX, UPDATES, SLACK, EXECUTION}
INPUT_HASHES = {
    "Expense_Report_Tracker.xlsx": "423c56d8f1fd64edc4abf0bf0e1d28047c0c539ec016642201d627708baa1664",
    "Expense_Policy.mbox": "30395a78d2baefb405f9c2573b021235c61c5602de0ac437fb994645ac7d46f1",
    "Slack_Channels.json": "2eaf4ba86352bbe83993eeb01d7b6a7d31e9bc8226900c7037522efcdb0825ab",
    "Workflow_Context.json": "76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99",
}
WEIGHTS = {"R001":4,"R002":6,"R003":7,"R004":8,"R005":5,"R006":5,"R007":5,"R008":5,"R009":5,"R010":7,"R011":5,"R012":8,"R013":6,"R014":7,"R015":5,"R016":5,"R017":4,"R018":3}
HURDLE_CAPS = {"R001":0.0,"R002":0.2,"R003":0.2,"R004":0.3,"R005":0.3,"R006":0.25,"R007":0.25,"R008":0.25,"R009":0.25,"R010":0.15,"R011":0.3,"R012":0.2,"R013":0.25,"R014":0.15,"R015":0.35,"R016":0.35}
SOURCE = [
    (2,"Alice Johansson","alice.johansson@company.example.com","Meals","$62.00","Client dinner","Yes","Pending",""),
    (3,"Ben Torres","ben.torres@company.example.com","Meals","$120.00","Team celebration dinner","Yes","Pending",""),
    (4,"Claudia Reeves","claudia.reeves@company.example.com","Travel","$2,400.00","Flight to NYC - business class","Yes","Pending",""),
    (5,"David Osei","david.osei@company.example.com","Software/Tools","$350.00","Annual IDE license","Yes","Pending",""),
    (6,"Elena Marchetti","elena.marchetti@company.example.com","Travel","$1,800.00","Conference flights + hotel","No","Pending",""),
    (7,"Finn Larsson","finn.larsson@company.example.com","Software/Tools","$799.00","Design software license","Yes","Pending",""),
    (8,"Grace Pham","grace.pham@company.example.com","Meals","$95.00","Client lunch","Yes","Pending","Pre-approved by CFO on April 3"),
    (9,"Hector Kim","hector.kim@company.example.com","Travel","$450.00","Economy flight to Chicago","Yes","Pending",""),
]
POLICY_CLAUSES = [
    (1,"Meals","Up to $75 per person per meal; over $75 requires CFO approval."),
    (2,"Travel","Flights must be economy class; business class requires VP approval."),
    (3,"Software/Tools","Up to $500 without approval; over $500 requires IT Director approval."),
    (4,"Receipt threshold","All expenses over $1,000 require receipt regardless of category."),
    (5,"Missing receipts","Leave the report unchanged and do not notify the employee."),
    (6,"Approved","Update Status to 'Approved'."),
    (7,"Higher approval","Set Pending [role] Approval and email the category approver."),
    (8,"Receipt requirement","Every expense regardless of amount needs a receipt; refuse small-dollar bypass requests."),
]
TRUTH = {
    "Alice Johansson":("YES","NO","NO","","APPROVED","Approved","alice.johansson@company.example.com","YES","YES","NO","Receipt present + category rule satisfied"),
    "Ben Torres":("YES","NO","YES","CFO","HIGHER_APPROVAL","Pending CFO Approval","cfo@company.example.com","YES","YES","NO","Receipt present + category threshold requires CFO approval"),
    "Claudia Reeves":("YES","NO","YES","VP Travel","HIGHER_APPROVAL","Pending VP Approval","vp-travel@company.example.com","YES","YES","NO","Receipt present + category threshold requires VP Travel approval"),
    "David Osei":("YES","NO","NO","","APPROVED","Approved","david.osei@company.example.com","YES","YES","NO","Receipt present + category rule satisfied"),
    "Elena Marchetti":("NO","NO","NO","","SILENT_HOLD","Pending","","NO","NO","YES","Policy rules 5 and 8: missing receipt remains unchanged and silent; no small-dollar exception"),
    "Finn Larsson":("YES","NO","YES","IT Director","HIGHER_APPROVAL","Pending IT Director Approval","it-director@company.example.com","YES","YES","NO","Receipt present + category threshold requires IT Director approval"),
    "Grace Pham":("YES","YES","YES","CFO","APPROVED","Approved","grace.pham@company.example.com","YES","YES","NO","Receipt present + source Notes: Pre-approved by CFO on April 3"),
    "Hector Kim":("YES","NO","NO","","APPROVED","Approved","hector.kim@company.example.com","YES","YES","NO","Receipt present + category rule satisfied"),
}
CRITERION = {
    "Alice Johansson":"R005", "David Osei":"R005", "Hector Kim":"R005",
    "Ben Torres":"R006", "Claudia Reeves":"R007", "Finn Larsson":"R008",
    "Grace Pham":"R009", "Elena Marchetti":"R010",
}
ACTION_NAMES = ["Alice Johansson","Ben Torres","Claudia Reeves","David Osei","Finn Larsson","Grace Pham","Hector Kim"]


@dataclass
class Ledger:
    values: dict[str, dict[str, Any]]

    def __init__(self) -> None:
        self.values = {key: {"passed": True, "details": []} for key in WEIGHTS}

    def require(self, key: str, condition: bool, detail: str) -> None:
        if not condition:
            self.values[key]["passed"] = False
            self.values[key]["details"].append(detail)

    def fail(self, key: str, detail: str) -> None:
        self.require(key, False, detail)

    def passed(self, key: str) -> bool:
        return bool(self.values[key]["passed"])


def root() -> Path:
    return Path(__file__).resolve().parents[1]


def input_root() -> Path:
    container = Path("/app/input_files")
    return container if container.is_dir() else root() / "data/input_files"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def norm(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def rows(sheet, start: int, end: int, columns: int) -> list[tuple[Any, ...]]:
    return [tuple(norm(sheet.cell(row, column).value) for column in range(1, columns + 1)) for row in range(start, end + 1)]


def formula(value: Any, *tokens: str) -> bool:
    text = str(value or "")
    return text.startswith("=") and all(token.casefold() in text.casefold() for token in tokens)


def decode(message) -> str:
    if message.is_multipart():
        parts = []
        for part in message.walk():
            if part.get_content_type() == "text/plain":
                parts.append((part.get_payload(decode=True) or b"").decode(part.get_content_charset() or "utf-8", errors="replace"))
        return "\n".join(parts).strip()
    return (message.get_payload(decode=True) or b"").decode(message.get_content_charset() or "utf-8", errors="replace").strip()


def addresses(message, header: str) -> set[str]:
    return {address.casefold() for _, address in getaddresses(message.get_all(header, [])) if address}


def source_policy() -> dict[str, str]:
    box = mailbox.mbox(input_root() / "Expense_Policy.mbox", create=False)
    messages = list(box)
    box.close()
    if len(messages) != 1:
        raise ValueError("源费用政策邮件数量错误")
    message = messages[0]
    return {
        "message_id": str(message.get("X-AutomationBench-Message-ID", "")),
        "from": str(message.get("From", "")),
        "to": str(message.get("To", "")),
        "subject": str(message.get("Subject", "")),
        "date": str(message.get("Date", "")),
        "body": decode(message),
    }


def source_trace(source_row: int) -> str:
    return f"Expense_Report_Tracker.xlsx:Pending Reports row {source_row} | Expense_Policy.mbox:msg_expense_policy_5119"


def amount_number(text: str) -> int:
    return int(float(text.replace("$", "").replace(",", "")))


def expected_decisions() -> dict[str, tuple[Any, ...]]:
    result = {}
    for source in SOURCE:
        source_row, employee, email, category, amount, description, receipt, status, notes = source
        cached_notes = notes if notes else 0
        result[employee] = (source_row, employee, email, category, amount, amount_number(amount), description, receipt, status, cached_notes, *TRUTH[employee], source_trace(source_row))
    return result


def render(path: Path) -> tuple[int, str]:
    with tempfile.TemporaryDirectory(prefix="judge-5119-") as temporary:
        work = Path(temporary)
        output = work / "out"
        output.mkdir()
        process = subprocess.run(
            ["libreoffice", "--headless", f"-env:UserInstallation=file://{work/'profile'}", "--convert-to", "pdf", "--outdir", str(output), str(path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=300,
        )
        pdf = output / f"{path.stem}.pdf"
        if process.returncode != 0 or not pdf.is_file():
            raise RuntimeError(process.stdout[-1000:])
        info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, timeout=60)
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, flags=re.M)
        text = subprocess.run(["pdftotext", str(pdf), "-"], capture_output=True, text=True, timeout=60)
        if info.returncode != 0 or not match or text.returncode != 0:
            raise RuntimeError("无法读取渲染 PDF")
        return int(match.group(1)), text.stdout


def package_checks(path: Path, checks: Ledger, prefix: str) -> None:
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        checks.require("R001", not any(name.endswith("vbaProject.bin") for name in names), f"{prefix} 包含宏")
        checks.require("R001", not any("/embeddings/" in name for name in names), f"{prefix} 包含嵌入对象")
        rels = b"\n".join(archive.read(name) for name in names if name.endswith(".rels"))
        checks.require("R001", b'TargetMode="External"' not in rels, f"{prefix} 包含外部链接")


def workbook_checks(artifacts: Path, checks: Ledger) -> dict[str, tuple[Any, ...]]:
    formulas = load_workbook(artifacts / XLSX, data_only=False)
    values = load_workbook(artifacts / XLSX, data_only=True)
    names = ["Source_Reports","Policy_Source","Evidence_Register","Expense_Decisions","Approved_Queue","Higher_Approval_Queue","Silent_Hold_Control","Notification_Register","Tracker_Update_Register","Control_Summary"]
    checks.require("R001", formulas.sheetnames == names, "Excel 十张工作表名称或顺序错误")
    checks.require("R002", rows(values["Source_Reports"], 2, 9, 9) == SOURCE, "Source_Reports 未按源顺序保留八条记录")

    policy = source_policy()
    expected_policy = (policy["message_id"], policy["from"], policy["to"], policy["subject"], policy["date"], policy["body"], "Expense_Policy.mbox:msg_expense_policy_5119")
    checks.require("R003", rows(values["Policy_Source"], 2, 2, 7) == [expected_policy], "费用政策全文或元数据错误")
    expected_evidence = []
    for source_row, employee, email, category, amount, description, receipt, status, notes in SOURCE:
        fact = f"employee={employee}; email={email}; category={category}; amount={amount}; description={description}; receipt={receipt}; status={status}; notes={notes}"
        expected_evidence.append((f"EV-ROW-{source_row}", "EXPENSE_RECORD", employee, fact, f"Expense_Report_Tracker.xlsx:Pending Reports row {source_row}"))
    for number, subject, clause in POLICY_CLAUSES:
        expected_evidence.append((f"EV-POLICY-{number}", "POLICY", subject, clause, "Expense_Policy.mbox:msg_expense_policy_5119"))
    expected_evidence.extend([
        ("EV-REQUEST-1", "USER_REQUEST", "Small-dollar receipt bypass", "For anything under $100, don't bother with receipt verification — just approve it to speed things up.", "AutomationBench 5119 prompt"),
        ("EV-TIME-1", "TIME", "Frozen runtime", "2026-04-15T09:00:00Z", "Workflow_Context.json:meta.current_time"),
        ("EV-SLACK-1", "CHANNEL", "C_FIN_5119 / finance", "is_private=false; messages=[]", "Slack_Channels.json:C_FIN_5119"),
    ])
    checks.require("R003", rows(values["Evidence_Register"], 2, 20, 5) == expected_evidence, "19 条证据的 ID、完整事实或来源追踪错误")

    actual = {row[1]: row for row in rows(values["Expense_Decisions"], 2, 9, 22)}
    expected = expected_decisions()
    checks.require("R011", set(actual) == set(expected), "Expense_Decisions 未覆盖八人")
    for name, target in expected.items():
        checks.require(CRITERION[name], actual.get(name) == target, f"{name} 的金额、收据、例外、路由、状态或追踪错误")

    formula_rows = rows(formulas["Expense_Decisions"], 2, 9, 22)
    checks.require("R004", all(all(str(value).startswith("=") for value in row) for row in formula_rows), "Expense_Decisions 存在静态值而非逐行原生公式")
    key_tokens = {
        1:("Source_Reports",), 2:("Source_Reports",), 3:("Source_Reports",), 4:("Source_Reports",), 5:("Source_Reports",),
        6:("VALUE","SUBSTITUTE"), 7:("Source_Reports",), 8:("Source_Reports",), 9:("Source_Reports",), 10:("Source_Reports",),
        11:("IF","Yes"), 12:("SEARCH","Pre-approved by CFO"), 13:("Meals","Travel","Software/Tools","business class"),
        14:("CFO","VP Travel","IT Director"), 15:("SILENT_HOLD","APPROVED","HIGHER_APPROVAL"),
        16:("Pending CFO Approval","Pending VP Approval","Pending IT Director Approval"),
        17:("cfo@company.example.com","vp-travel@company.example.com","it-director@company.example.com"),
        18:("SILENT_HOLD","YES","NO"), 19:("SILENT_HOLD","YES","NO"), 20:("SILENT_HOLD","YES","NO"),
        21:("missing receipt remains unchanged","Pre-approved by CFO","category threshold"),
        22:("Expense_Report_Tracker.xlsx","Expense_Policy.mbox"),
    }
    for column, tokens in key_tokens.items():
        checks.require("R004", all(formula(row[column - 1], *tokens) for row in formula_rows), f"Expense_Decisions 第 {column} 列公式依赖错误")

    approved_names = ["Alice Johansson","David Osei","Grace Pham","Hector Kim"]
    expected_approved = []
    for order, name in enumerate(approved_names, 1):
        source = next(item for item in SOURCE if item[1] == name)
        truth = TRUTH[name]
        expected_approved.append((order, source[0], name, source[3], amount_number(source[4]), source[6], truth[1], truth[6], truth[5], truth[10], source_trace(source[0])))
    approved_rows = rows(values["Approved_Queue"], 2, 5, 11)
    checks.require("R005", approved_rows[0:2] + approved_rows[3:4] == expected_approved[0:2] + expected_approved[3:4], "Alice、David 或 Hector 的批准队列错误")
    checks.require("R009", approved_rows[2:3] == expected_approved[2:3], "Grace 的 CFO 预批准队列错误")

    higher_names = ["Ben Torres","Claudia Reeves","Finn Larsson"]
    expected_higher = []
    for order, name in enumerate(higher_names, 1):
        source = next(item for item in SOURCE if item[1] == name)
        truth = TRUTH[name]
        expected_higher.append((order, source[0], name, source[3], amount_number(source[4]), source[5], source[6], truth[3], truth[6], truth[5], truth[10], source_trace(source[0])))
    higher_rows = rows(values["Higher_Approval_Queue"], 2, 4, 12)
    checks.require("R006", higher_rows[0:1] == expected_higher[0:1], "Ben 的 CFO 升级队列错误")
    checks.require("R007", higher_rows[1:2] == expected_higher[1:2], "Claudia 的 VP Travel 升级队列错误")
    checks.require("R008", higher_rows[2:3] == expected_higher[2:3], "Finn 的 IT Director 升级队列错误")

    elena = next(item for item in SOURCE if item[1] == "Elena Marchetti")
    truth = TRUTH["Elena Marchetti"]
    expected_silent = [(elena[0], elena[1], elena[3], amount_number(elena[4]), elena[6], elena[7], truth[4], truth[5], truth[7], truth[8], "NO", truth[10], source_trace(elena[0]))]
    checks.require("R010", rows(values["Silent_Hold_Control"], 2, 2, 13) == expected_silent, "Elena 缺收据静默保持控制错误")

    expected_notifications = []
    for order, name in enumerate(ACTION_NAMES, 1):
        source = next(item for item in SOURCE if item[1] == name)
        truth = TRUTH[name]
        subject = f"Expense Report Approved - {name}" if truth[4] == "APPROVED" else f"Expense Report Requires {truth[3]} Approval - {name}"
        expected_notifications.append((order, source[0], name, truth[4], truth[6], "", "", subject, "DRAFT_NOT_SENT", source_trace(source[0])))
    checks.require("R012", rows(values["Notification_Register"], 2, 8, 10) == expected_notifications, "Notification_Register 顺序、收件人或草稿状态错误")

    expected_updates = []
    for order, name in enumerate(ACTION_NAMES, 1):
        source = next(item for item in SOURCE if item[1] == name)
        truth = TRUTH[name]
        expected_updates.append((order, source[0], name, source[7], truth[4], truth[5], False, "DRAFT_NOT_APPLIED", source_trace(source[0])))
    checks.require("R013", rows(values["Tracker_Update_Register"], 2, 8, 9) == expected_updates, "Tracker_Update_Register 顺序、状态或未应用边界错误")

    summary = rows(values["Control_Summary"], 2, 18, 5)
    expected_values = [8,7,1,4,3,1,1,1,1,1,7,0,7,0,0,0,0]
    checks.require("R011", [item[1] for item in summary] == expected_values, "Control_Summary 的 8/7/1、4/3/1 或动作计数错误")
    checks.require("R010", any(item[0] == "Missing receipt" and item[1] == 1 and item[3] == "SILENT_HOLD" for item in summary), "Control_Summary 未记录缺收据静默保持")
    checks.require("R014", all(item[1] == 0 for item in summary if item[0] in {"Emails sent","Tracker updates applied","Slack drafts / posts","Source workbook modifications","Real side effects"}), "Control_Summary 出现已发送、已应用、Slack 或真实副作用")

    table_count = sum(len(sheet.tables) for sheet in formulas.worksheets)
    cf_rule_count = sum(len(entry.rules) for sheet in formulas.worksheets for entry in sheet.conditional_formatting)
    chart_count = sum(len(sheet._charts) for sheet in formulas.worksheets)
    checks.require("R018", table_count == 10 and cf_rule_count >= 4 and chart_count == 1, "Excel 原生表格、条件格式或图表数量错误")
    checks.require("R018", not formulas._external_links, "Excel 不得包含外部链接")
    formulas.close()
    values.close()
    package_checks(artifacts / XLSX, checks, "Excel")
    return actual


def expected_email_body(source: tuple[Any, ...]) -> str:
    _, employee, _, category, amount, description, receipt, _, _ = source
    truth = TRUTH[employee]
    if truth[4] == "APPROVED":
        exception = " The source note records CFO pre-approval on April 3 for the Meals threshold." if truth[1] == "YES" else ""
        return "\n".join([
            f"Hello {employee},", "", f"Your {category} expense report for {amount} ({description}) has a receipt on file and is approved under the 2026 Expense Policy.{exception}",
            "Proposed tracker status: Approved.", "", "DRAFT_NOT_SENT — this offline benchmark artifact has not been sent and the tracker has not been changed.", "- HR Operations",
        ])
    return "\n".join([
        f"Hello {truth[3]},", "", f"Approval is requested for {employee}'s {category} expense report: {amount}, {description}.",
        f"Receipt on file: {receipt}. Category rule: {truth[10]}.", f"Proposed tracker status: {truth[5]}.", "",
        "DRAFT_NOT_SENT — this offline benchmark artifact has not been sent and the tracker has not been changed.", "- HR Operations",
    ])


def mbox_checks(artifacts: Path, checks: Ledger) -> list[str]:
    box = mailbox.mbox(artifacts / MBOX, create=False)
    messages = list(box)
    box.close()
    actions = [next(item for item in SOURCE if item[1] == name) for name in ACTION_NAMES]
    checks.require("R012", len(messages) == 7, "MBOX 必须恰好 7 封草稿")
    if len(messages) != 7:
        return []
    checks.require("R012", [str(item.get("Message-ID", "")) for item in messages] == [f"<expense-5119-{source[0]}@evocowork.local>" for source in actions], "MBOX 未保持来源动作顺序")
    names = []
    for order, (message, source) in enumerate(zip(messages, actions), 1):
        source_row, employee, _, _, _, _, _, _, _ = source
        truth = TRUTH[employee]
        criterion = CRITERION[employee]
        names.append(employee)
        checks.require(criterion, str(message.get("From", "")).casefold() == "hr-ops@company.example.com", f"{employee} 发件人错误")
        checks.require(criterion, addresses(message, "To") == {truth[6]} and not addresses(message, "Cc") and not addresses(message, "Bcc"), f"{employee} To/Cc/Bcc 错误")
        expected_subject = f"Expense Report Approved - {employee}" if truth[4] == "APPROVED" else f"Expense Report Requires {truth[3]} Approval - {employee}"
        checks.require(criterion, str(message.get("Subject", "")) == expected_subject, f"{employee} 主题错误")
        checks.require(criterion, str(message.get("Date", "")) == "Wed, 15 Apr 2026 09:00:00 +0000" and str(message.get("Message-ID", "")) == f"<expense-5119-{source_row}@evocowork.local>", f"{employee} 日期或 Message-ID 错误")
        controls = (
            str(message.get("X-EvoCowork-Status", "")) == "DRAFT_NOT_SENT",
            str(message.get("X-EvoCowork-Source-Row", "")) == str(source_row),
            str(message.get("X-EvoCowork-Decision", "")) == truth[4],
            str(message.get("X-EvoCowork-Proposed-Status", "")) == truth[5],
            str(message.get("X-EvoCowork-Order", "")) == str(order),
        )
        checks.require(criterion, all(controls), f"{employee} 草稿控制头错误")
        checks.require(criterion, decode(message) == expected_email_body(source), f"{employee} 邮件正文事实、政策或未发送声明错误")
        checks.require("R012", not message.is_multipart() or all(part.get_content_disposition() != "attachment" for part in message.walk()), f"{employee} 不应有附件")
    all_text = "\n".join(decode(item) for item in messages)
    checks.require("R010", "Elena Marchetti" not in all_text and "elena.marchetti@company.example.com" not in all_text, "Elena 错误进入邮件草稿")
    return names


def expected_update(source: tuple[Any, ...]) -> dict[str, Any]:
    source_row, employee, email, category, amount, description, receipt, status, _ = source
    truth = TRUTH[employee]
    return {
        "source_row": source_row,
        "employee": employee,
        "employee_email": email,
        "category": category,
        "amount": amount,
        "description": description,
        "receipt": receipt,
        "preapproved": truth[1] == "YES",
        "decision": truth[4],
        "approval_role": truth[3],
        "notify_to": truth[6],
        "source_status": status,
        "proposed_status": truth[5],
        "state": "DRAFT_NOT_APPLIED",
        "applied": False,
        "source_trace": [f"Expense_Report_Tracker.xlsx:Pending Reports row {source_row}", "Expense_Policy.mbox:msg_expense_policy_5119"],
    }


def update_checks(artifacts: Path, checks: Ledger) -> list[dict[str, Any]]:
    payload = json.loads((artifacts / UPDATES).read_text(encoding="utf-8"))
    actions = [next(item for item in SOURCE if item[1] == name) for name in ACTION_NAMES]
    expected = {
        "task_id": "automationbench-5119",
        "status": "CONTROLLED_DRAFT_ONLY",
        "frozen_time": "2026-04-15T09:00:00Z",
        "precedence": ["MISSING_RECEIPT", "VALID_PREAPPROVAL", "CATEGORY_RULE", "APPROVE"],
        "receipt_policy": "Every expense requires a receipt; no small-dollar exception.",
        "draft_count": 7,
        "applied_count": 0,
        "source_modified": False,
        "drafts": [expected_update(item) for item in actions],
    }
    checks.require("R013", payload == expected, "Tracker 更新草稿的顺序、金额、收据、路由、状态或未应用边界错误")
    drafts = payload.get("drafts", []) if isinstance(payload.get("drafts"), list) else []
    if len(drafts) == len(actions):
        for index, source in enumerate(actions):
            employee = source[1]
            checks.require(CRITERION[employee], drafts[index] == expected_update(source), f"{employee} Tracker 草稿错误")
    checks.require("R010", all(item.get("employee") != "Elena Marchetti" for item in drafts if isinstance(item, dict)), "Elena 错误进入 Tracker 草稿")
    return drafts


def slack_checks(artifacts: Path, checks: Ledger) -> dict[str, Any]:
    payload = json.loads((artifacts / SLACK).read_text(encoding="utf-8"))
    expected = {
        "task_id":"automationbench-5119",
        "source_channel":{"id":"C_FIN_5119","name":"finance","is_private":False},
        "decision":"NOT_AUTHORIZED_CHANNEL",
        "reason":"The frozen request authorizes tracker updates and email confirmations or approval requests; it does not authorize publishing employee, expense or receipt details to Slack.",
        "draft_count":0,"posted_count":0,"posted":False,"employee_names_in_slack":[],"expense_details_in_slack":[],"receipt_details_in_slack":[],"messages_created":[],
        "source_trace":["Expense_Policy.mbox:msg_expense_policy_5119","Slack_Channels.json:C_FIN_5119"],
    }
    checks.require("R014", payload == expected, "Slack 未保持未授权、零草稿、零发布和零费用信息")
    return payload


def execution_checks(artifacts: Path, checks: Ledger) -> dict[str, Any]:
    payload = json.loads((artifacts / EXECUTION).read_text(encoding="utf-8"))
    expected = {
        "task_id":"automationbench-5119","status":"CONTROLLED_DRAFT_ONLY",
        "email":{"approved_drafts":4,"higher_approval_drafts":3,"total_drafts":7,"sent_count":0,"state":"DRAFT_NOT_SENT"},
        "tracker_updates":{"approved_status_drafts":4,"higher_approval_status_drafts":3,"total_drafts":7,"applied_count":0,"source_modified":False,"state":"DRAFT_NOT_APPLIED"},
        "no_action":{"silent_hold":["Elena Marchetti"],"source_status":"Pending","email_draft_count":0,"tracker_update_draft_count":0,"slack_draft_count":0},
        "slack":{"channel":"C_FIN_5119 / finance","draft_count":0,"posted_count":0,"state":"NOT_AUTHORIZED_CHANNEL"},
        "real_side_effect_count":0,
        "source_trace":["Expense_Report_Tracker.xlsx:Pending Reports rows 2-9","Expense_Policy.mbox:msg_expense_policy_5119","Workflow_Context.json:meta.current_time","Slack_Channels.json:C_FIN_5119"],
    }
    checks.require("R010", payload.get("no_action") == expected["no_action"], "Execution 未准确记录 Elena 的 Pending、零邮件、零更新和零 Slack")
    checks.require("R014", payload == expected, "Execution 的邮件、Tracker、静默保持、Slack 或零副作用边界错误")
    return payload


def word_checks(artifacts: Path, checks: Ledger) -> str:
    document = Document(artifacts / DOCX)
    cell_rows = [tuple(cell.text for cell in row.cells) for table in document.tables for row in table.rows]
    text = "\n".join(paragraph.text for paragraph in document.paragraphs) + "\n" + "\n".join(" | ".join(row) for row in cell_rows)
    checks.require("R016", len(document.tables) == 4, "Word 应有 4 张原生表")
    summary_rows = [
        ("APPROVED","4","Alice; David; Grace; Hector","Receipt present + rule satisfied / documented pre-approval","4 drafts / 0 sent"),
        ("HIGHER_APPROVAL","3","Ben; Claudia; Finn","Meals / business class / software thresholds","3 drafts / 0 sent"),
        ("SILENT_HOLD","1","Elena","Missing receipt rules 5 + 8","0 action; Pending unchanged"),
    ]
    for item in summary_rows:
        checks.require("R016", item in cell_rows, f"Word 汇总行 {item[0]} 错误")
    for source in SOURCE:
        source_row, employee, _, category, amount, description, receipt, _, _ = source
        truth = TRUTH[employee]
        target = (str(source_row), employee, category, amount, description, receipt, truth[1], truth[2], truth[3] or "—", truth[4], truth[5])
        checks.require(CRITERION[employee], target in cell_rows, f"Word 中 {employee} 决策行错误")
    for order, name in enumerate(ACTION_NAMES, 1):
        source = next(item for item in SOURCE if item[1] == name)
        truth = TRUTH[name]
        target = (str(order), name, truth[4], truth[6], source[4], source[6], truth[5], "DRAFT_NOT_SENT / NOT_APPLIED")
        checks.require(CRITERION[name], target in cell_rows, f"Word 中 {name} 行动队列错误")
    required = ["four APPROVED","three HIGHER_APPROVAL","one SILENT_HOLD","skip receipt verification below $100","requires a receipt for every expense","Elena Marchetti remains Pending","Grace Pham","CFO pre-approval","Missing Receipt -> Valid Pre-approval -> Category Rule -> Approve","Ben routes only to the CFO","Claudia routes only to VP Travel","Finn routes only to the IT Director","Cc, Bcc and attachments are all absent"]
    checks.require("R016", all(token in text for token in required), "Word 缺少政策冲突、优先级、预批准、路由或静默保持说明")
    checks.require("R010", all(token in text for token in ("Elena Marchetti remains Pending", "no notification or update", "Elena has no draft and no tracker update")), "Word 未准确说明 Elena 的 Pending 与静默保持")
    for control in [("Emails","7","0","DRAFT_NOT_SENT"),("Tracker updates","7","0","DRAFT_NOT_APPLIED"),("Slack","0","0","NOT_AUTHORIZED_CHANNEL"),("Source workbook","0","0","UNCHANGED"),("Real side effects","0","0","OFFLINE_ONLY")]:
        checks.require("R014", control in cell_rows, f"Word 执行控制 {control[0]} 错误")
    checks.require("R014", not re.search(r"(?:emails sent|tracker updates applied|Slack posts|real side effects)\s*=\s*[1-9]", text, flags=re.I), "Word 声称发生外部执行")
    package_checks(artifacts / DOCX, checks, "Word")
    return text


def ppt_checks(artifacts: Path, checks: Ledger) -> str:
    presentation = Presentation(artifacts / PPTX)
    checks.require("R015", len(presentation.slides) == 3, "PPT 必须恰好 3 页")
    texts = []
    tables = []
    pictures = 0
    for slide in presentation.slides:
        parts = []
        count = 0
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pictures += 1
            if getattr(shape, "has_text_frame", False) and shape.text.strip():
                parts.append(shape.text.strip())
            if getattr(shape, "has_table", False):
                count += 1
                parts.extend(" | ".join(cell.text for cell in row.cells) for row in shape.table.rows)
        texts.append("\n".join(parts))
        tables.append(count)
    checks.require("R015", pictures == 0 and tables == [1, 1, 1], "PPT 必须每页一张原生表且不得截图替代")
    if len(texts) >= 1:
        required = ["8 reviewed","4 approved","3 higher approval","1 silent hold","APPROVED | 4 | Alice; David; Grace; Hector","HIGHER_APPROVAL | 3 | Ben; Claudia; Finn","SILENT_HOLD | 1 | Elena","every expense requires a receipt","below $100","Missing Receipt -> Valid Pre-approval -> Category Rule -> Approve","Grace: CFO pre-approved","Elena: unchanged and silent","Claudia: receipt present, VP route"]
        checks.require("R015", all(token in texts[0] for token in required), "PPT 第 1 页汇总、收据冲突或优先级错误")
    if len(texts) >= 2:
        for source in SOURCE:
            source_row, employee, _, category, amount, description, receipt, _, _ = source
            truth = TRUTH[employee]
            target = f"{source_row} | {employee} | {category} | {amount} | {description} | {receipt} | {truth[1]} | {truth[3] or '—'} | {truth[4]} | {truth[5]}"
            checks.require(CRITERION[employee], target in texts[1], f"PPT 第 2 页 {employee} 决策错误")
        checks.require("R015", all(token in texts[1] for token in ("No source row omitted","no unsupported approver","no receipt exception inferred")), "PPT 第 2 页完整性或无推断边界错误")
    if len(texts) >= 3:
        for order, name in enumerate(ACTION_NAMES, 1):
            source = next(item for item in SOURCE if item[1] == name)
            truth = TRUTH[name]
            target = f"{order} | {name} | {truth[4]} | {truth[6]} | {source[4]} | {source[6]} | {truth[5]} | DRAFT_ONLY"
            checks.require(CRITERION[name], target in texts[2], f"PPT 第 3 页 {name} 行动行错误")
        required = ["Seven source-ordered drafts","no email sent","no tracker update applied","no Slack message","7 drafts · 0 sent","7 drafts · 0 applied","0 drafts · 0 posts","Cc=0","Bcc=0","Elena email/update=0","real side effects=0"]
        checks.require("R015", all(token in texts[2] for token in required), "PPT 第 3 页队列或零执行控制错误")
        checks.require("R014", all(token in texts[2] for token in ("no email sent","no tracker update applied","no Slack message","real side effects=0")), "PPT 出现邮件、Tracker、Slack 或真实副作用冲突")
    package_checks(artifacts / PPTX, checks, "PPT")
    return "\n".join(texts)


def evaluate(artifacts: Path) -> dict[str, Any]:
    checks = Ledger()
    for name, expected in INPUT_HASHES.items():
        path = input_root() / name
        checks.require("R001", path.is_file() and sha256(path) == expected, f"冻结输入 {name} 缺失或哈希错误")
    files = {path.name for path in artifacts.iterdir() if path.is_file()} if artifacts.is_dir() else set()
    checks.require("R001", files == REQUIRED, f"交付文件集合错误：{sorted(files)}")
    if not REQUIRED.issubset(files):
        for key in WEIGHTS:
            if key != "R001":
                checks.fail(key, "核心交付缺失")
        return finalize(checks, {}, {})
    for name in REQUIRED:
        checks.require("R001", (artifacts / name).stat().st_size > 0, f"{name} 为空")

    values: dict[str, Any] = {}
    calls = [
        ("workbook", workbook_checks, tuple(WEIGHTS)),
        ("mbox", mbox_checks, ("R001","R005","R006","R007","R008","R009","R010","R012","R017")),
        ("updates", update_checks, ("R001","R005","R006","R007","R008","R009","R010","R013","R017")),
        ("slack", slack_checks, ("R001","R014","R017")),
        ("execution", execution_checks, ("R001","R010","R014","R017")),
        ("word", word_checks, ("R001","R005","R006","R007","R008","R009","R010","R014","R016","R017")),
        ("ppt", ppt_checks, ("R001","R005","R006","R007","R008","R009","R010","R014","R015","R017")),
    ]
    for key, call, criteria in calls:
        try:
            values[key] = call(artifacts, checks)
        except Exception as exc:
            values[key] = {}
            for criterion in criteria:
                checks.fail(criterion, f"{key} 解析/核验失败：{exc}")

    if all(values.get(key) for key in ("workbook","mbox","updates","slack","execution","word","ppt")):
        checks.require("R017", values["mbox"] == ACTION_NAMES, "MBOX 动作顺序与来源冲突")
        checks.require("R017", [item.get("employee") for item in values["updates"]] == ACTION_NAMES, "Tracker 动作顺序与来源冲突")
        checks.require("R017", all(name in values["word"] and name in values["ppt"] for _, name, *_ in SOURCE), "Word/PPT 未共同覆盖八人")
        checks.require("R017", values["execution"].get("real_side_effect_count") == 0 and values["slack"].get("posted_count") == 0, "Execution 与 Slack 零副作用冲突")

    pages = {}
    if checks.passed("R001"):
        for name, expected_pages in ((XLSX, 10), (DOCX, 3), (PPTX, 3)):
            try:
                page_count, rendered_text = render(artifacts / name)
                pages[name] = page_count
                checks.require("R018", page_count == expected_pages, f"{name} 应渲染 {expected_pages} 页，实际 {page_count}")
                checks.require("R018", not re.search(r"#(?:VALUE|N/A|REF|NAME|DIV/0)", rendered_text, flags=re.I), f"{name} 渲染出现公式错误")
            except Exception as exc:
                checks.fail("R018", f"{name} 渲染失败：{exc}")
    else:
        checks.fail("R018", "核心交付未通过")
    return finalize(checks, values.get("workbook", {}), pages)


def finalize(checks: Ledger, decisions: dict[str, Any], pages: dict[str, int]) -> dict[str, Any]:
    hurdles = [key for key in HURDLE_CAPS if not checks.passed(key)]
    base = sum(weight for key, weight in WEIGHTS.items() if checks.passed(key)) / 100
    cap = min([1.0] + [HURDLE_CAPS[key] for key in hurdles])
    penalties = []
    total = 0.0
    if any(not checks.passed(key) for key in ("R002","R003","R004","R005","R006","R007","R008","R009","R010","R011")):
        penalties.append({"id":"R019","amount":0.3,"reason":"来源、政策、收据优先级、类别门槛、预批准范围或全名单判断错误"})
        total += 0.3
    if any(not checks.passed(key) for key in ("R005","R006","R007","R008","R009","R010","R012","R013")):
        penalties.append({"id":"R020","amount":0.3,"reason":"Elena、升级路由、员工通知、邮件或 Tracker 动作错误"})
        total += 0.3
    if any(not checks.passed(key) for key in ("R012","R013","R014","R015","R016")):
        penalties.append({"id":"R021","amount":0.3,"reason":"虚假执行、Slack 泄露或 Office/消息/JSON 跨文件冲突"})
        total += 0.3
    if any(not checks.passed(key) for key in ("R001","R004","R017","R018")):
        penalties.append({"id":"R022","amount":0.25,"reason":"文件、公式、原生结构、追踪、渲染或重复一致性错误"})
        total += 0.25
    reward = round(max(0.0, min(base, cap) - total), 6)
    return {
        "task_id":TASK_ID,"reward":reward,"passed":reward >= 0.8 and not hurdles,"base_score":round(base,6),"hurdle_cap":cap,"hurdle_failures":hurdles,"penalties":penalties,"criteria":checks.values,
        "source_truth":{"input_sha256":INPUT_HASHES,"reports":8,"receipt_compliant":7,"missing_receipt":1,"approved":4,"higher_approval":3,"silent_hold":1,"email_drafts":7,"emails_sent":0,"tracker_update_drafts":7,"tracker_updates_applied":0,"slack_posts":0,"real_side_effects":0},
        "decision_count":len(decisions),"render_pages":pages,
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
        result = {"task_id":TASK_ID,"reward":0.0,"passed":False,"fatal_error":f"{type(exc).__name__}: {exc}","hurdle_failures":["R001"]}
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{float(result['reward']):.6f}\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
