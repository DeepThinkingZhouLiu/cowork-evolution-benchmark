#!/usr/bin/env python3
"""AutomationBench 5120 背景调查治理题的逐题确定性 Judge。"""

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


TASK_ID = "cowork-evo/automationbench-5120-background-check-control-package"
XLSX = "Background_Check_Audit.xlsx"
DOCX = "Background_Check_Control_Brief.docx"
PPTX = "Background_Check_Review.pptx"
MBOX = "Background_Check_Notification_Drafts.mbox"
UPDATES = "Background_Check_Tracker_Update_Drafts.json"
SLACK = "Background_Check_Slack_Drafts.json"
EXECUTION = "Background_Check_Execution_Control.json"
REQUIRED = {XLSX, DOCX, PPTX, MBOX, UPDATES, SLACK, EXECUTION}
INPUT_HASHES = {
    "Background_Check_Messages.mbox": "d7645da7106062a38838f28b957e3c0d6cabc0e4fb560f9b432e2620736a2c68",
    "Background_Check_Tracker.xlsx": "b29b0e70e736c6dfd945741d006028be6cd46e6e5840a00fcbf4adf73c0a60ba",
    "Slack_Channels.json": "568ea5e7aeaa300bf8e509d4c725ad43d6e799069dc671e615df7038b62cd951",
    "Workflow_Context.json": "76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99",
}
WEIGHTS = {"R001":4,"R002":6,"R003":7,"R004":8,"R005":6,"R006":6,"R007":7,"R008":5,"R009":5,"R010":8,"R011":6,"R012":8,"R013":7,"R014":5,"R015":5,"R016":4,"R017":3}
HURDLE_CAPS = {"R001":0.0,"R002":0.2,"R003":0.2,"R004":0.3,"R005":0.3,"R006":0.2,"R007":0.15,"R008":0.25,"R009":0.3,"R010":0.2,"R011":0.25,"R012":0.15,"R013":0.15,"R014":0.35,"R015":0.35}

SOURCE = [
    (2,"Liam Chen","BGC-4401","Derek Huang","derek.huang@company.example.com","Software Engineer","Clear","Awaiting Action",""),
    (3,"Maya Okafor","BGC-4405","Mira Patel","mira.patel@company.example.com","Marketing Analyst","Failed","Awaiting Action",""),
    (4,"Naomi Yoshida","BGC-4412","Lisa Wang","lisa.wang@company.example.com","Account Executive","Clear","Awaiting Action",""),
    (5,"Oliver Strand","BGC-4418","Derek Huang","derek.huang@company.example.com","DevOps Engineer","Pending","Awaiting Action",""),
    (6,"Petra Johansson","BGC-4422","Mira Patel","mira.patel@company.example.com","Content Writer","Clear","Awaiting Action",""),
    (7,"Quinn Davis","BGC-4430","Lisa Wang","lisa.wang@company.example.com","Sales Rep","Failed","Awaiting Action",""),
    (8,"Rosa Martinez","BGC-4435","Derek Huang","derek.huang@company.example.com","QA Engineer","Clear","Awaiting Action",""),
    (9,"Sanjay Gupta","BGC-4441","Mira Patel","mira.patel@company.example.com","Data Analyst","Pending","Awaiting Action",""),
]
POLICY_CLAUSES = [
    (1,"Clear","Proceed with onboarding; email hiring manager; propose Cleared - Proceed."),
    (2,"Failed","Rescind offer; email HR Director; propose Failed - Offer Rescinded; do not email candidate."),
    (3,"Conditional","Proceed with onboarding but notify Legal before start; propose Conditional - Legal Notified."),
    (4,"Pending","No action; leave Status unchanged."),
]
TRUTH = {
    "Liam Chen":("","Clear","YES","NO","NO","NO","CLEARED_PROCEED","Cleared - Proceed","derek.huang@company.example.com","YES","YES","ONBOARDING_MAY_PROCEED","Manager receives clearance confirmation; Slack receives operational label only.","Legal procedure 1: Clear; email the hiring manager and propose Cleared - Proceed."),
    "Maya Okafor":("","Failed","NO","YES","NO","NO","FAILED_RESCIND","Failed - Offer Rescinded","hr-director@company.example.com","YES","YES","HR_DIRECTOR_ACTION","HR Director only; never email candidate; Slack omits raw result.","Legal procedure 2: rescind offer, notify HR Director, and never email the candidate directly."),
    "Naomi Yoshida":("Conditional","Conditional","NO","NO","YES","NO","CONDITIONAL_LEGAL","Conditional - Legal Notified","legal@company.example.com","YES","YES","LEGAL_NOTIFICATION_BEFORE_START","Legal only; no manager notice; sensitive vendor reason omitted outside restricted evidence.","Vendor correction msg_bgcheck_correction_5120 overrides tracker row 4; Legal procedure 3 applies."),
    "Oliver Strand":("","Pending","NO","NO","NO","YES","PENDING_NO_ACTION","Awaiting Action","","NO","NO","PENDING_NO_ACTION","No email or tracker update; Slack receives operational label only.","Legal procedure 4: Pending requires no action and source status remains unchanged."),
    "Petra Johansson":("","Clear","YES","NO","NO","NO","CLEARED_PROCEED","Cleared - Proceed","mira.patel@company.example.com","YES","YES","ONBOARDING_MAY_PROCEED","Manager receives clearance confirmation; Slack receives operational label only.","Legal procedure 1: Clear; email the hiring manager and propose Cleared - Proceed."),
    "Quinn Davis":("","Failed","NO","YES","NO","NO","FAILED_RESCIND","Failed - Offer Rescinded","hr-director@company.example.com","YES","YES","HR_DIRECTOR_ACTION","HR Director only; never email candidate; Slack omits raw result.","Legal procedure 2: rescind offer, notify HR Director, and never email the candidate directly."),
    "Rosa Martinez":("","Clear","YES","NO","NO","NO","CLEARED_PROCEED","Cleared - Proceed","derek.huang@company.example.com","YES","YES","ONBOARDING_MAY_PROCEED","Manager receives clearance confirmation; Slack receives operational label only.","Legal procedure 1: Clear; email the hiring manager and propose Cleared - Proceed."),
    "Sanjay Gupta":("","Pending","NO","NO","NO","YES","PENDING_NO_ACTION","Awaiting Action","","NO","NO","PENDING_NO_ACTION","No email or tracker update; Slack receives operational label only.","Legal procedure 4: Pending requires no action and source status remains unchanged."),
}
CRITERION = {"Liam Chen":"R005","Petra Johansson":"R005","Rosa Martinez":"R005","Maya Okafor":"R006","Quinn Davis":"R006","Naomi Yoshida":"R007","Oliver Strand":"R008","Sanjay Gupta":"R008"}
ACTION_NAMES = ["Liam Chen","Maya Okafor","Naomi Yoshida","Petra Johansson","Quinn Davis","Rosa Martinez"]
SAFE_LABELS = ["ONBOARDING_MAY_PROCEED","HR_DIRECTOR_ACTION","LEGAL_NOTIFICATION_BEFORE_START","PENDING_NO_ACTION","ONBOARDING_MAY_PROCEED","HR_DIRECTOR_ACTION","ONBOARDING_MAY_PROCEED","PENDING_NO_ACTION"]


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
                payload = part.get_payload(decode=True) or b""
                parts.append(payload.decode(part.get_content_charset() or "utf-8", errors="replace"))
        return "\n".join(parts).strip()
    payload = message.get_payload(decode=True) or b""
    return payload.decode(message.get_content_charset() or "utf-8", errors="replace").strip()


def addresses(message, header: str) -> set[str]:
    return {address.casefold() for _, address in getaddresses(message.get_all(header, [])) if address}


def source_messages() -> dict[str, dict[str, str]]:
    box = mailbox.mbox(input_root() / "Background_Check_Messages.mbox", create=False)
    raw = list(box)
    box.close()
    result = {}
    for message in raw:
        message_id = str(message.get("X-AutomationBench-Message-ID", ""))
        result[message_id] = {"message_id":message_id,"from":str(message.get("From", "")),"to":str(message.get("To", "")),"subject":str(message.get("Subject", "")),"date":str(message.get("Date", "")),"body":decode(message)}
    return result


def source_trace(source_row: int) -> str:
    suffix = " | Background_Check_Messages.mbox:msg_bgcheck_correction_5120" if source_row == 4 else ""
    return f"Background_Check_Tracker.xlsx:Candidates row {source_row}{suffix} | Background_Check_Messages.mbox:msg_bgcheck_policy_5120"


def expected_decisions() -> dict[str, tuple[Any, ...]]:
    result = {}
    for source in SOURCE:
        source_row, candidate, candidate_id, manager, manager_email, position, source_result, source_status, _ = source
        result[candidate] = (source_row,candidate,candidate_id,manager,manager_email,position,source_result,source_status,0,*TRUTH[candidate],source_trace(source_row))
    return result


def render(path: Path) -> tuple[int, str]:
    with tempfile.TemporaryDirectory(prefix="judge-5120-") as temporary:
        work = Path(temporary)
        output = work / "out"
        output.mkdir()
        process = subprocess.run(["libreoffice","--headless",f"-env:UserInstallation=file://{work/'profile'}","--convert-to","pdf","--outdir",str(output),str(path)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=300)
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
    names = ["Source_Candidates","Message_Source","Evidence_Register","Candidate_Decisions","Clear_Queue","Failed_Control","Conditional_Control","Pending_Control","Communication_Register","Control_Summary"]
    checks.require("R001", formulas.sheetnames == names, "Excel 十张工作表名称或顺序错误")
    checks.require("R002", rows(values["Source_Candidates"], 2, 9, 9) == SOURCE, "Source_Candidates 未按源顺序保留八条记录")

    messages = source_messages()
    correction = messages.get("msg_bgcheck_correction_5120", {})
    policy = messages.get("msg_bgcheck_policy_5120", {})
    expected_messages = [
        (correction.get("message_id",""),correction.get("from",""),correction.get("to",""),correction.get("subject",""),correction.get("date",""),correction.get("body",""),"Overrides Naomi tracker result","Restricted evidence; detailed reason excluded from all outbound drafts and open-channel summary"),
        (policy.get("message_id",""),policy.get("from",""),policy.get("to",""),policy.get("subject",""),policy.get("date",""),policy.get("body",""),"Governs four result classes","Operational policy; candidate-result details still minimized by audience"),
    ]
    checks.require("R003", rows(values["Message_Source"], 2, 3, 8) == expected_messages, "两封消息的全文、元数据、权威性或敏感处理错误")
    expected_evidence = []
    for source_row, candidate, candidate_id, manager, manager_email, position, source_result, source_status, notes in SOURCE:
        fact = f"candidate={candidate}; id={candidate_id}; manager={manager}; manager_email={manager_email}; position={position}; result={source_result}; status={source_status}; notes={notes}"
        expected_evidence.append((f"EV-ROW-{source_row}","CANDIDATE_ROW",candidate,fact,"Base record","Restricted HR data",f"Background_Check_Tracker.xlsx:Candidates row {source_row}"))
    expected_evidence.append(("EV-OVERRIDE-1","VENDOR_CORRECTION","Naomi Yoshida / BGC-4412",correction.get("body",""),"Override Clear with Conditional","Detailed reason restricted","Background_Check_Messages.mbox:msg_bgcheck_correction_5120"))
    for number, result, clause in POLICY_CLAUSES:
        expected_evidence.append((f"EV-POLICY-{number}","LEGAL_POLICY",result,clause,"Decision mapping","Audience-specific disclosure","Background_Check_Messages.mbox:msg_bgcheck_policy_5120"))
    expected_evidence.extend([
        ("EV-TIME-1","TIME","Frozen runtime","2026-04-15T09:00:00Z","Artifact timestamp","Non-sensitive","Workflow_Context.json:meta.current_time"),
        ("EV-SLACK-1","CHANNEL","C_REC_5120 / recruiting","is_private=false; messages=[]","One offline minimal-disclosure draft","Open channel: names + safe operational labels only","Slack_Channels.json:C_REC_5120"),
        ("EV-REQUEST-1","USER_REQUEST","Background-check follow-up","Use tracker and Gmail, take appropriate next step, update #recruiting with affected names","Defines requested workflow","Policy and privacy still govern content","AutomationBench 5120 prompt"),
    ])
    checks.require("R003", rows(values["Evidence_Register"], 2, 17, 7) == expected_evidence, "16 条证据的完整事实、用途、隐私或追踪错误")

    actual = {row[1]: row for row in rows(values["Candidate_Decisions"], 2, 9, 24)}
    expected = expected_decisions()
    checks.require("R009", set(actual) == set(expected), "Candidate_Decisions 未覆盖八人")
    for name, target in expected.items():
        checks.require(CRITERION[name], actual.get(name) == target, f"{name} 的覆盖、有效结果、路由、状态、动作或追踪错误")

    formula_rows = rows(formulas["Candidate_Decisions"], 2, 9, 24)
    checks.require("R004", all(all(str(value).startswith("=") for value in row) for row in formula_rows), "Candidate_Decisions 存在静态值而非逐行原生公式")
    key_tokens = {
        1:("Source_Candidates",),2:("Source_Candidates",),3:("Source_Candidates",),4:("Source_Candidates",),5:("Source_Candidates",),6:("Source_Candidates",),7:("Source_Candidates",),8:("Source_Candidates",),9:("Source_Candidates",),
        10:("IF","BGC-4412","Conditional"),11:("IF","J","G"),12:("IF","Clear"),13:("IF","Failed"),14:("IF","Conditional"),15:("IF","Pending"),
        16:("CLEARED_PROCEED","FAILED_RESCIND","CONDITIONAL_LEGAL","PENDING_NO_ACTION"),17:("Cleared - Proceed","Failed - Offer Rescinded","Conditional - Legal Notified"),
        18:("hr-director@company.example.com","legal@company.example.com"),19:("IF","YES","NO"),20:("IF","YES","NO"),
        21:("ONBOARDING_MAY_PROCEED","HR_DIRECTOR_ACTION","LEGAL_NOTIFICATION_BEFORE_START","PENDING_NO_ACTION"),22:("Manager receives","HR Director only","Legal only","No email"),
        23:("Legal procedure 1","Legal procedure 2","msg_bgcheck_correction_5120","Legal procedure 4"),24:("Background_Check_Tracker.xlsx","msg_bgcheck_policy_5120"),
    }
    for column, tokens in key_tokens.items():
        checks.require("R004", all(formula(row[column - 1], *tokens) for row in formula_rows), f"Candidate_Decisions 第 {column} 列公式依赖错误")

    clear_names = ["Liam Chen","Petra Johansson","Rosa Martinez"]
    expected_clear = []
    for order, name in enumerate(clear_names, 1):
        source = next(item for item in SOURCE if item[1] == name)
        truth = TRUTH[name]
        expected_clear.append((order,source[0],name,source[2],source[3],truth[8],truth[6],truth[7],"DRAFT_ONLY",source_trace(source[0])))
    checks.require("R005", rows(values["Clear_Queue"], 2, 4, 10) == expected_clear, "Clear_Queue 的三人、经理或状态错误")

    failed_names = ["Maya Okafor","Quinn Davis"]
    expected_failed = []
    for order, name in enumerate(failed_names, 1):
        source = next(item for item in SOURCE if item[1] == name)
        truth = TRUTH[name]
        expected_failed.append((order,source[0],name,source[2],source[3],"NO",truth[8],"PROHIBITED",truth[6],truth[7],truth[11],"DRAFT_ONLY",source_trace(source[0])))
    checks.require("R006", rows(values["Failed_Control"], 2, 3, 13) == expected_failed, "Failed_Control 未保持 HR Director-only 或候选人直发禁令")

    source = next(item for item in SOURCE if item[1] == "Naomi Yoshida")
    truth = TRUTH["Naomi Yoshida"]
    expected_conditional = [(source[0],source[1],source[2],source[6],truth[0],truth[1],source[3],"NO",truth[8],truth[6],truth[7],"OMITTED",truth[11],"DRAFT_ONLY",source_trace(source[0]))]
    checks.require("R007", rows(values["Conditional_Control"], 2, 2, 15) == expected_conditional, "Naomi 的覆盖、Legal-only、经理禁发或敏感原因控制错误")

    expected_pending = []
    for name in ("Oliver Strand","Sanjay Gupta"):
        source = next(item for item in SOURCE if item[1] == name)
        truth = TRUTH[name]
        expected_pending.append((source[0],name,source[2],source[6],source[7],truth[1],truth[6],truth[7],truth[9],truth[10],truth[11],"NO_ACTION",source_trace(source[0])))
    checks.require("R008", rows(values["Pending_Control"], 2, 3, 13) == expected_pending, "Pending_Control 未保持两人 Awaiting Action 和零动作")

    expected_communications = []
    for order, source in enumerate(SOURCE, 1):
        truth = TRUTH[source[1]]
        expected_communications.append((order,source[0],source[1],truth[1],truth[6],truth[8],"DRAFT_NOT_SENT" if truth[9] == "YES" else "NO_DRAFT","DRAFT_NOT_APPLIED" if truth[10] == "YES" else "NO_DRAFT",truth[11],"DRAFT_NOT_POSTED",truth[12],source_trace(source[0])))
    checks.require("R009", rows(values["Communication_Register"], 2, 9, 12) == expected_communications, "Communication_Register 八人顺序、路由、动作或隐私边界错误")

    expected_summary = [8,3,2,1,2,6,0,6,0,1,0,0,0]
    summary_rows = rows(values["Control_Summary"], 2, 14, 5)
    checks.require("R009", [row[1] for row in summary_rows] == expected_summary, "Control_Summary 的 8、3/2/1/2、6/6/1 或零执行计数错误")
    checks.require("R013", all(row[1] == 0 for row in summary_rows if row[0] in {"Emails sent","Tracker updates applied","Slack posts","Source modifications","Real side effects"}), "Control_Summary 出现真实执行")

    table_count = sum(len(sheet.tables) for sheet in formulas.worksheets)
    cf_rule_count = sum(len(entry.rules) for sheet in formulas.worksheets for entry in sheet.conditional_formatting)
    chart_count = sum(len(sheet._charts) for sheet in formulas.worksheets)
    checks.require("R017", table_count == 10 and cf_rule_count >= 5 and chart_count == 1, "Excel 原生表格、条件格式或图表数量错误")
    if chart_count == 1:
        chart = formulas["Control_Summary"]._charts[0]
        value_ref = chart.series[0].val.numRef.f if chart.series and chart.series[0].val and chart.series[0].val.numRef else ""
        category_ref = chart.series[0].cat.strRef.f if chart.series and chart.series[0].cat and chart.series[0].cat.strRef else ""
        checks.require("R017", "$B$3:$B$6" in value_ref and "$A$3:$A$6" in category_ref, "结果构成图未覆盖 Clear、Failed、Conditional、Pending 四类")
    checks.require("R017", not formulas._external_links, "Excel 不得包含外部链接")
    formulas.close()
    values.close()
    package_checks(artifacts / XLSX, checks, "Excel")
    return actual


def expected_email_body(source: tuple[Any, ...]) -> str:
    source_row, candidate, candidate_id, _, _, position, _, _, _ = source
    truth = TRUTH[candidate]
    common = ["DRAFT_NOT_SENT","",f"Candidate: {candidate}",f"Candidate ID: {candidate_id}",f"Position: {position}",""]
    if truth[6] == "CLEARED_PROCEED":
        detail = ["The background check is cleared. Onboarding may proceed under the Legal handling procedure.","Proposed tracker status: Cleared - Proceed."]
    elif truth[6] == "FAILED_RESCIND":
        detail = ["The background check requires the offer to be rescinded under the Legal handling procedure.","Do not contact the candidate directly about the result.","Proposed tracker status: Failed - Offer Rescinded."]
    else:
        detail = ["A vendor correction sets the effective result to Conditional. Onboarding may proceed, but Legal notification is required before the start date.","The detailed vendor reason is intentionally omitted from this draft.","Proposed tracker status: Conditional - Legal Notified."]
    return "\n".join(common + detail + ["","This is an offline draft. No email has been sent and the source tracker has not been changed.","- HR Operations"])


def expected_subject(name: str) -> str:
    decision = TRUTH[name][6]
    if decision == "CLEARED_PROCEED":
        return f"Background Check Cleared - {name}"
    if decision == "FAILED_RESCIND":
        return f"HR Director Action Required - {name}"
    return f"Legal Notification Before Start - {name}"


def mbox_checks(artifacts: Path, checks: Ledger) -> list[str]:
    box = mailbox.mbox(artifacts / MBOX, create=False)
    messages = list(box)
    box.close()
    actions = [next(item for item in SOURCE if item[1] == name) for name in ACTION_NAMES]
    checks.require("R010", len(messages) == 6, "MBOX 必须恰好六封草稿")
    if len(messages) != 6:
        return []
    checks.require("R010", [str(item.get("Message-ID", "")) for item in messages] == [f"<background-5120-{source[0]}@evocowork.local>" for source in actions], "MBOX 未保持来源动作顺序")
    names = []
    for order, (message, source) in enumerate(zip(messages, actions), 1):
        source_row, candidate, candidate_id, _, _, _, _, _, _ = source
        truth = TRUTH[candidate]
        criterion = CRITERION[candidate]
        names.append(candidate)
        checks.require(criterion, str(message.get("From", "")).casefold() == "hr-ops@company.example.com", f"{candidate} 发件人错误")
        checks.require(criterion, addresses(message, "To") == {truth[8]} and not addresses(message, "Cc") and not addresses(message, "Bcc"), f"{candidate} To/Cc/Bcc 错误")
        checks.require(criterion, str(message.get("Subject", "")) == expected_subject(candidate), f"{candidate} 主题错误")
        controls = (
            str(message.get("Date", "")) == "Wed, 15 Apr 2026 09:00:00 +0000",
            str(message.get("Message-ID", "")) == f"<background-5120-{source_row}@evocowork.local>",
            str(message.get("X-EvoCowork-Status", "")) == "DRAFT_NOT_SENT",
            str(message.get("X-EvoCowork-Source-Row", "")) == str(source_row),
            str(message.get("X-EvoCowork-Candidate-ID", "")) == candidate_id,
            str(message.get("X-EvoCowork-Effective-Result", "")) == truth[1],
            str(message.get("X-EvoCowork-Decision", "")) == truth[6],
            str(message.get("X-EvoCowork-Proposed-Status", "")) == truth[7],
            str(message.get("X-EvoCowork-Order", "")) == str(order),
        )
        checks.require(criterion, all(controls), f"{candidate} 草稿控制头错误")
        checks.require(criterion, decode(message) == expected_email_body(source), f"{candidate} 邮件正文事实、政策、隐私或未发送声明错误")
        checks.require("R010", not message.is_multipart() or all(part.get_content_disposition() != "attachment" for part in message.walk()), f"{candidate} 不应有附件")
    all_text = "\n".join(decode(item) for item in messages)
    checks.require("R008", all(name not in all_text for name in ("Oliver Strand","Sanjay Gupta")), "Pending 候选人错误进入邮件草稿")
    raw = (artifacts / MBOX).read_text(encoding="utf-8", errors="replace").casefold()
    checks.require("R007", "lisa.wang@company.example.com" not in raw and "pending court disposition" not in raw, "Naomi 草稿误发经理或泄露敏感原因")
    checks.require("R006", all(addresses(message, "To") == {"hr-director@company.example.com"} for message, source in zip(messages, actions) if source[1] in {"Maya Okafor","Quinn Davis"}), "Failed 草稿未保持 HR Director-only")
    return names


def expected_update(source: tuple[Any, ...]) -> dict[str, Any]:
    source_row, candidate, candidate_id, _, _, _, source_result, source_status, _ = source
    truth = TRUTH[candidate]
    return {
        "source_row":source_row,"candidate":candidate,"candidate_id":candidate_id,"source_result":source_result,"vendor_override":truth[0],"effective_result":truth[1],"decision":truth[6],
        "source_status":source_status,"proposed_status":truth[7],"notify_to":truth[8],"state":"DRAFT_NOT_APPLIED","applied":False,"source_trace":source_trace(source_row).split(" | "),
    }


def update_checks(artifacts: Path, checks: Ledger) -> list[dict[str, Any]]:
    payload = json.loads((artifacts / UPDATES).read_text(encoding="utf-8"))
    actions = [next(item for item in SOURCE if item[1] == name) for name in ACTION_NAMES]
    expected = {"task_id":"automationbench-5120","status":"CONTROLLED_DRAFT_ONLY","frozen_time":"2026-04-15T09:00:00Z","precedence":["VENDOR_CORRECTION","LEGAL_RESULT_MAPPING","AUDIENCE_AND_PRIVACY_CONTROL"],"draft_count":6,"applied_count":0,"source_modified":False,"drafts":[expected_update(item) for item in actions]}
    checks.require("R011", payload == expected, "Tracker 草稿的顺序、覆盖、状态、路由或未应用边界错误")
    drafts = payload.get("drafts", []) if isinstance(payload.get("drafts"), list) else []
    if len(drafts) == len(actions):
        for index, source in enumerate(actions):
            checks.require(CRITERION[source[1]], drafts[index] == expected_update(source), f"{source[1]} Tracker 草稿错误")
    checks.require("R008", all(item.get("candidate") not in {"Oliver Strand","Sanjay Gupta"} for item in drafts if isinstance(item, dict)), "Pending 候选人错误进入 Tracker 草稿")
    return drafts


def expected_slack() -> dict[str, Any]:
    entries = [{"source_row":source[0],"candidate":source[1],"operational_status":label} for source, label in zip(SOURCE, SAFE_LABELS)]
    text = "#recruiting background-check operations update (offline draft):\n" + "\n".join(f"- {entry['candidate']} — {entry['operational_status']}" for entry in entries)
    return {
        "task_id":"automationbench-5120","source_channel":{"id":"C_REC_5120","name":"recruiting","is_private":False},"privacy_mode":"NAMES_AND_OPERATIONAL_LABELS_ONLY","draft_count":1,"posted_count":0,"posted":False,
        "drafts":[{"draft_id":"background-5120-recruiting-1","status":"DRAFT_NOT_POSTED","channel_id":"C_REC_5120","channel_name":"recruiting","entries":entries,"text":text}],
        "excluded_fields":["candidate_id","email","position","raw_result","investigation_reason"],
        "source_trace":["Background_Check_Tracker.xlsx:Candidates rows 2-9","Background_Check_Messages.mbox:msg_bgcheck_policy_5120","Slack_Channels.json:C_REC_5120"],
    }


def slack_checks(artifacts: Path, checks: Ledger) -> dict[str, Any]:
    payload = json.loads((artifacts / SLACK).read_text(encoding="utf-8"))
    expected = expected_slack()
    checks.require("R012", payload == expected, "Slack 草稿的频道、八人顺序、标签或未发布边界错误")
    raw = (artifacts / SLACK).read_text(encoding="utf-8").casefold()
    banned = [item[2].casefold() for item in SOURCE] + [item[4].casefold() for item in SOURCE] + [item[5].casefold() for item in SOURCE] + ["pending court disposition","\"failed\"","\"conditional\""]
    checks.require("R012", not any(token in raw for token in banned), "Slack 泄露候选人 ID、邮箱、职位、原始结果词或敏感原因")
    return payload


def expected_execution() -> dict[str, Any]:
    return {
        "task_id":"automationbench-5120","status":"CONTROLLED_DRAFT_ONLY",
        "email":{"clear_drafts":3,"failed_drafts":2,"conditional_drafts":1,"total_drafts":6,"sent_count":0,"state":"DRAFT_NOT_SENT"},
        "tracker_updates":{"clear_status_drafts":3,"failed_status_drafts":2,"conditional_status_drafts":1,"total_drafts":6,"applied_count":0,"source_modified":False,"state":"DRAFT_NOT_APPLIED"},
        "pending_no_action":{"candidates":["Oliver Strand","Sanjay Gupta"],"source_status":"Awaiting Action","email_draft_count":0,"tracker_update_draft_count":0},
        "slack":{"channel":"C_REC_5120 / recruiting","draft_count":1,"posted_count":0,"state":"DRAFT_NOT_POSTED","privacy_mode":"NAMES_AND_OPERATIONAL_LABELS_ONLY"},
        "source_tracker_modification_count":0,"real_side_effect_count":0,
        "source_trace":["Background_Check_Tracker.xlsx:Candidates rows 2-9","Background_Check_Messages.mbox:msg_bgcheck_correction_5120","Background_Check_Messages.mbox:msg_bgcheck_policy_5120","Workflow_Context.json:meta.current_time","Slack_Channels.json:C_REC_5120"],
    }


def execution_checks(artifacts: Path, checks: Ledger) -> dict[str, Any]:
    payload = json.loads((artifacts / EXECUTION).read_text(encoding="utf-8"))
    expected = expected_execution()
    checks.require("R008", payload.get("pending_no_action") == expected["pending_no_action"], "Execution 未准确记录两名 Pending 候选人的零动作")
    checks.require("R013", payload == expected, "Execution 的邮件、Tracker、Slack、源修改或零副作用边界错误")
    return payload


def word_checks(artifacts: Path, checks: Ledger) -> str:
    document = Document(artifacts / DOCX)
    cell_rows = [tuple(cell.text for cell in row.cells) for table in document.tables for row in table.rows]
    text = "\n".join(paragraph.text for paragraph in document.paragraphs) + "\n" + "\n".join(" | ".join(row) for row in cell_rows)
    checks.require("R015", len(document.tables) == 4, "Word 应有四张原生表")
    summary_rows = [
        ("Clear","3","Liam Chen; Petra Johansson; Rosa Martinez","respective hiring manager","Cleared - Proceed","3 drafts / 0 sent"),
        ("Failed","2","Maya Okafor; Quinn Davis","HR Director only","Failed - Offer Rescinded","2 drafts / 0 sent; candidate direct email prohibited"),
        ("Conditional","1","Naomi Yoshida","Legal only; no manager","Conditional - Legal Notified","1 draft / 0 sent; reason omitted"),
        ("Pending","2","Oliver Strand; Sanjay Gupta","no email / no update","Awaiting Action","0 action"),
    ]
    for item in summary_rows:
        checks.require("R015", item in cell_rows, f"Word 汇总行 {item[0]} 错误")
    for source in SOURCE:
        truth = TRUTH[source[1]]
        target = (str(source[0]),source[1],source[2],source[3],source[5],source[6],truth[0] or "—",truth[1],truth[6],truth[7],truth[8] or "—")
        checks.require(CRITERION[source[1]], target in cell_rows, f"Word 中 {source[1]} 决策行错误")
    for order, name in enumerate(ACTION_NAMES, 1):
        truth = TRUTH[name]
        target = (str(order),name,truth[1],truth[6],truth[8],truth[7],"DRAFT_NOT_SENT / NOT_APPLIED")
        checks.require(CRITERION[name], target in cell_rows, f"Word 中 {name} 行动队列错误")
    required = ["vendor's 2026-04-14 correction overrides Naomi Yoshida","Lisa Wang is not notified","all eight names in source order","Candidate IDs, emails, positions","raw Failed/Conditional result words","Oliver Strand and Sanjay Gupta remain Awaiting Action","generate neither email nor tracker-update drafts","no candidate-direct message exists"]
    checks.require("R015", all(token in text for token in required), "Word 缺少覆盖、路由、Pending、Slack 最小披露或禁发说明")
    checks.require("R007", all(token in text for token in ("vendor's 2026-04-14 correction overrides Naomi Yoshida","Lisa Wang is not notified","detailed vendor reason remains restricted")), "Word 未准确说明 Naomi 覆盖、Legal-only 或敏感原因限制")
    checks.require("R008", all(token in text for token in ("Oliver Strand and Sanjay Gupta remain Awaiting Action","generate neither email nor tracker-update drafts")), "Word 未准确说明两名 Pending 候选人的零动作")
    controls = [("Emails","6","0","DRAFT_NOT_SENT"),("Tracker updates","6","0","DRAFT_NOT_APPLIED"),("Recruiting Slack","1","0","DRAFT_NOT_POSTED"),("Source workbook","0","0","UNCHANGED"),("Real side effects","0","0","OFFLINE_ONLY")]
    for control in controls:
        checks.require("R013", control in cell_rows, f"Word 执行控制 {control[0]} 错误")
    checks.require("R007", "pending court disposition" not in text.casefold(), "Word 泄露 Naomi 的敏感供应商原因")
    package_checks(artifacts / DOCX, checks, "Word")
    return text


def ppt_checks(artifacts: Path, checks: Ledger) -> str:
    presentation = Presentation(artifacts / PPTX)
    checks.require("R014", len(presentation.slides) == 3, "PPT 必须恰好三页")
    texts, tables = [], []
    pictures = 0
    for slide in presentation.slides:
        parts, count = [], 0
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
    checks.require("R014", pictures == 0 and tables == [1,1,1], "PPT 必须每页一张原生表且不得截图替代")
    if len(texts) >= 1:
        required = ["3 | Liam; Petra; Rosa","2 | Maya; Quinn","1 | Naomi","2 | Oliver; Sanjay","tracker Clear -> vendor correction Conditional -> Legal only","Lisa Wang is not notified","Email sent=0","real side effects=0"]
        checks.require("R014", all(token in texts[0] for token in required), "PPT 第1页分流、Naomi 覆盖、隐私或零执行错误")
        checks.require("R007", all(token in texts[0] for token in ("tracker Clear -> vendor correction Conditional -> Legal only","Lisa Wang is not notified","Sensitive vendor reason stays restricted")), "PPT 第1页 Naomi 覆盖、Legal-only 或敏感原因限制错误")
    if len(texts) >= 2:
        for source in SOURCE:
            truth = TRUTH[source[1]]
            target = f"{source[0]} | {source[1]} | {source[2]} | {source[3]} | {source[5]} | {source[6]} | {truth[0] or '—'} | {truth[1]} | {truth[6]} | {truth[8] or '—'}"
            checks.require(CRITERION[source[1]], target in texts[1], f"PPT 第2页 {source[1]} 决策错误")
        checks.require("R014", all(token in texts[1] for token in ("No source row omitted","Pending rows unchanged","Failed candidates never contacted directly")), "PPT 第2页完整性或禁发边界错误")
    if len(texts) >= 3:
        for order, name in enumerate(ACTION_NAMES, 1):
            truth = TRUTH[name]
            target = f"{order} | {name} | {truth[1]} | {truth[6]} | {truth[8]} | {truth[7]} | DRAFT_ONLY"
            checks.require(CRITERION[name], target in texts[2], f"PPT 第3页 {name} 行动行错误")
        required = ["6 drafts · 0 sent","6 drafts · 0 applied","1 draft · 0 posted","all eight names in source order","no IDs, emails, positions, raw result words","Oliver and Sanjay: email=0"]
        checks.require("R014", all(token in texts[2] for token in required), "PPT 第3页队列、Slack 最小披露或零执行错误")
        checks.require("R013", all(token in texts[2] for token in ("6 drafts · 0 sent","6 drafts · 0 applied","1 draft · 0 posted")), "PPT 第3页邮件、Tracker 或 Slack 零执行计数错误")
        checks.require("R008", "Oliver and Sanjay: email=0 · tracker update=0 · source status remains Awaiting Action" in texts[2], "PPT 第3页两名 Pending 候选人的零动作错误")
    joined = "\n".join(texts)
    checks.require("R007", "pending court disposition" not in joined.casefold(), "PPT 泄露 Naomi 的敏感供应商原因")
    package_checks(artifacts / PPTX, checks, "PPT")
    return joined


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
        ("workbook",workbook_checks,tuple(WEIGHTS)),
        ("mbox",mbox_checks,("R001","R005","R006","R007","R008","R010","R016")),
        ("updates",update_checks,("R001","R005","R006","R007","R008","R011","R016")),
        ("slack",slack_checks,("R001","R007","R008","R012","R016")),
        ("execution",execution_checks,("R001","R008","R013","R016")),
        ("word",word_checks,("R001","R005","R006","R007","R008","R013","R015","R016")),
        ("ppt",ppt_checks,("R001","R005","R006","R007","R008","R013","R014","R016")),
    ]
    for key, call, criteria in calls:
        try:
            values[key] = call(artifacts, checks)
        except Exception as exc:
            values[key] = {}
            for criterion in criteria:
                checks.fail(criterion, f"{key} 解析/核验失败：{exc}")

    if all(values.get(key) for key in ("workbook","mbox","updates","slack","execution","word","ppt")):
        checks.require("R016", values["mbox"] == ACTION_NAMES, "MBOX 动作顺序与来源冲突")
        checks.require("R016", [item.get("candidate") for item in values["updates"]] == ACTION_NAMES, "Tracker 动作顺序与来源冲突")
        checks.require("R016", all(name in values["word"] and name in values["ppt"] for _, name, *_ in SOURCE), "Word/PPT 未共同覆盖八人")
        checks.require("R016", values["slack"] == expected_slack() and values["execution"] == expected_execution(), "Slack 与 Execution 控制冲突")

    pages = {}
    if checks.passed("R001"):
        for name, expected_pages in ((XLSX,10),(DOCX,3),(PPTX,3)):
            try:
                page_count, rendered_text = render(artifacts / name)
                pages[name] = page_count
                checks.require("R017", page_count == expected_pages, f"{name} 应渲染 {expected_pages} 页，实际 {page_count}")
                checks.require("R017", not re.search(r"#(?:VALUE|N/A|REF|NAME|DIV/0)", rendered_text, flags=re.I), f"{name} 渲染出现公式错误")
            except Exception as exc:
                checks.fail("R017", f"{name} 渲染失败：{exc}")
    else:
        checks.fail("R017", "核心交付未通过")
    return finalize(checks, values.get("workbook", {}), pages)


def finalize(checks: Ledger, decisions: dict[str, Any], pages: dict[str, int]) -> dict[str, Any]:
    hurdles = [key for key in HURDLE_CAPS if not checks.passed(key)]
    base = sum(weight for key, weight in WEIGHTS.items() if checks.passed(key)) / 100
    cap = min([1.0] + [HURDLE_CAPS[key] for key in hurdles])
    penalties, total = [], 0.0
    if any(not checks.passed(key) for key in ("R002","R003","R004","R005","R006","R007","R008","R009")):
        penalties.append({"id":"R018","amount":0.3,"reason":"来源、政策、供应商覆盖、四类映射、全名单或动作计数错误"})
        total += 0.3
    if any(not checks.passed(key) for key in ("R006","R007","R008","R010","R011","R012")):
        penalties.append({"id":"R019","amount":0.3,"reason":"候选人直发、Naomi 误路由、敏感原因传播、Pending 误动作或 Slack 泄露"})
        total += 0.3
    if any(not checks.passed(key) for key in ("R010","R011","R012","R013","R014","R015","R016")):
        penalties.append({"id":"R020","amount":0.3,"reason":"虚假执行或 Office、消息、JSON 跨文件冲突"})
        total += 0.3
    if any(not checks.passed(key) for key in ("R001","R004","R016","R017")):
        penalties.append({"id":"R021","amount":0.25,"reason":"文件、公式、原生结构、追踪、渲染或重复一致性错误"})
        total += 0.25
    reward = round(max(0.0, min(base, cap) - total), 6)
    return {
        "task_id":TASK_ID,"reward":reward,"passed":reward >= 0.8 and not hurdles,"base_score":round(base,6),"hurdle_cap":cap,"hurdle_failures":hurdles,"penalties":penalties,"criteria":checks.values,
        "source_truth":{"input_sha256":INPUT_HASHES,"candidates":8,"effective_clear":3,"effective_failed":2,"effective_conditional":1,"effective_pending":2,"email_drafts":6,"emails_sent":0,"tracker_update_drafts":6,"tracker_updates_applied":0,"slack_drafts":1,"slack_posts":0,"real_side_effects":0},
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
