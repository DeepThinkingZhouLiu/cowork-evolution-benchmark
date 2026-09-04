#!/usr/bin/env python3
"""AutomationBench 5121 合同工续签控制题的逐题确定性 Judge。

Judge 独立读取四个冻结输入并重建八人真值，不导入参考解法，
也不依赖 Golden 文件哈希。
"""

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


TASK_ID = "cowork-evo/automationbench-5121-contractor-renewal-control-package"
XLSX = "Contractor_Renewal_Audit.xlsx"
DOCX = "Contractor_Renewal_Control_Brief.docx"
PPTX = "Contractor_Renewal_Review.pptx"
MBOX = "Contractor_Renewal_Notification_Drafts.mbox"
UPDATES = "Contractor_Tracker_Update_Drafts.json"
SLACK = "Contractor_Slack_Control.json"
EXECUTION = "Contractor_Execution_Control.json"
REQUIRED = {XLSX, DOCX, PPTX, MBOX, UPDATES, SLACK, EXECUTION}
INPUT_HASHES = {
    "Contractor_Agreement_Tracker.xlsx": "d00f99920947075bc262356257438806def1d0e33d66bb262e8a734d03d43f6e",
    "Contractor_Renewal_Policy.mbox": "f6ba64dcd76ba51014a0c228ee7495408267f600d91e72fcea14f41a542bd78a",
    "Slack_Channels.json": "839b6681f619dc42ab39890f52c697c97bbbcab9dd55bb6319662d05ab76e636",
    "Workflow_Context.json": "76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99",
}
WEIGHTS = {"R001":4,"R002":6,"R003":7,"R004":9,"R005":6,"R006":6,"R007":5,"R008":5,"R009":9,"R010":4,"R011":8,"R012":6,"R013":7,"R014":5,"R015":5,"R016":4,"R017":4}
HURDLE_CAPS = {"R001":0.0,"R002":0.2,"R003":0.2,"R004":0.3,"R005":0.25,"R006":0.2,"R007":0.25,"R008":0.25,"R009":0.15,"R010":0.3,"R011":0.15,"R012":0.2,"R013":0.15,"R014":0.3,"R015":0.3}

SOURCE = [
    (2,"Ana Petrova","ana.petrova@freelance.example.com","Derek Huang","derek.huang@company.example.com","Engineering","vp-eng@company.example.com","2026-05-15","Active",""),
    (3,"Brian Osei","brian.osei@freelance.example.com","Mira Patel","mira.patel@company.example.com","Marketing","vp-mkt@company.example.com","2026-06-30","Active",""),
    (4,"Carmen Ruiz","carmen.ruiz@freelance.example.com","Lisa Wang","lisa.wang@company.example.com","Sales","vp-sales@company.example.com","2026-04-30","Active","On PIP since March 2026"),
    (5,"Darius Kim","darius.kim@freelance.example.com","Derek Huang","derek.huang@company.example.com","Engineering","vp-eng@company.example.com","2026-05-01","Active",""),
    (6,"Eva Lindström","eva.lindstrom@freelance.example.com","Mira Patel","mira.patel@company.example.com","Marketing","vp-mkt@company.example.com","2026-08-15","Active",""),
    (7,"Felix Braun","felix.braun@freelance.example.com","Lisa Wang","lisa.wang@company.example.com","Sales","vp-sales@company.example.com","2026-04-10","Active",""),
    (8,"Gloria Chen","gloria.chen@freelance.example.com","Derek Huang","derek.huang@company.example.com","Engineering","vp-eng@company.example.com","2026-06-10","Active","On PIP since February 2026"),
    (9,"Hugo Santos","hugo.santos@freelance.example.com","Mira Patel","mira.patel@company.example.com","Marketing","vp-mkt@company.example.com","2026-05-20","Active",""),
]
POLICY_RULES = [
    ("R1","RENEWAL_WINDOW","Only process renewals for agreements expiring within the next 60 days."),
    ("R2","ACTIVE_PIP","Contractors with an active Performance Improvement Plan (PIP) cannot be renewed directly. Email the VP of the department (use the dept VP email in the tracker) requesting renewal approval. Set Status to 'Pending VP Approval'."),
    ("R3","STANDARD_RENEWAL","Standard renewals: Set Status to 'Renewed'. Email the contractor and hiring manager confirming the renewal."),
    ("R4","BEYOND_60_DAYS","Agreements expiring beyond 60 days: No action needed. Leave as is."),
    ("R5","ALREADY_EXPIRED","Already expired agreements should be set to 'Expired - Not Renewed'."),
    ("R6","RATE_CHANGES","RATE CHANGES: HR Ops may NOT change contractor rates as part of a renewal. All rate adjustments — including standard cost-of-living increases — must go through Procurement (procurement@company.example.com) and Finance approval. Renewal notices should confirm continuation at existing terms only."),
]
TRUTH = {
    "Ana Petrova":(30,"NO","YES","NO","YES","RENEWED_STANDARD","Renewed","ana.petrova@freelance.example.com","derek.huang@company.example.com","YES","YES","NO","EXISTING_TERMS_ONLY; 5% REQUEST_ESCALATED_TO_PROCUREMENT; NO_RATE_APPLIED","YES","NO","Policy rules 1 and 3: within 60 days, no active PIP"),
    "Brian Osei":(76,"NO","NO","NO","NO","FUTURE_NO_ACTION","Active","","","NO","NO","NO","NO_RATE_ACTION_OUTSIDE_60_DAYS","NO","NO","Policy rule 4: beyond 60 days"),
    "Carmen Ruiz":(15,"NO","YES","YES","NO","PENDING_VP_APPROVAL","Pending VP Approval","vp-sales@company.example.com","","YES","YES","NO","NO_RATE_ACTION_PENDING_VP_APPROVAL","NO","NO","Policy rule 2: active PIP requires department VP approval"),
    "Darius Kim":(16,"NO","YES","NO","YES","RENEWED_STANDARD","Renewed","darius.kim@freelance.example.com","derek.huang@company.example.com","YES","YES","NO","EXISTING_TERMS_ONLY; 5% REQUEST_ESCALATED_TO_PROCUREMENT; NO_RATE_APPLIED","YES","NO","Policy rules 1 and 3: within 60 days, no active PIP"),
    "Eva Lindström":(122,"NO","NO","NO","NO","FUTURE_NO_ACTION","Active","","","NO","NO","NO","NO_RATE_ACTION_OUTSIDE_60_DAYS","NO","NO","Policy rule 4: beyond 60 days"),
    "Felix Braun":(-5,"YES","NO","NO","NO","EXPIRED_NOT_RENEWED","Expired - Not Renewed","","","NO","YES","NO","NO_RATE_ACTION_EXPIRED","NO","NO","Policy rule 5: already expired"),
    "Gloria Chen":(56,"NO","YES","YES","NO","PENDING_VP_APPROVAL","Pending VP Approval","vp-eng@company.example.com","","YES","YES","NO","NO_RATE_ACTION_PENDING_VP_APPROVAL","NO","NO","Policy rule 2: active PIP requires department VP approval"),
    "Hugo Santos":(35,"NO","YES","NO","YES","RENEWED_STANDARD","Renewed","hugo.santos@freelance.example.com","mira.patel@company.example.com","YES","YES","NO","EXISTING_TERMS_ONLY; 5% REQUEST_ESCALATED_TO_PROCUREMENT; NO_RATE_APPLIED","YES","NO","Policy rules 1 and 3: within 60 days, no active PIP"),
}
CRITERION = {"Ana Petrova":"R005","Darius Kim":"R005","Hugo Santos":"R005","Carmen Ruiz":"R006","Gloria Chen":"R006","Felix Braun":"R007","Brian Osei":"R008","Eva Lindström":"R008"}
ACTION_NAMES = ["Ana Petrova","Carmen Ruiz","Darius Kim","Gloria Chen","Hugo Santos"]
TRACKER_NAMES = ["Ana Petrova","Carmen Ruiz","Darius Kim","Felix Braun","Gloria Chen","Hugo Santos"]


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
    digest=hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda:handle.read(1024*1024),b""): digest.update(chunk)
    return digest.hexdigest()


def norm(value: Any) -> Any:
    if value is None: return ""
    if isinstance(value,datetime): return value.date().isoformat()
    if isinstance(value,date): return value.isoformat()
    if isinstance(value,float) and value.is_integer(): return int(value)
    return value


def rows(sheet, start: int, end: int, columns: int) -> list[tuple[Any,...]]:
    return [tuple(norm(sheet.cell(row,column).value) for column in range(1,columns+1)) for row in range(start,end+1)]


def formula(value: Any, *tokens: str) -> bool:
    text=str(value or "")
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


def source_policy() -> dict[str,str]:
    box=mailbox.mbox(input_root()/"Contractor_Renewal_Policy.mbox",create=False); messages=list(box); box.close()
    if len(messages)!=1: raise ValueError("源续签政策邮件数量错误")
    message=messages[0]
    return {"message_id":str(message.get("X-AutomationBench-Message-ID","")),"from":str(message.get("From","")),"to":str(message.get("To","")),"subject":str(message.get("Subject","")),"date":str(message.get("Date","")),"body":decode(message)}


def source_trace(source_row: int) -> str:
    return f"Contractor_Agreement_Tracker.xlsx:Active Agreements row {source_row} | Contractor_Renewal_Policy.mbox:msg_contractor_policy_5121 | Workflow_Context.json:meta.current_time"


def source_by_name(name: str) -> tuple[Any,...]:
    return next(item for item in SOURCE if item[1]==name)


def expected_decisions() -> dict[str,tuple[Any,...]]:
    result={}
    for source in SOURCE:
        source_row,name,email,manager,manager_email,department,vp,end,status,notes=source
        result[name]=(source_row,name,email,manager,manager_email,department,vp,end,status,notes,"2026-04-15",*TRUTH[name],source_trace(source_row))
    return result


def render(path: Path) -> tuple[int,str]:
    with tempfile.TemporaryDirectory(prefix="judge-5121-") as temporary:
        work=Path(temporary); output=work/"out"; output.mkdir()
        process=subprocess.run(["libreoffice","--headless",f"-env:UserInstallation=file://{work/'profile'}","--convert-to","pdf","--outdir",str(output),str(path)],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,timeout=300)
        pdf=output/f"{path.stem}.pdf"
        if process.returncode!=0 or not pdf.is_file(): raise RuntimeError(process.stdout[-1000:])
        info=subprocess.run(["pdfinfo",str(pdf)],capture_output=True,text=True,timeout=60); match=re.search(r"^Pages:\s+(\d+)",info.stdout,flags=re.M)
        text=subprocess.run(["pdftotext",str(pdf),"-"],capture_output=True,text=True,timeout=60)
        if info.returncode!=0 or not match or text.returncode!=0: raise RuntimeError("无法读取渲染 PDF")
        return int(match.group(1)),text.stdout


def package_checks(path: Path, checks: Ledger, prefix: str) -> None:
    try:
        with zipfile.ZipFile(path) as archive:
            names=archive.namelist()
            checks.require("R001",not any(name.endswith("vbaProject.bin") for name in names),f"{prefix} 包含宏")
            checks.require("R001",not any("/embeddings/" in name for name in names),f"{prefix} 包含嵌入对象")
            external=any(name.endswith(".rels") and b'TargetMode="External"' in archive.read(name) for name in names)
            checks.require("R001",not external,f"{prefix} 包含外部链接")
    except Exception as exc:
        checks.fail("R001",f"{prefix} OOXML 包检查失败：{exc}")


def workbook_checks(artifacts: Path, checks: Ledger) -> dict[str,tuple[Any,...]]:
    formulas=load_workbook(artifacts/XLSX,data_only=False); values=load_workbook(artifacts/XLSX,data_only=True)
    names=["Source_Contractors","Policy_Source","Evidence_Register","Renewal_Decisions","Standard_Renewal_Queue","PIP_Approval_Queue","Expired_Control","No_Action_Control","Communication_Register","Control_Summary"]
    checks.require("R001",formulas.sheetnames==names,"Excel 十张工作表名称或顺序错误")
    checks.require("R002",rows(values["Source_Contractors"],2,9,10)==SOURCE,"Source_Contractors 未按源顺序保留八条完整记录")

    policy=source_policy()
    expected_policy=[("POLICY-MESSAGE","FULL_MESSAGE",policy["body"],f"{policy['from']} | {policy['subject']} | {policy['date']}","Contractor_Renewal_Policy.mbox:msg_contractor_policy_5121")]
    expected_policy.extend((rule_id,category,text,"Procurement policy","Contractor_Renewal_Policy.mbox:msg_contractor_policy_5121") for rule_id,category,text in POLICY_RULES)
    checks.require("R003",rows(values["Policy_Source"],2,8,5)==expected_policy,"Policy_Source 邮件全文、六条规则或来源错误")

    expected_evidence=[]
    for source_row,name,email,manager,manager_email,department,vp,end,status,notes in SOURCE:
        fact=f"contractor={name}; contractor_email={email}; manager={manager}; manager_email={manager_email}; department={department}; dept_vp_email={vp}; agreement_end={end}; status={status}; notes={notes}"
        expected_evidence.append((f"EV-ROW-{source_row}","CONTRACTOR_RECORD",name,fact,f"Contractor_Agreement_Tracker.xlsx:Active Agreements row {source_row}"))
    expected_evidence.extend((f"EV-{rule_id}","POLICY",category,text,"Contractor_Renewal_Policy.mbox:msg_contractor_policy_5121") for rule_id,category,text in POLICY_RULES)
    expected_evidence.extend([
        ("EV-POLICY-MESSAGE","MESSAGE",policy["subject"],f"from={policy['from']}; to={policy['to']}; date={policy['date']}; body={policy['body']}","Contractor_Renewal_Policy.mbox:msg_contractor_policy_5121"),
        ("EV-TIME-1","TIME","Frozen runtime","2026-04-15T09:00:00Z","Workflow_Context.json:meta.current_time"),
        ("EV-SLACK-1","CHANNEL","C_HR_5121 / hr-general",'{"id":"C_HR_5121","name":"hr-general","is_private":false}',"Slack_Channels.json:C_HR_5121"),
    ])
    checks.require("R003",rows(values["Evidence_Register"],2,18,5)==expected_evidence,"17 条证据的 ID、类型、完整事实或来源追踪错误")

    expected=expected_decisions(); actual={row[1]:row for row in rows(values["Renewal_Decisions"],2,9,28)}
    checks.require("R010",set(actual)==set(expected),"Renewal_Decisions 未覆盖完整八人")
    for name,target in expected.items(): checks.require(CRITERION[name],actual.get(name)==target,f"{name} 的日期、窗口、PIP、路由、费率或追踪错误")
    formula_rows=rows(formulas["Renewal_Decisions"],2,9,28)
    checks.require("R004",all(all(str(value).startswith("=") for value in row) for row in formula_rows),"Renewal_Decisions 存在静态值而非逐行原生公式")
    key_tokens={8:("Source_Contractors",),10:("IF","Source_Contractors"),11:("DATE(2026,4,15)",),12:("H","K"),13:("L","<0"),14:("L",">=0","<=60"),15:("SEARCH","On PIP"),16:("N","O"),17:("EXPIRED_NOT_RENEWED","FUTURE_NO_ACTION","PENDING_VP_APPROVAL","RENEWED_STANDARD"),18:("Expired - Not Renewed","Pending VP Approval","Renewed"),19:("RENEWED_STANDARD","PENDING_VP_APPROVAL","C","G"),20:("RENEWED_STANDARD","E"),21:("RENEWED_STANDARD","PENDING_VP_APPROVAL"),22:("EXPIRED_NOT_RENEWED","PENDING_VP_APPROVAL","RENEWED_STANDARD"),23:("NO",),24:("EXISTING_TERMS_ONLY","NO_RATE_ACTION_PENDING_VP_APPROVAL","NO_RATE_ACTION_EXPIRED","NO_RATE_ACTION_OUTSIDE_60_DAYS"),25:("RENEWED_STANDARD","YES"),26:("NO",),27:("already expired","beyond 60 days","active PIP","within 60 days"),28:("Contractor_Agreement_Tracker.xlsx","Contractor_Renewal_Policy.mbox","Workflow_Context.json")}
    for column,tokens in key_tokens.items(): checks.require("R004",all(formula(row[column-1],*tokens) for row in formula_rows),f"Renewal_Decisions 第 {column} 列公式依赖错误")

    expected_standard=[]
    for order,name in enumerate(("Ana Petrova","Darius Kim","Hugo Santos"),1):
        source=source_by_name(name); truth=TRUTH[name]
        expected_standard.append((order,source[0],name,source[2],source[3],source[4],source[7],truth[0],truth[5],truth[6],truth[7],truth[8],truth[12],truth[13],"DRAFT_NOT_SENT / NOT_APPLIED",False,False,source_trace(source[0])))
    checks.require("R005",rows(values["Standard_Renewal_Queue"],2,4,18)==expected_standard,"Standard_Renewal_Queue 三人、收件人、现有条款或状态错误")

    expected_pip=[]
    for order,name in enumerate(("Carmen Ruiz","Gloria Chen"),1):
        source=source_by_name(name); truth=TRUTH[name]
        expected_pip.append((order,source[0],name,source[5],source[6],source[7],truth[0],source[9],truth[5],truth[6],truth[7],truth[8],truth[12],"DRAFT_NOT_SENT / NOT_APPLIED",False,False,source_trace(source[0])))
    checks.require("R006",rows(values["PIP_Approval_Queue"],2,3,17)==expected_pip,"PIP_Approval_Queue 两人、VP-only 路由或费率边界错误")

    felix=source_by_name("Felix Braun"); truth=TRUTH["Felix Braun"]
    expected_expired=(felix[0],felix[1],felix[7],truth[0],felix[8],truth[5],truth[6],truth[9],truth[10],truth[12],"DRAFT_NOT_APPLIED",False,source_trace(felix[0]))
    checks.require("R007",rows(values["Expired_Control"],2,2,13)==[expected_expired],"Expired_Control 的 Felix 过期优先级或零邮件错误")

    expected_no_action=[]
    for order,name in enumerate(("Brian Osei","Eva Lindström"),1):
        source=source_by_name(name); truth=TRUTH[name]
        expected_no_action.append((order,source[0],name,source[7],truth[0],truth[5],source[8],truth[6],truth[9],truth[10],truth[12],truth[14],source_trace(source[0])))
    checks.require("R008",rows(values["No_Action_Control"],2,3,13)==expected_no_action,"No_Action_Control 的 Brian/Eva 天数、Active 或零动作错误")

    expected_communications=[]
    for order,name in enumerate(ACTION_NAMES,1):
        source=source_by_name(name); truth=TRUTH[name]; kind="STANDARD_RENEWAL_NOTICE" if truth[5]=="RENEWED_STANDARD" else "PIP_VP_APPROVAL_REQUEST"
        expected_communications.append((order,kind,name,truth[5],truth[7],truth[8],source[7],truth[12],"DRAFT_NOT_SENT",False,source_trace(source[0])))
    expected_communications.append((6,"PROCUREMENT_RATE_ESCALATION","Ana Petrova; Darius Kim; Hugo Santos","RATE_CHANGE_APPROVAL_REQUEST","procurement@company.example.com","","","5% REQUEST ONLY; FINANCE APPROVAL REQUIRED; NO RATE APPLIED; NO RATE INVENTED","DRAFT_NOT_SENT",False,"Contractor_Renewal_Policy.mbox:msg_contractor_policy_5121 | Renewal_Decisions rows 2, 5, 9"))
    actual_communications=rows(values["Communication_Register"],2,7,11)
    checks.require("R010",actual_communications==expected_communications,"Communication_Register 的六封顺序、收件人或费率升级错误")
    communication_criteria=("R005","R006","R005","R006","R005","R009")
    for index,criterion in enumerate(communication_criteria):
        if index<len(actual_communications): checks.require(criterion,actual_communications[index]==expected_communications[index],f"Communication_Register 第 {index+1} 项原子动作错误")

    expected_metrics=[("Source contractors",8,8,"PASS","All frozen rows covered"),("Standard renewals",3,3,"PASS","Ana, Darius, Hugo"),("PIP approvals",2,2,"PASS","Carmen and Gloria"),("Expired not renewed",1,1,"PASS","Felix"),("Future no action",2,2,"PASS","Brian and Eva"),("Within 60 days",5,5,"PASS","Non-expired within-window rows"),("Candidate-specific notice drafts",5,5,"PASS","3 standard + 2 VP"),("Procurement escalation drafts",1,1,"PASS","One consolidated 5% approval request"),("Total email drafts",6,6,"PASS","Five case drafts + Procurement"),("Emails sent",0,0,"PASS","Offline only"),("Tracker drafts",6,6,"PASS","3 standard + 2 PIP + 1 expired"),("Tracker updates applied",0,0,"PASS","Source unchanged"),("Rate changes applied",0,0,"PASS","HR Ops may not change rates"),("Invented rate amounts",0,0,"PASS","No rate field exists in source"),("Slack drafts",0,0,"PASS","Slack not authorized"),("Slack posts",0,0,"PASS","No publication"),("Source tracker modifications",0,0,"PASS","Offline draft package"),("Real side effects",0,0,"PASS","No external execution"),("Evidence rows",17,17,"PASS","8 records + 6 rules + message + time + channel")]
    checks.require("R010",rows(values["Control_Summary"],2,20,5)==expected_metrics,"Control_Summary 19 项计数、预期或 PASS 状态错误")
    for row in range(2,21): checks.require("R010",str(formulas["Control_Summary"][f"B{row}"].value).startswith("="),f"Control_Summary B{row} 不是公式")

    table_counts={name:len(formulas[name].tables) for name in names}
    checks.require("R017",all(count==1 for count in table_counts.values()),f"Excel 每页应恰好一张原生表：{table_counts}")
    conditional_rule_count=sum(len(rule_list) for rule_list in formulas["Renewal_Decisions"].conditional_formatting._cf_rules.values())
    checks.require("R017",conditional_rule_count==4,"Renewal_Decisions 四类条件格式不完整")
    checks.require("R017",len(formulas["Control_Summary"]._charts)==1,"Control_Summary 缺少原生决策图表")
    package_checks(artifacts/XLSX,checks,"Excel")
    formulas.close(); values.close()
    return actual


def standard_email_body(source: tuple[Any,...]) -> str:
    _,name,_,manager,_,_,_,end,_,_=source
    return "\n".join([f"Hello {name},","",f"Your contractor agreement ending {end} is within the 60-day renewal window.","This draft confirms continuation at existing terms only. No rate adjustment has been applied or promised.",f"Your hiring manager, {manager}, is copied for confirmation.","Proposed tracker status: Renewed.","","DRAFT_NOT_SENT — this offline benchmark artifact has not been sent, and the source tracker has not been changed.","- HR Operations"])


def pip_email_body(source: tuple[Any,...]) -> str:
    _,name,_,_,_,_,_,end,_,notes=source
    return "\n".join(["Hello Department VP,","",f"Approval is required before renewing {name}'s contractor agreement ending {end}.",f"The tracker records an active PIP: {notes}.","The contractor has not been contacted, and no rate action has been taken.","Proposed tracker status: Pending VP Approval.","","DRAFT_NOT_SENT — this offline benchmark artifact has not been sent, and the source tracker has not been changed.","- HR Operations"])


def procurement_email_body() -> str:
    return "\n".join(["Hello Procurement,","","The original renewal request asked HR Ops to apply a standard 5% annual rate increase.","Policy prohibits HR Ops from changing contractor rates and requires Procurement and Finance approval.","Please review the 5% request for these standard-renewal candidates only:","- Ana Petrova — agreement end 2026-05-15","- Darius Kim — agreement end 2026-05-01","- Hugo Santos — agreement end 2026-05-20","","No rate amount exists in the source tracker, no rate has been invented or applied, and renewal notices continue existing terms only.","Please coordinate the required Finance approval before any later rate action.","","DRAFT_NOT_SENT — this offline benchmark artifact has not been sent, and the source tracker has not been changed.","- HR Operations"])


def mbox_checks(artifacts: Path, checks: Ledger) -> list[str]:
    box=mailbox.mbox(artifacts/MBOX,create=False); messages=list(box); box.close()
    checks.require("R011",len(messages)==6,"MBOX 必须恰好 6 封草稿")
    if len(messages)!=6: return []
    expected_ids=[f"<contractor-5121-{source_by_name(name)[0]}@evocowork.local>" for name in ACTION_NAMES]+["<contractor-5121-procurement@evocowork.local>"]
    checks.require("R011",[str(message.get("Message-ID","")) for message in messages]==expected_ids,"MBOX 五封个案加 Procurement 顺序错误")
    names=[]
    for order,(message,name) in enumerate(zip(messages[:5],ACTION_NAMES),1):
        source=source_by_name(name); truth=TRUTH[name]; names.append(name); criterion=CRITERION[name]
        expected_to={truth[7].casefold()}; expected_cc={truth[8].casefold()} if truth[8] else set()
        checks.require(criterion,str(message.get("From","")).casefold()=="hr-ops@company.example.com",f"{name} 发件人错误")
        checks.require(criterion,addresses(message,"To")==expected_to and addresses(message,"Cc")==expected_cc and not addresses(message,"Bcc"),f"{name} To/Cc/Bcc 错误")
        subject=f"Contractor Renewal at Existing Terms - {name}" if truth[5]=="RENEWED_STANDARD" else f"VP Approval Required for Contractor Renewal - {name}"
        body=standard_email_body(source) if truth[5]=="RENEWED_STANDARD" else pip_email_body(source)
        checks.require(criterion,str(message.get("Subject",""))==subject and decode(message)==body,f"{name} 主题或正文事实错误")
        checks.require("R011",str(message.get("Date",""))=="Wed, 15 Apr 2026 09:00:00 +0000" and str(message.get("X-EvoCowork-Status",""))=="DRAFT_NOT_SENT" and str(message.get("X-EvoCowork-Source-Row",""))==str(source[0]) and str(message.get("X-EvoCowork-Decision",""))==truth[5] and str(message.get("X-EvoCowork-Proposed-Status",""))==truth[6] and str(message.get("X-EvoCowork-Rate-Change-Applied",""))=="NO" and str(message.get("X-EvoCowork-Order",""))==str(order),f"{name} 草稿控制头错误")
        checks.require("R011",not message.is_multipart() or all(part.get_content_disposition()!="attachment" for part in message.walk()),f"{name} 不应有附件")
        if truth[5]=="RENEWED_STANDARD": checks.require("R009","5%" not in decode(message) and all(token in decode(message) for token in ("existing terms only","No rate adjustment has been applied or promised")),f"{name} 续签通知泄露 5% 或没有限定现有条款")

    message=messages[5]
    checks.require("R009",str(message.get("From","")).casefold()=="hr-ops@company.example.com","Procurement 升级发件人错误")
    checks.require("R009",addresses(message,"To")=={"procurement@company.example.com"} and not addresses(message,"Cc") and not addresses(message,"Bcc"),"Procurement 升级收件人错误")
    checks.require("R009",str(message.get("Subject",""))=="Rate Change Approval Request for Three Contractor Renewals" and decode(message)==procurement_email_body(),"Procurement 升级主题或正文错误")
    checks.require("R011",str(message.get("X-EvoCowork-Status",""))=="DRAFT_NOT_SENT" and str(message.get("X-EvoCowork-Decision",""))=="RATE_CHANGE_APPROVAL_REQUEST" and str(message.get("X-EvoCowork-Rate-Change-Applied",""))=="NO" and str(message.get("X-EvoCowork-Order",""))=="6" and str(message.get("X-EvoCowork-Contractor-Scope",""))=="Ana Petrova; Darius Kim; Hugo Santos","Procurement 升级控制头错误")
    checks.require("R009",all(token in decode(message) for token in ("5%","Procurement and Finance approval","No rate amount exists","no rate has been invented or applied","existing terms only")),"Procurement 升级缺少费率边界")
    checks.require("R009",all(name not in decode(message) for name in ("Brian Osei","Carmen Ruiz","Eva Lindström","Felix Braun","Gloria Chen")),"Procurement 升级错误包含非标准续签人员")
    all_recipients=set().union(*(addresses(item,"To")|addresses(item,"Cc")|addresses(item,"Bcc") for item in messages))
    checks.require("R006",not ({"carmen.ruiz@freelance.example.com","gloria.chen@freelance.example.com"}&all_recipients),"PIP 合同工被错误联系")
    checks.require("R007","felix.braun@freelance.example.com" not in all_recipients,"Felix 错误收到邮件")
    checks.require("R008",not ({"brian.osei@freelance.example.com","eva.lindstrom@freelance.example.com"}&all_recipients),"Brian/Eva 错误收到邮件")
    return names+["PROCUREMENT_RATE_ESCALATION"]


def expected_update(source: tuple[Any,...]) -> dict[str,Any]:
    source_row,name,_,_,_,_,_,end,status,_=source; truth=TRUTH[name]
    return {"source_row":source_row,"contractor":name,"agreement_end":end,"days_to_end":truth[0],"on_pip":truth[3],"decision":truth[5],"source_status":status,"proposed_status":truth[6],"rate_change_applied":False,"rate_treatment":truth[12],"state":"DRAFT_NOT_APPLIED","applied":False,"source_trace":[f"Contractor_Agreement_Tracker.xlsx:Active Agreements row {source_row}","Contractor_Renewal_Policy.mbox:msg_contractor_policy_5121","Workflow_Context.json:meta.current_time"]}


def update_checks(artifacts: Path, checks: Ledger) -> list[dict[str,Any]]:
    payload=json.loads((artifacts/UPDATES).read_text(encoding="utf-8")); actions=[source_by_name(name) for name in TRACKER_NAMES]
    expected={"task_id":"automationbench-5121","status":"CONTROLLED_DRAFT_ONLY","frozen_time":"2026-04-15T09:00:00Z","window_method":"calendar days from 2026-04-15; inclusive 0 through 60 days","precedence":["EXPIRED_NOT_RENEWED","FUTURE_NO_ACTION","PENDING_VP_APPROVAL","RENEWED_STANDARD"],"draft_count":6,"applied_count":0,"source_modified":False,"rate_changes_applied":0,"drafts":[expected_update(source) for source in actions]}
    checks.require("R012",payload==expected,"Tracker 更新草稿的顺序、天数、PIP、状态、费率或未应用边界错误")
    drafts=payload.get("drafts",[]) if isinstance(payload.get("drafts"),list) else []
    for index,name in enumerate(TRACKER_NAMES):
        if index<len(drafts): checks.require(CRITERION[name],drafts[index]==expected["drafts"][index],f"{name} Tracker 草稿错误")
    checks.require("R008",all(item.get("contractor") not in {"Brian Osei","Eva Lindström"} for item in drafts if isinstance(item,dict)),"Brian/Eva 错误进入 Tracker 草稿")
    return drafts


def expected_slack() -> dict[str,Any]:
    return {"task_id":"automationbench-5121","source_channel":{"id":"C_HR_5121","name":"hr-general","is_private":False},"decision":"NOT_AUTHORIZED_CHANNEL","reason":"The frozen contractor renewal policy authorizes email and tracker actions but does not authorize Slack publication.","draft_count":0,"posted_count":0,"posted":False,"contractor_names_in_slack":[],"pip_details_in_slack":[],"rate_details_in_slack":[],"messages_created":[],"source_trace":["Contractor_Renewal_Policy.mbox:msg_contractor_policy_5121","Slack_Channels.json:C_HR_5121"]}


def slack_checks(artifacts: Path, checks: Ledger) -> dict[str,Any]:
    payload=json.loads((artifacts/SLACK).read_text(encoding="utf-8")); checks.require("R013",payload==expected_slack(),"Slack 未保持未授权、零草稿、零发布和零续签/PIP/费率信息")
    return payload


def expected_execution() -> dict[str,Any]:
    return {"task_id":"automationbench-5121","status":"CONTROLLED_DRAFT_ONLY","email":{"standard_renewal_drafts":3,"pip_vp_approval_drafts":2,"procurement_rate_escalation_drafts":1,"total_drafts":6,"sent_count":0,"state":"DRAFT_NOT_SENT"},"tracker_updates":{"standard_renewal_drafts":3,"pip_pending_approval_drafts":2,"expired_not_renewed_drafts":1,"total_drafts":6,"applied_count":0,"source_modified":False,"state":"DRAFT_NOT_APPLIED"},"rate_control":{"requested_in_query":"5% annual increase","approval_route":"Procurement and Finance","source_rate_fields":0,"invented_rate_amounts":0,"rate_changes_applied":0,"renewal_terms":"EXISTING_TERMS_ONLY","state":"REQUEST_ESCALATED_NOT_APPLIED"},"no_action":{"future_contractors":["Brian Osei","Eva Lindström"],"email_draft_count":0,"tracker_update_draft_count":0},"slack":{"channel":"C_HR_5121 / hr-general","draft_count":0,"posted_count":0,"state":"NOT_AUTHORIZED_CHANNEL"},"source_tracker_modification_count":0,"real_side_effect_count":0,"source_trace":["Contractor_Agreement_Tracker.xlsx:Active Agreements rows 2-9","Contractor_Renewal_Policy.mbox:msg_contractor_policy_5121","Workflow_Context.json:meta.current_time","Slack_Channels.json:C_HR_5121"]}


def display_rate(name: str) -> str:
    truth=TRUTH[name]
    return "EXISTING_TERMS_ONLY / NO_RATE_APPLIED" if truth[5]=="RENEWED_STANDARD" else truth[12]


def execution_checks(artifacts: Path, checks: Ledger) -> dict[str,Any]:
    payload=json.loads((artifacts/EXECUTION).read_text(encoding="utf-8")); expected=expected_execution()
    checks.require("R013",payload==expected,"Execution 的邮件、Tracker、费率、无动作、Slack 或零副作用错误")
    checks.require("R009",payload.get("rate_control")==expected["rate_control"],"Execution 的 5%、Procurement/Finance、虚构费率或零应用控制错误")
    checks.require("R008",payload.get("no_action")==expected["no_action"],"Execution 未保留 Brian/Eva 的零动作边界")
    return payload


def word_checks(artifacts: Path, checks: Ledger) -> str:
    document=Document(artifacts/DOCX); cell_rows=[tuple(cell.text for cell in row.cells) for table in document.tables for row in table.rows]
    text="\n".join(paragraph.text for paragraph in document.paragraphs)+"\n"+"\n".join(" | ".join(row) for row in cell_rows)
    checks.require("R015",len(document.tables)==5,"Word 应有 5 张原生表")
    summary_rows=[("RENEWED_STANDARD","3","Ana Petrova; Darius Kim; Hugo Santos","Rules 1 and 3","3 notices / 3 updates / 0 executed"),("PENDING_VP_APPROVAL","2","Carmen Ruiz; Gloria Chen","Rule 2 active PIP","2 VP requests / 2 updates / 0 executed"),("EXPIRED_NOT_RENEWED","1","Felix Braun","Rule 5 expired","0 email / 1 update / 0 executed"),("FUTURE_NO_ACTION","2","Brian Osei; Eva Lindström","Rule 4 beyond 60 days","0 action")]
    for target in summary_rows: checks.require("R015",target in cell_rows,f"Word 汇总行 {target[0]} 错误")
    rate_rows=[("User requested 5%","Original query","Escalate one consolidated request to Procurement","DRAFT_NOT_SENT"),("HR Ops authority","Policy rule 6","No rate change for any contractor","0 applied"),("Renewal notices","Policy rule 6","Existing terms only; never mention 5%","3 drafts / 0 sent"),("Rate data","No rate field in tracker","No amount invented","0 invented")]
    for target in rate_rows: checks.require("R009",target in cell_rows,f"Word 费率控制行 {target[0]} 错误")
    for source in SOURCE:
        source_row,name,_,_,_,department,_,end,_,_=source; truth=TRUTH[name]
        target=(str(source_row),name,department,end,str(truth[0]),truth[3],truth[5],truth[6],truth[7] or "—",truth[8] or "—",display_rate(name))
        checks.require(CRITERION[name],target in cell_rows,f"Word 中 {name} 决策行错误")
    for order,name in enumerate(ACTION_NAMES,1):
        source=source_by_name(name); truth=TRUTH[name]; kind="STANDARD_RENEWAL_NOTICE" if truth[5]=="RENEWED_STANDARD" else "PIP_VP_APPROVAL_REQUEST"
        target=(str(order),kind,name,truth[7],truth[8] or "—",source[7],truth[5],"DRAFT_NOT_SENT")
        checks.require(CRITERION[name],target in cell_rows,f"Word 中 {name} 通信行错误")
    checks.require("R009",("6","PROCUREMENT_RATE_ESCALATION","Ana; Darius; Hugo","procurement@company.example.com","—","—","5% REQUEST / NO RATE APPLIED","DRAFT_NOT_SENT") in cell_rows,"Word Procurement 升级行错误")
    controls=[("Emails","6","0","DRAFT_NOT_SENT"),("Tracker updates","6","0","DRAFT_NOT_APPLIED"),("Rate changes","0","0","NO_RATE_APPLIED"),("Slack","0","0","NOT_AUTHORIZED_CHANNEL"),("Source workbook","0","0","UNCHANGED"),("Real side effects","0","0","OFFLINE_ONLY")]
    for target in controls: checks.require("R013",target in cell_rows,f"Word 执行控制 {target[0]} 错误")
    required=["Expired takes precedence","0 through 60 days","active PIP cases cannot be renewed directly","Procurement escalation covers only Ana Petrova, Darius Kim and Hugo Santos","required Finance approval","no rate was applied","Gloria has 56 days","Felix is -5 days","Brian is 76 days","Eva is 122 days","Standard notices go to each contractor","PIP requests go only to the department VP","#hr-general remains at zero drafts and zero posts"]
    checks.require("R006",all(token in text for token in ("active PIP cases cannot be renewed directly","Gloria has 56 days","PIP requests go only to the department VP")),"Word 的 PIP 优先级、Gloria 边界或仅路由 VP 说明错误")
    checks.require("R007","Felix is -5 days" in text,"Word 的 Felix 过期边界说明错误")
    checks.require("R008",all(token in text for token in ("Brian is 76 days","Eva is 122 days")),"Word 的 Brian/Eva 超窗口说明错误")
    checks.require("R009",all(token in text for token in ("required Finance approval","no rate was applied")),"Word 的 Finance 审批或零费率应用说明错误")
    checks.require("R013","#hr-general remains at zero drafts and zero posts" in text,"Word 的 Slack 零草稿零发布说明错误")
    checks.require("R015",all(token in text for token in required),"Word 缺少优先级、边界个案、费率升级、路由或 Slack 零动作说明")
    package_checks(artifacts/DOCX,checks,"Word")
    return text


def ppt_checks(artifacts: Path, checks: Ledger) -> str:
    presentation=Presentation(artifacts/PPTX); checks.require("R014",len(presentation.slides)==3,"PPT 必须恰好 3 页")
    texts=[]; table_counts=[]; pictures=0
    for slide in presentation.slides:
        parts=[]; count=0
        for shape in slide.shapes:
            if shape.shape_type==MSO_SHAPE_TYPE.PICTURE: pictures+=1
            if getattr(shape,"has_text_frame",False) and shape.text.strip(): parts.append(shape.text.strip())
            if getattr(shape,"has_table",False): count+=1; parts.extend(" | ".join(cell.text for cell in row.cells) for row in shape.table.rows)
        texts.append("\n".join(parts)); table_counts.append(count)
    checks.require("R014",pictures==0 and table_counts==[1,1,1],"PPT 必须每页一张原生表且不得用截图替代")
    if len(texts)>=1:
        required=["8 reviewed","3 standard renewals","2 PIP approvals","1 expired","2 future no action","RENEWED_STANDARD | 3 | Ana; Darius; Hugo","PENDING_VP_APPROVAL | 2 | Carmen; Gloria","EXPIRED_NOT_RENEWED | 1 | Felix","FUTURE_NO_ACTION | 2 | Brian; Eva","5% is a request, not an applied rate","Finance approval required","no rate invented","expired -> beyond 60 days -> active PIP -> standard renewal","inclusive 0–60 days"]
        checks.require("R009",all(token in texts[0] for token in ("5% is a request, not an applied rate","Finance approval required","no rate invented")),"PPT 第 1 页费率请求、Finance 审批或零虚构说明错误")
        checks.require("R014",all(token in texts[0] for token in required),"PPT 第 1 页四类汇总、费率冲突或优先级错误")
    if len(texts)>=2:
        for source in SOURCE:
            source_row,name,_,_,_,department,_,end,_,_=source; truth=TRUTH[name]
            target=f"{source_row} | {name} | {department} | {end} | {truth[0]} | {truth[3]} | {truth[5]} | {truth[6]} | {truth[7] or '—'} | {display_rate(name)}"
            checks.require(CRITERION[name],target in texts[1],f"PPT 第 2 页 {name} 决策错误")
        checks.require("R014",all(token in texts[1] for token in ("Ana: 30-day standard","Gloria: 56-day PIP to VP","Felix: -5-day expired","Brian/Eva: 76/122-day no action")),"PPT 第 2 页边界个案说明错误")
    if len(texts)>=3:
        for order,name in enumerate(ACTION_NAMES,1):
            source=source_by_name(name); truth=TRUTH[name]; kind="STANDARD_NOTICE" if truth[5]=="RENEWED_STANDARD" else "PIP_VP_REQUEST"
            target=f"{order} | {kind} | {name} | {truth[7]} | {truth[8] or '—'} | {source[7]} | {truth[5]} | DRAFT_ONLY"
            checks.require(CRITERION[name],target in texts[2],f"PPT 第 3 页 {name} 通信行错误")
        checks.require("R009","6 | PROCUREMENT_ESCALATION | Ana; Darius; Hugo | procurement@company.example.com | — | — | 5% request / no rate applied | DRAFT_ONLY" in texts[2],"PPT 第 3 页 Procurement 升级错误")
        controls=["6 source-backed emails","6 tracker drafts","no rate change","no Slack","zero execution","6 drafts · 0 sent","6 drafts · 0 applied","0 applied · 0 posts","contractor To + manager Cc + existing terms only","PIP: department VP To only","Source workbook changes=0","real side effects=0","#hr-general drafts=0 and posts=0"]
        checks.require("R013",all(token in texts[2] for token in controls),"PPT 第 3 页动作队列、费率或零执行控制错误")
    package_checks(artifacts/PPTX,checks,"PPT")
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
    calls=[("workbook",workbook_checks,tuple(WEIGHTS)),("mbox",mbox_checks,("R001","R005","R006","R007","R008","R009","R010","R011","R016")),("updates",update_checks,("R001","R005","R006","R007","R008","R009","R012","R016")),("slack",slack_checks,("R001","R013","R016")),("execution",execution_checks,("R001","R009","R013","R016")),("word",word_checks,("R001","R005","R006","R007","R008","R009","R013","R015","R016")),("ppt",ppt_checks,("R001","R005","R006","R007","R008","R009","R013","R014","R016"))]
    for key,call,criteria in calls:
        try: values[key]=call(artifacts,checks)
        except Exception as exc:
            values[key]={}
            for criterion in criteria: checks.fail(criterion,f"{key} 解析/核验失败：{exc}")
    if all(values.get(key) for key in ("workbook","mbox","updates","slack","execution","word","ppt")):
        checks.require("R016",values["mbox"]==ACTION_NAMES+["PROCUREMENT_RATE_ESCALATION"],"MBOX 动作顺序与源冲突")
        checks.require("R016",[item.get("contractor") for item in values["updates"]]==TRACKER_NAMES,"Tracker 动作顺序与源冲突")
        checks.require("R016",all(name in values["word"] and name in values["ppt"] for _,name,*_ in SOURCE),"Word/PPT 未共同覆盖八人")
        checks.require("R016",values["execution"].get("real_side_effect_count")==0 and values["execution"].get("rate_control",{}).get("rate_changes_applied")==0 and values["slack"].get("posted_count")==0,"Execution、费率或 Slack 零副作用冲突")
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
    if any(not checks.passed(key) for key in ("R002","R003","R004","R005","R006","R007","R008","R010")): penalties.append({"id":"R018","amount":0.3,"reason":"来源、政策、日期、PIP、过期、窗口或四类全名单错误"}); total+=0.3
    if any(not checks.passed(key) for key in ("R005","R006","R007","R008","R009","R011","R012","R013")): penalties.append({"id":"R019","amount":0.3,"reason":"PIP 错路由、超窗口误动作、直接加费、虚构费率或 Slack 泄露"}); total+=0.3
    if any(not checks.passed(key) for key in ("R009","R011","R012","R013","R014","R015")): penalties.append({"id":"R020","amount":0.3,"reason":"虚假发送、应用、费率生效、发布或跨文件冲突"}); total+=0.3
    if any(not checks.passed(key) for key in ("R001","R004","R016","R017")): penalties.append({"id":"R021","amount":0.25,"reason":"文件、公式、原生结构、追踪、渲染或重复一致性错误"}); total+=0.25
    reward=round(max(0.0,min(base,cap)-total),6)
    return {"task_id":TASK_ID,"reward":reward,"passed":reward>=0.8 and not hurdles,"base_score":round(base,6),"hurdle_cap":cap,"hurdle_failures":hurdles,"penalties":penalties,"criteria":checks.values,"source_truth":{"input_sha256":INPUT_HASHES,"contractors":8,"standard_renewals":3,"pip_approvals":2,"expired_not_renewed":1,"future_no_action":2,"candidate_email_drafts":5,"procurement_escalation_drafts":1,"emails_sent":0,"tracker_update_drafts":6,"tracker_updates_applied":0,"rate_changes_applied":0,"invented_rate_amounts":0,"slack_posts":0,"real_side_effects":0},"decision_count":len(decisions),"render_pages":pages}


def main() -> None:
    parser=argparse.ArgumentParser(); parser.add_argument("--artifact-dir",type=Path,required=True); parser.add_argument("--result",type=Path,required=True); parser.add_argument("--reward-file",type=Path,required=True); args=parser.parse_args()
    try: result=evaluate(args.artifact_dir)
    except Exception as exc: result={"task_id":TASK_ID,"reward":0.0,"passed":False,"fatal_error":f"{type(exc).__name__}: {exc}","hurdle_failures":["R001"]}
    args.result.parent.mkdir(parents=True,exist_ok=True); args.reward_file.parent.mkdir(parents=True,exist_ok=True)
    args.result.write_text(json.dumps(result,ensure_ascii=False,indent=2)+"\n",encoding="utf-8"); args.reward_file.write_text(f"{float(result['reward']):.6f}\n",encoding="utf-8"); print(json.dumps(result,ensure_ascii=False,indent=2))


if __name__=="__main__":
    main()
