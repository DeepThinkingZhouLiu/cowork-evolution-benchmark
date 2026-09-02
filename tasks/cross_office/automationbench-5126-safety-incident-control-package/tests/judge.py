#!/usr/bin/env python3
"""AutomationBench 5126 安全事故治理题的逐题确定性 Judge。

Judge 独立读取四个冻结输入并重建八起事故真值，
不导入参考解法，也不依赖 Golden 文件哈希。
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


TASK_ID="cowork-evo/automationbench-5126-safety-incident-control-package"
XLSX="Safety_Incident_Audit.xlsx"; DOCX="Safety_Incident_Control_Brief.docx"; PPTX="Safety_Incident_Review.pptx"; MBOX="Safety_Notification_Drafts.mbox"
UPDATES="Safety_Tracker_Update_Drafts.json"; SLACK="Safety_Slack_Summary_Draft.json"; EXECUTION="Safety_Execution_Control.json"
REQUIRED={XLSX,DOCX,PPTX,MBOX,UPDATES,SLACK,EXECUTION}
INPUT_HASHES={
    "Safety_Incident_Tracker.xlsx":"961eb03fb913283fc116d114793b15aa39277d5c043eccf5a48629e7341653f3",
    "Safety_Manual.mbox":"81e2472039256c090ac1fad2014d76c6590c5c4d10ec4bd0c70b82f0b009457b",
    "Slack_Channels.json":"6f4850069e074815ee1b160108497d37c4a1601cf2f37a15d65b11aef7fa9f92",
    "Workflow_Context.json":"76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99",
}
WEIGHTS={"R001":4,"R002":5,"R003":7,"R004":9,"R005":5,"R006":6,"R007":8,"R008":9,"R009":5,"R010":4,"R011":8,"R012":5,"R013":5,"R014":5,"R015":4,"R016":4,"R017":4,"R018":3}
HURDLE_CAPS={"R001":0.0,"R002":0.2,"R003":0.2,"R004":0.3,"R005":0.2,"R006":0.2,"R007":0.15,"R008":0.1,"R009":0.25,"R010":0.3,"R011":0.15,"R012":0.2,"R013":0.15,"R014":0.15,"R015":0.3,"R016":0.3}
SOURCE=[
    (2,"Ana Petrova","mgr-a@company.example.com","2026-04-14","Paper cut from filing cabinet","Minor","First aid - bandage applied","0","Pending",""),
    (3,"Ben Torres","mgr-b@company.example.com","2026-04-13","Slipped on wet floor, sprained ankle","Moderate","Urgent care visit, ankle brace","2","Pending",""),
    (4,"Carlos Mendoza","mgr-a@company.example.com","2026-04-12","Fell from loading dock, broken leg","Serious","Hospitalization - emergency surgery","TBD","Pending",""),
    (5,"Diana Osei","mgr-b@company.example.com","2026-04-14","Ergonomic discomfort - wrist pain","Minor","Ice pack, ergonomic assessment scheduled","0","Pending",""),
    (6,"Ethan Park","mgr-a@company.example.com","2026-04-11","Chemical splash in lab - eye injury","Serious","Hospitalization - potential loss of vision","5+","Pending",""),
    (7,"Fiona Strand","mgr-b@company.example.com","2026-04-14","Bumped head on low clearance doorway","Minor","Ice pack applied","0","Pending",""),
    (8,"George Tan","mgr-a@company.example.com","2026-04-10","Repetitive strain - back pain from lifting","Moderate","Physical therapy referral","1","Pending",""),
    (9,"Holly Ng","mgr-b@company.example.com","2026-04-13","Machine guard failure - finger amputation","Serious","Hospitalization - surgical amputation","TBD","Pending",""),
]
POLICY_ROWS=[
    ("P001","Classification","Minor","First aid only; no lost work time","Safety_Manual.mbox:Chapter 2"),
    ("P002","Classification","Moderate","Medical treatment; fewer than 3 lost work days","Safety_Manual.mbox:Chapter 2"),
    ("P003","Classification","Serious","Hospitalization, more than 3 lost work days, or restricted duty","Safety_Manual.mbox:Chapter 2"),
    ("P004","Internal routing","Minor","Manager only; Status=Documented","Safety_Manual.mbox:Chapter 3"),
    ("P005","Internal routing","Moderate","Safety Officer + manager; Status=Under Review","Safety_Manual.mbox:Chapter 3"),
    ("P006","Internal routing","Serious","Safety Officer + manager + VP Operations; Status=Investigation","Safety_Manual.mbox:Chapter 3"),
    ("P007","Workers' compensation","Moderate/Serious","Claim form required; Benefits handles separately","Safety_Manual.mbox:Chapter 4"),
    ("P008","OSHA","Hospitalization/amputation/loss of eye","Legal immediately; OSHA within 24 hours; precedence","Safety_Manual.mbox:Chapter 5"),
    ("P009","Return to work","Return from injury leave","Assessment required","Safety_Manual.mbox:Chapter 6"),
    ("P010","Retention","All incidents","Retain records for 5 years","Safety_Manual.mbox:Chapter 7"),
]
TRUTH={
    "Ana Petrova":("NO","NO","NO","NO","Documented","NO","NO","mgr-a@company.example.com","Safety Incident Documentation - Ana Petrova - 2026-04-14","P004,P010"),
    "Ben Torres":("NO","NO","NO","NO","Under Review","YES","NO","safety-officer@company.example.com; mgr-b@company.example.com","Safety Incident Under Review - Ben Torres - 2026-04-13","P005,P007,P010"),
    "Carlos Mendoza":("YES","NO","NO","YES","Investigation","YES","YES","legal@company.example.com; safety-officer@company.example.com; mgr-a@company.example.com; vp-ops@company.example.com","URGENT OSHA Review - Carlos Mendoza - 2026-04-12","P006,P007,P008,P010"),
    "Diana Osei":("NO","NO","NO","NO","Documented","NO","NO","mgr-b@company.example.com","Safety Incident Documentation - Diana Osei - 2026-04-14","P004,P010"),
    "Ethan Park":("YES","NO","POTENTIAL_VISION_LOSS","YES","Investigation","YES","YES","legal@company.example.com; safety-officer@company.example.com; mgr-a@company.example.com; vp-ops@company.example.com","URGENT OSHA Review - Ethan Park - 2026-04-11","P006,P007,P008,P010"),
    "Fiona Strand":("NO","NO","NO","NO","Documented","NO","NO","mgr-b@company.example.com","Safety Incident Documentation - Fiona Strand - 2026-04-14","P004,P010"),
    "George Tan":("NO","NO","NO","NO","Under Review","YES","NO","safety-officer@company.example.com; mgr-a@company.example.com","Safety Incident Under Review - George Tan - 2026-04-10","P005,P007,P010"),
    "Holly Ng":("YES","YES","NO","YES","Investigation","YES","YES","legal@company.example.com; safety-officer@company.example.com; mgr-b@company.example.com; vp-ops@company.example.com","URGENT OSHA Review - Holly Ng - 2026-04-13","P006,P007,P008,P010"),
}
PRIMARY={"Ana Petrova":"R005","Diana Osei":"R005","Fiona Strand":"R005","Ben Torres":"R006","George Tan":"R006","Carlos Mendoza":"R007","Ethan Park":"R007","Holly Ng":"R007"}


@dataclass
class Ledger:
    values: dict[str,dict[str,Any]]
    def __init__(self)->None: self.values={key:{"passed":True,"details":[]} for key in WEIGHTS}
    def require(self,key:str,condition:bool,detail:str)->None:
        if not condition: self.values[key]["passed"]=False; self.values[key]["details"].append(detail)
    def fail(self,key:str,detail:str)->None: self.require(key,False,detail)
    def passed(self,key:str)->bool: return bool(self.values[key]["passed"])


def root()->Path: return Path(__file__).resolve().parents[1]
def input_root()->Path:
    container=Path("/app/input_files"); return container if container.is_dir() else root()/"data/input_files"
def sha256(path:Path)->str:
    digest=hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda:handle.read(1024*1024),b""): digest.update(chunk)
    return digest.hexdigest()
def norm(value:Any)->Any:
    if value is None:return ""
    if isinstance(value,datetime):return value.date().isoformat()
    if isinstance(value,date):return value.isoformat()
    if isinstance(value,float) and value.is_integer():return int(value)
    return value
def rows(sheet,start:int,end:int,columns:int)->list[tuple[Any,...]]: return [tuple(norm(sheet.cell(row,column).value) for column in range(1,columns+1)) for row in range(start,end+1)]
def formula(value:Any,*tokens:str)->bool:
    text=str(value or ""); return text.startswith("=") and all(token.casefold() in text.casefold() for token in tokens)
def decode(message)->str:
    if message.is_multipart():
        parts=[]
        for part in message.walk():
            if part.get_content_type()=="text/plain": parts.append((part.get_payload(decode=True) or b"").decode(part.get_content_charset() or "utf-8",errors="replace"))
        return "\n".join(parts).strip()
    return (message.get_payload(decode=True) or b"").decode(message.get_content_charset() or "utf-8",errors="replace").strip()
def address_list(message,header:str)->list[str]: return [address.casefold() for _,address in getaddresses(message.get_all(header,[])) if address]
def by_name(name:str)->tuple[Any,...]: return next(item for item in SOURCE if item[1]==name)
def trace(row:int)->str: return f"Safety_Incident_Tracker.xlsx:Recent Incidents row {row} | Safety_Manual.mbox:msg_safety_manual_5126 | Workflow_Context.json:meta.current_time"


def source_manual()->dict[str,str]:
    box=mailbox.mbox(input_root()/"Safety_Manual.mbox",create=False); messages=list(box); box.close()
    if len(messages)!=1: raise ValueError("源 MBOX 必须恰好一封")
    message=messages[0]; result={"message_id":str(message.get("X-AutomationBench-Message-ID","")),"thread_id":str(message.get("X-AutomationBench-Thread-ID","")),"from":str(message.get("From","")),"to":str(message.get("To","")),"subject":str(message.get("Subject","")),"date":str(message.get("Date","")),"labels":str(message.get("X-AutomationBench-Labels","")),"body":decode(message)}
    if result["message_id"]!="msg_safety_manual_5126" or result["from"].casefold()!="safety-officer@company.example.com" or result["subject"]!="Workplace Safety Incident Reporting Manual - 2026": raise ValueError("源安全手册身份错误")
    tokens=("Minor: First aid only","Moderate: Medical treatment needed","Serious: Hospitalization","Minor: Email the employee's manager","Moderate: Email the Safety Officer","Serious: Email the Safety Officer","workers' comp claim form","hospitalization, amputation, or loss of an eye","reported to OSHA within 24 hours","legal@company.example.com","retained for 5 years")
    if not all(token in result["body"] for token in tokens): raise ValueError("源安全手册内容错误")
    return result


def expected_decision(source:tuple[Any,...])->tuple[Any,...]:
    source_row,employee,manager,incident_date,description,severity,treatment,lost_days,status,notes=source; truth=TRUTH[employee]
    return (source_row,employee,manager,incident_date,description,severity,treatment,lost_days,status,"2026-04-15T09:00:00Z","YES" if severity=="Minor" else "NO","YES" if severity=="Moderate" else "NO","YES" if severity=="Serious" else "NO",*truth[:7],truth[7],truth[8],"AGGREGATE_ONLY","YES","DRAFT_NOT_SENT","NO","NO","NO",5,truth[9],trace(source_row))


def expected_evidence(manual:dict[str,str])->list[tuple[Any,...]]:
    values=[]
    for index,source in enumerate(SOURCE,1): values.append((f"E{index:03d}","SOURCE_INCIDENT",source[1],f"{source[3]} | {source[5]} | {source[6]}",f"Safety_Incident_Tracker.xlsx:Recent Incidents row {source[0]}"))
    values.extend([
        ("E009","MANUAL","Reporting manual",manual["subject"],"Safety_Manual.mbox:msg_safety_manual_5126"),("E010","TIME","Frozen time","2026-04-15T09:00:00Z","Workflow_Context.json:meta.current_time"),("E011","CHANNEL","Safety channel","C_SAF_5126 | safety | public","Slack_Channels.json:C_SAF_5126"),("E012","CONTROL","Minor count","3","Incident_Decisions:Severity"),("E013","CONTROL","Moderate count","2","Incident_Decisions:Severity"),("E014","CONTROL","Serious count","3","Incident_Decisions:Severity"),("E015","CONTROL","OSHA/Legal immediate count","3","Incident_Decisions:OSHA Reportable"),("E016","CONTROL","Workers comp follow-up count","5","Incident_Decisions:Workers Comp Required"),("E017","CONTROL","Email draft count","8","Communication_Register"),("E018","CONTROL","Tracker update draft count","8","Tracker update payload"),("E019","CONTROL","Slack aggregate draft count","1","Slack summary payload"),("E020","CONTROL","Emails sent","0","Execution control"),("E021","CONTROL","Tracker updates applied","0","Execution control"),("E022","CONTROL","Real side effects","0","Execution control"),
    ]); return values


def render(path:Path)->tuple[int,str]:
    with tempfile.TemporaryDirectory(prefix="judge-5126-") as temporary:
        work=Path(temporary); output=work/"out"; output.mkdir(); process=subprocess.run(["libreoffice","--headless",f"-env:UserInstallation=file://{work/'profile'}","--convert-to","pdf","--outdir",str(output),str(path)],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,timeout=300); pdf=output/f"{path.stem}.pdf"
        if process.returncode!=0 or not pdf.is_file(): raise RuntimeError(process.stdout[-1000:])
        info=subprocess.run(["pdfinfo",str(pdf)],capture_output=True,text=True,timeout=60); match=re.search(r"^Pages:\s+(\d+)",info.stdout,flags=re.M); text=subprocess.run(["pdftotext",str(pdf),"-"],capture_output=True,text=True,timeout=60)
        if info.returncode!=0 or not match or text.returncode!=0: raise RuntimeError("无法读取渲染 PDF")
        return int(match.group(1)),text.stdout


def package_checks(path:Path,checks:Ledger,prefix:str)->None:
    try:
        with zipfile.ZipFile(path) as archive:
            names=archive.namelist(); checks.require("R001",not any(name.endswith("vbaProject.bin") for name in names),f"{prefix}包含宏"); checks.require("R001",not any("/embeddings/" in name for name in names),f"{prefix}包含嵌入对象"); checks.require("R001",not any(name.endswith(".rels") and b'TargetMode="External"' in archive.read(name) for name in names),f"{prefix}包含外部链接")
    except Exception as exc: checks.fail("R001",f"{prefix} OOXML包检查失败：{exc}")


def workbook_checks(artifacts:Path,checks:Ledger)->dict[str,tuple[Any,...]]:
    formulas=load_workbook(artifacts/XLSX,data_only=False); values=load_workbook(artifacts/XLSX,data_only=True)
    names=["Source_Incidents","Source_Manual","Policy_Source","Evidence_Register","Incident_Decisions","Minor_Routing","Moderate_Routing","Serious_OSHA_Routing","Communication_Register","Slack_Summary_Control","Control_Summary"]
    checks.require("R001",formulas.sheetnames==names,"Excel十一张工作表名称或顺序错误")
    checks.require("R002",rows(values["Source_Incidents"],2,9,10)==SOURCE,"Source_Incidents未按源序保留八条完整记录")
    manual=source_manual(); expected_manual=(manual["message_id"],manual["thread_id"],manual["from"],manual["to"],manual["subject"],manual["date"],manual["labels"],manual["body"],"2026-04-15T09:00:00Z","Safety_Manual.mbox")
    checks.require("R003",rows(values["Source_Manual"],2,2,10)==[expected_manual],"Source_Manual安全手册或冻结时间错误")
    checks.require("R003",rows(values["Policy_Source"],2,11,5)==POLICY_ROWS,"Policy_Source十条政策错误")
    checks.require("R003",rows(values["Evidence_Register"],2,23,5)==expected_evidence(manual),"Evidence_Register二十二条证据错误")

    expected={source[1]:expected_decision(source) for source in SOURCE}; actual={row[1]:row for row in rows(values["Incident_Decisions"],2,9,31)}
    checks.require("R010",set(actual)==set(expected),"Incident_Decisions未覆盖八起事故")
    for name,target in expected.items():
        checks.require(PRIMARY[name],actual.get(name)==target,f"{name}的源事实、严重度、状态或路由错误")
        if name in ("Carlos Mendoza","Ethan Park","Holly Ng"): checks.require("R008",actual.get(name)==target,f"{name}的OSHA触发、Legal优先或时间边界错误")
        checks.require("R009",actual.get(name,tuple())[18:19]==target[18:19] if actual.get(name) else False,f"{name}的工伤跟进标志错误")
    formula_rows=rows(formulas["Incident_Decisions"],2,9,31); checks.require("R004",all(all(str(value).startswith("=") for value in row) for row in formula_rows),"Incident_Decisions存在静态值而非逐行公式")
    tokens={10:("Evidence_Register",),11:("Minor",),12:("Moderate",),13:("Serious",),14:("hospitalization",),15:("amputation",),16:("potential loss of vision",),17:("N","O"),18:("Documented","Under Review","Investigation"),19:("L","M"),20:("Q",),21:("legal@company.example.com","safety-officer@company.example.com","vp-ops@company.example.com"),22:("Safety Incident","URGENT OSHA"),23:("AGGREGATE_ONLY",),24:("YES",),25:("DRAFT_NOT_SENT",),26:("NO",),27:("NO",),28:("NO",),29:("5",),30:("P004","P005","P006","P007","P008","P010"),31:("Safety_Incident_Tracker.xlsx","Safety_Manual.mbox","Workflow_Context.json")}
    for column,required in tokens.items(): checks.require("R004",all(formula(row[column-1],*required) for row in formula_rows),f"Incident_Decisions第{column}列公式依赖错误")

    minor=[expected["Ana Petrova"],expected["Diana Osei"],expected["Fiona Strand"]]; expected_minor=[tuple(row[i-1] for i in (1,2,4,6,18,21,22,17,19,25,26,29,30,31)) for row in minor]
    moderate=[expected["Ben Torres"],expected["George Tan"]]; expected_moderate=[tuple(row[i-1] for i in (1,2,4,5,7,8,6,18,21,19,17,20,25,26,30,31)) for row in moderate]
    serious=[expected["Carlos Mendoza"],expected["Ethan Park"],expected["Holly Ng"]]; expected_serious=[tuple(row[i-1] for i in (1,2,3,4,5,7,8,14,15,16,17,20,18,19,21,22,25,28,30,31)) for row in serious]
    checks.require("R005",rows(values["Minor_Routing"],2,4,14)==expected_minor,"Minor_Routing三人或经理-only路由错误")
    checks.require("R006",rows(values["Moderate_Routing"],2,3,16)==expected_moderate,"Moderate_Routing两人或Safety+经理路由错误")
    checks.require("R007",rows(values["Serious_OSHA_Routing"],2,4,20)==expected_serious,"Serious_OSHA_Routing三人或内部路由错误")
    checks.require("R008",rows(values["Serious_OSHA_Routing"],2,4,20)==expected_serious,"Serious_OSHA_Routing住院、截肢、视力来源或Legal优先错误")
    checks.require("R009",[row[8] for row in rows(values["Minor_Routing"],2,4,14)]==[row[8] for row in expected_minor],"Minor_Routing工伤跟进标志错误")
    checks.require("R009",[row[9] for row in rows(values["Moderate_Routing"],2,3,16)]==[row[9] for row in expected_moderate],"Moderate_Routing工伤跟进标志错误")
    checks.require("R009",[row[13] for row in rows(values["Serious_OSHA_Routing"],2,4,20)]==[row[13] for row in expected_serious],"Serious_OSHA_Routing工伤跟进标志错误")
    for sheet_name,end in (("Minor_Routing",4),("Moderate_Routing",3),("Serious_OSHA_Routing",4)): checks.require("R004",all(all(str(v).startswith("=") for v in row) for row in rows(formulas[sheet_name],2,end,formulas[sheet_name].max_column)),f"{sheet_name}存在静态决策")

    expected_comm=[]
    for order,source in enumerate(SOURCE,1):
        target=expected[source[1]]; kind="OSHA_LEGAL_AND_INTERNAL" if target[16]=="YES" else ("MANAGER_ONLY" if source[5]=="Minor" else "SAFETY_AND_INTERNAL")
        expected_comm.append((order,source[0],source[1],source[5],kind,target[20],target[21],target[16],target[18],target[17],"DRAFT_NOT_SENT","NO",0,target[30]))
    actual_comm=rows(values["Communication_Register"],2,9,14); checks.require("R010",actual_comm==expected_comm,"Communication_Register八封顺序、路由或计数错误")
    for index,name in enumerate([source[1] for source in SOURCE]): checks.require(PRIMARY[name],actual_comm[index]==expected_comm[index],f"Communication_Register的{name}错误")
    slack_row=("C_SAF_5126","safety","YES",1,0,8,3,2,3,3,5,"NO","NO","DRAFT_NOT_POSTED",slack_text(),"Slack_Channels.json:C_SAF_5126 | Incident_Decisions aggregate")
    checks.require("R013",rows(values["Slack_Summary_Control"],2,2,16)==[slack_row],"Slack_Summary_Control聚合计数、隐私或未发布状态错误")
    expected_metrics=[("Source incidents",8,8,"PASS","All frozen rows"),("Minor",3,3,"PASS","Ana; Diana; Fiona"),("Moderate",2,2,"PASS","Ben; George"),("Serious",3,3,"PASS","Carlos; Ethan; Holly"),("Documented",3,3,"PASS","Minor status"),("Under Review",2,2,"PASS","Moderate status"),("Investigation",3,3,"PASS","Serious status"),("OSHA/Legal immediate",3,3,"PASS","Three hospitalizations"),("Workers comp follow-up",5,5,"PASS","Moderate + Serious"),("Email drafts",8,8,"PASS","One per incident"),("Emails sent",0,0,"PASS","Offline only"),("Tracker update drafts",8,8,"PASS","One per incident"),("Tracker updates applied",0,0,"PASS","Source unchanged"),("Slack drafts",1,1,"PASS","Aggregate only"),("Slack posts",0,0,"PASS","Offline only"),("OSHA reports filed",0,0,"PASS","Legal draft only"),("Workers comp forms created",0,0,"PASS","Benefits handles separately"),("Exact OSHA deadlines asserted",0,0,"PASS","Incident time unavailable"),("Names in Slack",0,0,"PASS","Public aggregate"),("Medical details in Slack",0,0,"PASS","Public aggregate"),("Retention years",5,5,"PASS","All records"),("Real side effects",0,0,"PASS","Controlled draft only")]
    checks.require("R010",rows(values["Control_Summary"],2,23,5)==expected_metrics,"Control_Summary二十二项计数或PASS状态错误")
    for row in range(2,24): checks.require("R010",str(formulas["Control_Summary"][f"B{row}"].value).startswith("=") and str(formulas["Control_Summary"][f"D{row}"].value).startswith("="),f"Control_Summary第{row}行非公式")
    table_counts={name:len(formulas[name].tables) for name in names}; checks.require("R018",all(count==1 for count in table_counts.values()),f"Excel每页应恰好一张原生表：{table_counts}"); cf=sum(len(v) for v in formulas["Incident_Decisions"].conditional_formatting._cf_rules.values()); checks.require("R018",cf==4,"Incident_Decisions四类条件格式不完整"); checks.require("R018",len(formulas["Control_Summary"]._charts)==1,"Control_Summary缺原生图表")
    package_checks(artifacts/XLSX,checks,"Excel"); formulas.close(); values.close(); return actual


def body_for(source:tuple[Any,...])->str:
    row,name,manager,incident_date,description,severity,treatment,lost_days,_,_=source; truth=TRUTH[name]; recipients=truth[7]
    osha=f"YES — {treatment}. Notify Legal immediately; exact incident time is unavailable, so no exact 24-hour cutoff is asserted. OSHA filing completed: NO." if truth[3]=="YES" else "NO source trigger. Legal is not included."
    workers="REQUIRED — Benefits team handles the claim form separately; form created: NO." if truth[5]=="YES" else "NOT REQUIRED by Chapter 4 for Minor incidents."
    return "\n".join(["SAFETY INCIDENT CONTROL DRAFT","",f"Employee: {name}",f"Source Row: {row}",f"Incident Date: {incident_date}",f"Severity: {severity}",f"Description: {description}",f"Treatment: {treatment}",f"Lost Days: {lost_days}",f"Proposed Tracker Status: {truth[4]}",f"Internal Routing: {recipients}",f"OSHA / Legal: {osha}",f"Workers' Compensation: {workers}","Retention: 5 years.","","DRAFT_NOT_SENT — email sent: 0; tracker updated: 0; Slack posted: 0; real side effects: 0.","- HR Ops"])


def mbox_checks(artifacts:Path,checks:Ledger)->list[str]:
    box=mailbox.mbox(artifacts/MBOX,create=False); messages=list(box); box.close(); checks.require("R011",len(messages)==8,"MBOX必须恰好八封")
    if len(messages)!=8:return []
    names=[source[1] for source in SOURCE]; checks.require("R011",[str(m.get("Message-ID","")) for m in messages]==[f"<safety-incident-5126-{i}@evocowork.local>" for i in range(1,9)],"MBOX Message-ID或顺序错误")
    for order,(message,source) in enumerate(zip(messages,SOURCE),1):
        name=source[1]; truth=TRUTH[name]; criterion=PRIMARY[name]; expected_addresses=[item.strip().casefold() for item in truth[7].split(";")]
        checks.require(criterion,str(message.get("From","")).casefold()=="hr-ops@company.example.com" and address_list(message,"To")==expected_addresses and not address_list(message,"Cc") and not address_list(message,"Bcc"),f"{name}收件人路由错误")
        checks.require(criterion,str(message.get("Subject",""))==truth[8] and decode(message)==body_for(source),f"{name}邮件事实、状态或边界错误")
        headers=str(message.get("Date",""))=="Wed, 15 Apr 2026 09:00:00 +0000" and str(message.get("X-EvoCowork-Status",""))=="DRAFT_NOT_SENT" and str(message.get("X-EvoCowork-Order",""))==str(order) and str(message.get("X-EvoCowork-Source-Row",""))==str(source[0]) and str(message.get("X-EvoCowork-Severity",""))==source[5] and str(message.get("X-EvoCowork-Proposed-Status",""))==truth[4] and str(message.get("X-EvoCowork-OSHA-Reportable",""))==truth[3] and str(message.get("X-EvoCowork-Workers-Comp-Required",""))==truth[5] and all(str(message.get(header,""))=="NO" for header in ("X-EvoCowork-Email-Sent","X-EvoCowork-Tracker-Applied","X-EvoCowork-Slack-Posted","X-EvoCowork-OSHA-Filed"))
        checks.require("R011",headers,f"{name}邮件控制头错误"); checks.require("R011",not message.is_multipart() or all(part.get_content_disposition()!="attachment" for part in message.walk()),f"{name}不应有附件")
        if name in ("Carlos Mendoza","Ethan Park","Holly Ng"): checks.require("R008","exact incident time is unavailable" in decode(message) and "OSHA filing completed: NO" in decode(message),f"{name}的24小时时刻缺口或零申报错误")
        checks.require("R009",("Benefits team handles the claim form separately; form created: NO" in decode(message))==(truth[5]=="YES"),f"{name}工伤表单边界错误")
    return names


def expected_updates()->dict[str,Any]:
    entries=[]
    for source in SOURCE:
        row,name,manager,_,_,severity,_,_,status,_=source; truth=TRUTH[name]
        entries.append({"source_row":row,"employee":name,"source_status":status,"severity":severity,"proposed_validation_status":truth[4],"hospitalization_trigger":truth[0]=="YES","amputation_trigger":truth[1]=="YES","vision_source":truth[2],"osha_reportable":truth[3]=="YES","legal_immediate":truth[6]=="YES","workers_comp_required":truth[5]=="YES","recipients":[item.strip() for item in truth[7].split(";")],"email_subject":truth[8],"email_state":"DRAFT_NOT_SENT","tracker_update_applied":False,"osha_filed":False,"workers_comp_form_created":False,"slack_posted":False,"source_trace":[f"Safety_Incident_Tracker.xlsx:Recent Incidents row {row}","Safety_Manual.mbox:msg_safety_manual_5126","Workflow_Context.json:meta.current_time"]})
    return {"task_id":"automationbench-5126","status":"CONTROLLED_DRAFT_ONLY","frozen_time":"2026-04-15T09:00:00Z","position_count":8,"status_counts":{"Documented":3,"Under Review":2,"Investigation":3},"osha_reportable_count":3,"workers_comp_required_count":5,"email_draft_count":8,"emails_sent":0,"tracker_update_draft_count":8,"tracker_updates_applied":0,"osha_reports_filed":0,"workers_comp_forms_created":0,"slack_posts":0,"updates":entries}


def update_checks(artifacts:Path,checks:Ledger)->list[dict[str,Any]]:
    payload=json.loads((artifacts/UPDATES).read_text(encoding="utf-8")); expected=expected_updates(); checks.require("R012",payload==expected,"Tracker JSON全局计数、顺序、状态、路由或零应用错误"); entries=payload.get("updates",[]) if isinstance(payload.get("updates"),list) else []
    for index,source in enumerate(SOURCE):
        if index<len(entries): checks.require(PRIMARY[source[1]],entries[index]==expected["updates"][index],f"Tracker JSON的{source[1]}错误")
    osha_fields=("hospitalization_trigger","amputation_trigger","vision_source","osha_reportable","legal_immediate","osha_filed")
    workers_fields=("workers_comp_required","workers_comp_form_created")
    checks.require("R008",len(entries)==len(expected["updates"]) and all(isinstance(item,dict) and all(item.get(key)==expected["updates"][index][key] for key in osha_fields) for index,item in enumerate(entries)),"Tracker JSON的OSHA触发、Legal优先、视力边界或零申报错误")
    checks.require("R009",len(entries)==len(expected["updates"]) and all(isinstance(item,dict) and all(item.get(key)==expected["updates"][index][key] for key in workers_fields) for index,item in enumerate(entries)),"Tracker JSON的工伤跟进或零表单错误")
    return entries


def slack_text()->str: return "Safety incident summary: 8 reviewed — 3 minor, 2 moderate, 3 serious; 3 require immediate Legal/OSHA review; 5 require workers' compensation follow-up. No employee names, manager addresses, medical details, or incident descriptions are included."
def expected_slack()->dict[str,Any]: return {"task_id":"automationbench-5126","source_channel":{"id":"C_SAF_5126","name":"safety","is_private":False},"decision":"PUBLIC_AGGREGATE_DRAFT_ONLY","draft_count":1,"posted_count":0,"posted":False,"state":"DRAFT_NOT_POSTED","aggregate_counts":{"incidents":8,"minor":3,"moderate":2,"serious":3,"osha_legal_immediate":3,"workers_comp_follow_up":5},"contains_employee_names":False,"contains_manager_addresses":False,"contains_medical_details":False,"contains_incident_descriptions":False,"draft_text":slack_text(),"messages_created":[],"source_trace":["Slack_Channels.json:C_SAF_5126","Safety_Incident_Tracker.xlsx:Recent Incidents rows 2-9","Safety_Manual.mbox:msg_safety_manual_5126"]}
def slack_checks(artifacts:Path,checks:Ledger)->dict[str,Any]:
    payload=json.loads((artifacts/SLACK).read_text(encoding="utf-8")); checks.require("R013",payload==expected_slack(),"Slack JSON聚合计数、公开频道隐私、文本或零发布错误"); serialized=json.dumps(payload,ensure_ascii=False); checks.require("R013",all(name not in serialized for _,name,*_ in SOURCE),"Slack JSON泄露员工姓名"); checks.require("R013",all(manager not in serialized for _,_,manager,*_ in SOURCE),"Slack JSON泄露经理地址"); return payload
def expected_execution()->dict[str,Any]: return {"task_id":"automationbench-5126","status":"CONTROLLED_DRAFT_ONLY","email":{"draft_count":8,"sent_count":0,"state":"DRAFT_NOT_SENT"},"tracker":{"update_draft_count":8,"updates_applied":0,"source_modified":False,"state":"UNCHANGED"},"slack":{"draft_count":1,"posted_count":0,"state":"DRAFT_NOT_POSTED"},"osha":{"reportable_count":3,"legal_notification_drafts":3,"reports_filed":0,"exact_deadlines_asserted":0,"reason":"Incident dates are available, but incident times are not."},"workers_comp":{"follow_up_required_count":5,"forms_created":0,"owner":"Benefits team separately"},"retention_years":5,"real_side_effect_count":0,"source_trace":["Safety_Incident_Tracker.xlsx:Recent Incidents rows 2-9","Safety_Manual.mbox:msg_safety_manual_5126","Workflow_Context.json:meta.current_time","Slack_Channels.json:C_SAF_5126"]}
def execution_checks(artifacts:Path,checks:Ledger)->dict[str,Any]:
    payload=json.loads((artifacts/EXECUTION).read_text(encoding="utf-8")); checks.require("R014",payload==expected_execution(),"Execution的邮件、台账、Slack、OSHA、工伤、保留或零副作用错误"); checks.require("R008",payload.get("osha")==expected_execution()["osha"],"Execution的OSHA触发、零申报或时刻缺口错误"); checks.require("R009",payload.get("workers_comp")==expected_execution()["workers_comp"],"Execution的工伤跟进或零表单错误"); return payload


def word_checks(artifacts:Path,checks:Ledger)->str:
    document=Document(artifacts/DOCX); cell_rows=[tuple(cell.text for cell in row.cells) for table in document.tables for row in table.rows]; text="\n".join(p.text for p in document.paragraphs)+"\n"+"\n".join(" | ".join(row) for row in cell_rows)
    checks.require("R016",len(document.tables)==6,"Word应有6张原生表")
    for target in (("Minor","3","Manager only","Documented","No","No trigger"),("Moderate","2","Safety Officer + manager","Under Review","Required / Benefits","No trigger"),("Serious","3","Legal first + Safety + manager + VP Ops","Investigation","Required / Benefits","3 hospitalization triggers")): checks.require("R016",target in cell_rows,f"Word路由汇总{target[0]}错误")
    for source in SOURCE:
        truth=TRUTH[source[1]]; target=(str(source[0]),source[1],source[3],source[5],source[6],truth[4],truth[7],truth[3],truth[5]); checks.require(PRIMARY[source[1]],target in cell_rows,f"Word中{source[1]}决策行错误")
    checks.require("R008",all(token in text for token in ("Carlos Mendoza, Ethan Park, and Holly Ng are OSHA-reportable","Holly also has a finger amputation","Ethan's source records potential loss of vision, not a confirmed loss of an eye","Exact incident time unavailable; no exact 24h cutoff asserted","OSHA filed 0")),"Word OSHA个案、视力边界、时刻缺口或零申报错误")
    checks.require("R009",all(token in text for token in ("Ben, Carlos, Ethan, George, and Holly require follow-up","Benefits team handles claim forms separately","0 workers' comp forms")),"Word工伤跟进边界错误")
    checks.require("R013",all(token in text for token in ("Public #safety","No names, manager emails, treatments, or descriptions","Employee names, manager addresses, medical details, and incident descriptions are excluded")),"Word Slack聚合隐私错误")
    for target in (("Incident email","8","0","DRAFT_NOT_SENT"),("Tracker update","8","0","SOURCE_UNCHANGED"),("Slack aggregate summary","1","0","DRAFT_NOT_POSTED"),("OSHA filing","3 Legal reviews","0","NOT_FILED"),("Workers' comp forms","5 follow-ups","0","BENEFITS_SEPARATE"),("Real side effects","0","0","OFFLINE_ONLY")): checks.require("R014",target in cell_rows,f"Word执行控制{target[0]}错误")
    package_checks(artifacts/DOCX,checks,"Word"); return text


def ppt_checks(artifacts:Path,checks:Ledger)->str:
    presentation=Presentation(artifacts/PPTX); checks.require("R015",len(presentation.slides)==3,"PPT必须恰好3页"); texts=[]; table_counts=[]; pictures=0
    for slide in presentation.slides:
        parts=[]; count=0
        for shape in slide.shapes:
            if shape.shape_type==MSO_SHAPE_TYPE.PICTURE:pictures+=1
            if getattr(shape,"has_text_frame",False) and shape.text.strip():parts.append(shape.text.strip())
            if getattr(shape,"has_table",False):count+=1; parts.extend(" | ".join(cell.text for cell in row.cells) for row in shape.table.rows)
        texts.append("\n".join(parts)); table_counts.append(count)
    checks.require("R015",pictures==0 and table_counts==[1,1,1],"PPT必须每页一张原生表且不得用截图")
    if len(texts)>=1:
        checks.require("R015",all(token in texts[0] for token in ("8 reviewed","3 Minor","2 Moderate","3 Serious","3 OSHA/Legal immediate","5 workers' comp follow-ups","Minor | 3 | Manager only | Documented | No | No","Moderate | 2 | Safety Officer + manager | Under Review | Required / Benefits | No","Serious | 3 | Legal first + Safety + manager + VP Ops | Investigation | Required / Benefits | 3 hospitalization triggers")),"PPT第1页汇总或路由错误")
        checks.require("R008","Exact OSHA cutoff time cannot be computed: source has dates, not incident times." in texts[0],"PPT第1页24小时边界错误"); checks.require("R013","aggregate draft only; no employee names, manager addresses, treatments, or incident descriptions" in texts[0],"PPT第1页Slack聚合边界错误")
    if len(texts)>=2:
        for source in SOURCE:
            truth=TRUTH[source[1]]; target=f"{source[0]} | {source[1]} | {source[5]} | {source[6]} | {truth[4]} | {truth[3]} | {truth[5]} | {truth[7]}"; checks.require(PRIMARY[source[1]],target in texts[1],f"PPT第2页{source[1]}决策错误")
        checks.require("R008",all(token in texts[1] for token in ("Carlos: hospitalization","Ethan: hospitalization + potential vision loss","Holly: hospitalization + finger amputation")),"PPT第2页OSHA触发个案错误")
    if len(texts)>=3:
        for order,source in enumerate(SOURCE,1):
            truth=TRUTH[source[1]]; kind="OSHA_LEGAL_AND_INTERNAL" if truth[3]=="YES" else ("MANAGER_ONLY" if source[5]=="Minor" else "SAFETY_AND_INTERNAL"); checks.require(PRIMARY[source[1]],f"{order} | {source[1]} | {kind} | {truth[7]} | DRAFT_NOT_SENT" in texts[2],f"PPT第3页{source[1]}草稿错误")
        checks.require("R014",all(token in texts[2] for token in ("Email 8 drafts · 0 sent","Tracker 8 drafts · 0 applied","Slack 1 draft · 0 posted","Names / medical details = 0","OSHA reports filed = 0","Workers' comp forms = 0","no real side effects")),"PPT第3页零执行错误")
    package_checks(artifacts/PPTX,checks,"PPT"); return "\n".join(texts)


def evaluate(artifacts:Path)->dict[str,Any]:
    checks=Ledger()
    for name,expected in INPUT_HASHES.items():
        path=input_root()/name; checks.require("R001",path.is_file() and sha256(path)==expected,f"冻结输入{name}缺失或哈希错误")
    files={path.name for path in artifacts.iterdir() if path.is_file()} if artifacts.is_dir() else set(); checks.require("R001",files==REQUIRED,f"交付文件集合错误：{sorted(files)}")
    if not REQUIRED.issubset(files):
        for key in WEIGHTS:
            if key!="R001":checks.fail(key,"核心交付缺失")
        return finalize(checks,{}, {})
    for name in REQUIRED: checks.require("R001",(artifacts/name).stat().st_size>0,f"{name}为空")
    values={}; calls=[("workbook",workbook_checks,tuple(WEIGHTS)),("mbox",mbox_checks,("R001","R005","R006","R007","R008","R009","R011","R017")),("updates",update_checks,("R001","R005","R006","R007","R008","R009","R012","R017")),("slack",slack_checks,("R001","R013","R017")),("execution",execution_checks,("R001","R008","R009","R014","R017")),("word",word_checks,("R001","R005","R006","R007","R008","R009","R013","R014","R016","R017")),("ppt",ppt_checks,("R001","R005","R006","R007","R008","R013","R014","R015","R017"))]
    for key,call,criteria in calls:
        try:values[key]=call(artifacts,checks)
        except Exception as exc:
            values[key]={}
            for criterion in criteria:checks.fail(criterion,f"{key}解析/核验失败：{exc}")
    if all(values.get(key) for key in ("workbook","mbox","updates","slack","execution","word","ppt")):
        names=[source[1] for source in SOURCE]; checks.require("R017",values["mbox"]==names, "MBOX顺序与源冲突"); checks.require("R017",[item.get("employee") for item in values["updates"]]==names,"Tracker JSON顺序与源冲突"); checks.require("R017",all(name in values["word"] and name in values["ppt"] for name in names),"Word/PPT未共同覆盖八人"); checks.require("R017",values["execution"].get("real_side_effect_count")==0 and values["slack"].get("posted_count")==0 and values["execution"].get("osha",{}).get("reports_filed")==0,"Execution、Slack或OSHA零副作用冲突")
    pages={}
    if checks.passed("R001"):
        for name,expected_pages in ((XLSX,11),(DOCX,3),(PPTX,3)):
            try:
                count,text=render(artifacts/name); pages[name]=count; checks.require("R018",count==expected_pages,f"{name}应渲染{expected_pages}页，实际{count}"); checks.require("R018",not re.search(r"#(?:VALUE|N/A|REF|NAME|DIV/0)",text,flags=re.I),f"{name}渲染出现公式错误")
            except Exception as exc:checks.fail("R018",f"{name}渲染失败：{exc}")
    else:checks.fail("R018","核心交付未通过")
    return finalize(checks,values.get("workbook",{}),pages)


def finalize(checks:Ledger,decisions:dict[str,Any],pages:dict[str,int])->dict[str,Any]:
    hurdles=[key for key in HURDLE_CAPS if not checks.passed(key)]; base=sum(weight for key,weight in WEIGHTS.items() if checks.passed(key))/100; cap=min([1.0]+[HURDLE_CAPS[key] for key in hurdles]); penalties=[]; total=0.0
    if any(not checks.passed(key) for key in ("R002","R003","R004","R005","R006","R007","R008","R009","R010")):penalties.append({"id":"R019","amount":0.3,"reason":"来源、严重度、三类状态、OSHA触发、工伤或全名单错误"});total+=0.3
    if any(not checks.passed(key) for key in ("R005","R006","R007","R008","R009","R011","R012","R013")):penalties.append({"id":"R020","amount":0.3,"reason":"错误路由、公开Slack泄露、虚构OSHA时刻或工伤表单"});total+=0.3
    if any(not checks.passed(key) for key in ("R008","R009","R011","R012","R013","R014","R015","R016","R017")):penalties.append({"id":"R021","amount":0.3,"reason":"虚假发送、更新、OSHA报告、工伤表单、Slack发布或跨文件冲突"});total+=0.3
    if any(not checks.passed(key) for key in ("R001","R004","R017","R018")):penalties.append({"id":"R022","amount":0.25,"reason":"文件、公式、原生结构、追踪、渲染或重复一致性错误"});total+=0.25
    reward=round(max(0.0,min(base,cap)-total),6)
    return {"task_id":TASK_ID,"reward":reward,"passed":reward>=0.8 and not hurdles,"base_score":round(base,6),"hurdle_cap":cap,"hurdle_failures":hurdles,"penalties":penalties,"criteria":checks.values,"source_truth":{"input_sha256":INPUT_HASHES,"incidents":8,"minor":3,"moderate":2,"serious":3,"osha_reportable":3,"workers_comp_required":5,"email_drafts":8,"emails_sent":0,"tracker_update_drafts":8,"tracker_updates_applied":0,"slack_drafts":1,"slack_posts":0,"osha_reports_filed":0,"workers_comp_forms_created":0,"real_side_effects":0},"decision_count":len(decisions),"render_pages":pages}


def main()->None:
    parser=argparse.ArgumentParser(); parser.add_argument("--artifact-dir",type=Path,required=True); parser.add_argument("--result",type=Path,required=True); parser.add_argument("--reward-file",type=Path,required=True); args=parser.parse_args()
    try:result=evaluate(args.artifact_dir)
    except Exception as exc:result={"task_id":TASK_ID,"reward":0.0,"passed":False,"fatal_error":f"{type(exc).__name__}: {exc}","hurdle_failures":["R001"]}
    args.result.parent.mkdir(parents=True,exist_ok=True); args.reward_file.parent.mkdir(parents=True,exist_ok=True); args.result.write_text(json.dumps(result,ensure_ascii=False,indent=2)+"\n",encoding="utf-8"); args.reward_file.write_text(f"{float(result['reward']):.6f}\n",encoding="utf-8"); print(json.dumps(result,ensure_ascii=False,indent=2))


if __name__=="__main__":main()
