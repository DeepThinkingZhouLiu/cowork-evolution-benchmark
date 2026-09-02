#!/usr/bin/env python3
"""向 AutomationBench 5117 Golden 逐项注入本题专属错误，验证 Judge 拒绝坏答案。"""

from __future__ import annotations

import argparse
import importlib.util
import json
import mailbox
import shutil
import sys
import tempfile
import zipfile
from email.generator import BytesGenerator
from email.message import EmailMessage
from email.policy import SMTP
from io import BytesIO
from pathlib import Path
from typing import Callable

from docx import Document
from lxml import etree
from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches


ROOT=Path(__file__).resolve().parents[1];GOLDEN=ROOT/"solution"/"Golden"
XLSX="Grievance_Routing_Audit.xlsx";DOCX="Grievance_Routing_Control_Brief.docx";PPTX="Grievance_Routing_Review.pptx";MBOX="Grievance_Notification_Drafts.mbox";UPDATES="Grievance_Tracker_Update_Drafts.json";SLACK="Grievance_Slack_Control.json";EXECUTION="Grievance_Execution_Control.json"
FILES=(XLSX,DOCX,PPTX,MBOX,UPDATES,SLACK,EXECUTION)
CASES=["GR-2026-041","GR-2026-042","GR-2026-043","GR-2026-044","GR-2026-045","GR-2026-046","GR-2026-047","GR-2026-048"]
EMPLOYEES=["Anna Svensson","Ben Torres","Carmen Obi","David Holm","Elise Braun","Frank Yoshida","Grace Petrov","Hugo Strand"]
MANAGERS=["Robert Kim","Lisa Wang","Derek Huang","Mira Patel"]
DESCRIPTIONS=["Noise complaint about adjacent team area","Overtime not being properly tracked","Feels excluded from team activities","Exposed wiring in server room","Manager Robert Kim making inappropriate comments","Disagrees with new remote work policy application","Passed over for promotion due to gender","Persistent bullying from peer"]
ROUTE_CRITERIA=["R005","R006","R008","R009","R007","R006","R008","R006"]
Mutator=Callable[[Path],None];Case=tuple[str,Mutator,str]


def load_judge():
    spec=importlib.util.spec_from_file_location("grievance_5117_judge",ROOT/"tests"/"judge.py")
    if spec is None or spec.loader is None:raise RuntimeError("无法载入本题 Judge")
    module=importlib.util.module_from_spec(spec);sys.modules[spec.name]=module;spec.loader.exec_module(module);module.render=lambda path:({".xlsx":9,".docx":3,".pptx":3}[path.suffix],"render-ok");return module


def rewrite_zip(path:Path,transform:Callable[[str,bytes],bytes|None])->None:
    temporary=path.with_suffix(path.suffix+".negative")
    with zipfile.ZipFile(path,"r") as source,zipfile.ZipFile(temporary,"w",zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data=transform(info.filename,source.read(info.filename))
            if data is not None:target.writestr(info,data)
    temporary.replace(path)


def sheet_xml_path(archive:zipfile.ZipFile,sheet_name:str)->str:
    workbook=etree.fromstring(archive.read("xl/workbook.xml"));relationships=etree.fromstring(archive.read("xl/_rels/workbook.xml.rels"));namespaces={"m":"http://schemas.openxmlformats.org/spreadsheetml/2006/main","r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships","pr":"http://schemas.openxmlformats.org/package/2006/relationships"}
    relationship_id=workbook.xpath(f'//m:sheet[@name="{sheet_name}"]/@r:id',namespaces=namespaces)[0];target=relationships.xpath(f'//pr:Relationship[@Id="{relationship_id}"]/@Target',namespaces=namespaces)[0]
    return target.lstrip("/") if target.startswith("/") else "xl/"+target.lstrip("/")


def edit_cell(path:Path,sheet_name:str,coordinate:str,*,value=None,remove_formula:bool=False)->None:
    with zipfile.ZipFile(path,"r") as archive:target_sheet=sheet_xml_path(archive,sheet_name)
    changed=False;namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    def transform(name:str,data:bytes)->bytes:
        nonlocal changed
        if name!=target_sheet:return data
        root=etree.fromstring(data);cells=root.xpath(f'//m:c[@r="{coordinate}"]',namespaces={"m":namespace})
        if not cells:raise ValueError(f"{sheet_name}!{coordinate} 不存在")
        cell=cells[0]
        if remove_formula:
            formula=cell.find(f"{{{namespace}}}f")
            if formula is None:raise ValueError(f"{sheet_name}!{coordinate} 没有公式")
            cell.remove(formula)
        if value is not None:
            inline=cell.find(f"{{{namespace}}}is")
            if inline is not None:cell.remove(inline)
            node=cell.find(f"{{{namespace}}}v")
            if node is None:node=etree.SubElement(cell,f"{{{namespace}}}v")
            if isinstance(value,bool):node.text="1" if value else "0";cell.set("t","b")
            elif isinstance(value,(int,float)):node.text=str(value);cell.attrib.pop("t",None)
            else:node.text=str(value);cell.set("t","str")
        changed=True;return etree.tostring(root,xml_declaration=True,encoding="UTF-8",standalone=True)
    rewrite_zip(path,transform)
    if not changed:raise ValueError("目标单元格未修改")


def openpyxl_mutation(path:Path,mutator:Callable)->None:
    workbook=load_workbook(path);mutator(workbook);workbook.save(path)
def mutate_json(path:Path,mutator:Callable[[dict],None])->None:
    payload=json.loads(path.read_text(encoding="utf-8"));mutator(payload);path.write_text(json.dumps(payload,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
def decode_message(message)->str:
    if message.is_multipart():return "\n".join((part.get_payload(decode=True) or b"").decode(part.get_content_charset() or "utf-8",errors="replace") for part in message.walk() if part.get_content_type()=="text/plain").strip()
    return (message.get_payload(decode=True) or b"").decode(message.get_content_charset() or "utf-8",errors="replace").strip()
def mbox_states(path:Path)->list[dict]:
    box=mailbox.mbox(path,create=False);messages=list(box);box.close();result=[]
    for message in messages:result.append({"from":str(message.get("From","")),"to":str(message.get("To","")),"cc":str(message.get("Cc","")),"bcc":str(message.get("Bcc","")),"subject":str(message.get("Subject","")),"date":str(message.get("Date","")),"message_id":str(message.get("Message-ID","")),"status":str(message.get("X-EvoCowork-Status","")),"route":str(message.get("X-EvoCowork-Route","")),"order":str(message.get("X-EvoCowork-Order","")),"confidentiality":str(message.get("X-EvoCowork-Confidentiality","")),"body":decode_message(message)})
    return result
def write_mbox_states(path:Path,states:list[dict])->None:
    chunks=[]
    for state in states:
        message=EmailMessage(policy=SMTP)
        for key,header in (("from","From"),("to","To"),("cc","Cc"),("bcc","Bcc"),("subject","Subject"),("date","Date"),("message_id","Message-ID"),("status","X-EvoCowork-Status"),("route","X-EvoCowork-Route"),("order","X-EvoCowork-Order"),("confidentiality","X-EvoCowork-Confidentiality")):
            if state.get(key):message[header]=state[key]
        message.set_content(state["body"].rstrip()+"\n",charset="utf-8");buffer=BytesIO();BytesGenerator(buffer,policy=SMTP,mangle_from_=True,maxheaderlen=0).flatten(message);chunks.append(b"From MAILER-DAEMON Wed Apr 15 09:00:01 2026\n"+buffer.getvalue().replace(b"\r\n",b"\n")+b"\n")
    path.write_bytes(b"".join(chunks))
def mutate_message(path:Path,index:int,mutator:Callable[[dict],None])->None:
    states=mbox_states(path);mutator(states[index]);write_mbox_states(path,states)
def reorder_messages(path:Path)->None:
    states=mbox_states(path);states[0],states[1]=states[1],states[0];write_mbox_states(path,states)


def replace_docx_text(path:Path,old:str,new:str,*,row_token:str|None=None)->None:
    document=Document(path);changed=False
    if row_token is None:
        for paragraph in document.paragraphs:
            if old in paragraph.text:paragraph.text=paragraph.text.replace(old,new,1);changed=True;break
    if not changed:
        for table in document.tables:
            for row in table.rows:
                if row_token and row_token not in " | ".join(cell.text for cell in row.cells):continue
                for cell in row.cells:
                    if old in cell.text:cell.text=cell.text.replace(old,new,1);changed=True;break
                if changed:break
            if changed:break
    if not changed:raise ValueError(f"Word 找不到 {old!r}")
    document.save(path)
def inject_docx_text(path:Path,text:str)->None:
    document=Document(path);document.add_paragraph(text);document.save(path)
def replace_pptx_text(path:Path,old:str,new:str)->None:
    presentation=Presentation(path);changed=False
    for slide in presentation.slides:
        for shape in slide.shapes:
            if getattr(shape,"has_text_frame",False) and old in shape.text:shape.text=shape.text.replace(old,new,1);changed=True;break
            if getattr(shape,"has_table",False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        if old in cell.text:cell.text=cell.text.replace(old,new,1);changed=True;break
                    if changed:break
            if changed:break
        if changed:break
    if not changed:raise ValueError(f"PPT 找不到 {old!r}")
    presentation.save(path)
def edit_pptx_table_cell(path:Path,slide_index:int,row_index:int,column_index:int,value:str)->None:
    presentation=Presentation(path);tables=[shape.table for shape in presentation.slides[slide_index].shapes if getattr(shape,"has_table",False)]
    if len(tables)!=1:raise ValueError("PPT 目标页原生表数不是1")
    tables[0].cell(row_index,column_index).text=value;presentation.save(path)
def edit_pptx_shape_text(path:Path,slide_index:int,shape_index:int,value:str)->None:
    presentation=Presentation(path);shape=presentation.slides[slide_index].shapes[shape_index]
    if not getattr(shape,"has_text_frame",False):raise ValueError("PPT 目标形状没有文本")
    shape.text=value;presentation.save(path)
def inject_pptx_text(path:Path,text:str)->None:
    presentation=Presentation(path);shape=presentation.slides[0].shapes.add_textbox(Inches(0.1),Inches(7.1),Inches(5),Inches(0.2));shape.text=text;presentation.save(path)
def drop_ppt_slide(path:Path)->None:
    presentation=Presentation(path);presentation.slides._sldIdLst.remove(presentation.slides._sldIdLst[-1]);presentation.save(path)
def remove_path(root:Path,name:str)->None:(root/name).unlink()
def corrupt_path(root:Path,name:str)->None:(root/name).write_bytes(b"not-a-valid-artifact")


def build_cases()->list[Case]:
    cases=[]
    for name in FILES:cases.append((f"missing-{name}",lambda root,n=name:remove_path(root,n),"R001"))
    cases.append(("extra-file",lambda root:(root/"unexpected.txt").write_text("unexpected",encoding="utf-8"),"R001"))
    for name in FILES:cases.append((f"corrupt-{name}",lambda root,n=name:corrupt_path(root,n),"R001" if name in (XLSX,DOCX,PPTX) else {MBOX:"R010",UPDATES:"R011",SLACK:"R012",EXECUTION:"R013"}[name]))

    source_columns=[("source-row","A",99),("case-id","B","GR-WRONG"),("employee","C","Wrong Person"),("manager","D","Wrong Manager"),("manager-email","E","wrong@company.example.com"),("category","F","Wrong Category"),("severity","G","Level 9"),("description","H","Wrong description"),("status","I","Closed"),("notes","J","Wrong note")]
    for row,case_id in enumerate(CASES,2):
        for field,column,value in source_columns:cases.append((f"source-{case_id}-{field}",lambda root,r=row,c=column,v=value:edit_cell(root/XLSX,"Restricted_Source",f"{c}{r}",value=v),"R002"))

    policy_edits=[("handbook-id","A2","wrong-id"),("handbook-from","B2","wrong@company.example.com"),("handbook-to","C2","wrong@company.example.com"),("handbook-subject","D2","Wrong handbook"),("handbook-date","E2","Thu, 16 Apr 2026 09:00:00 +0000"),("handbook-body","F2","Wrong procedures"),("handbook-trace","G2","wrong-source"),("override-id","A3","wrong-id"),("override-from","B3","wrong@company.example.com"),("override-subject","D3","Wrong override"),("override-body","F3","No discrimination override"),("override-trace","G3","wrong-source")]
    for name,cell,value in policy_edits:cases.append((name,lambda root,c=cell,v=value:edit_cell(root/XLSX,"Policy_Source",c,value=v),"R003"))
    for row in range(2,14):cases.append((f"evidence-row-{row}",lambda root,r=row:edit_cell(root/XLSX,"Evidence_Register",f"D{r}",value="Wrong evidence"),"R003"))

    for row in range(2,10):
        for column in "ABCDEFGHIJ":cases.append((f"formula-source-{column}{row}",lambda root,r=row,c=column:edit_cell(root/XLSX,"Routing_Decisions",f"{c}{r}",remove_formula=True),"R004"))
        for column in "KLMNOPQRS":cases.append((f"formula-route-{column}{row}",lambda root,r=row,c=column:edit_cell(root/XLSX,"Routing_Decisions",f"{c}{r}",remove_formula=True),"R004"))
    for sheet,start,end in (("Manager_Queue",2,2),("HRBP_Queue",2,5),("Legal_HR_Director_Queue",2,4)):
        for row in range(start,end+1):
            for column in "BCDEFGHI":cases.append((f"formula-{sheet}-{column}{row}",lambda root,s=sheet,r=row,c=column:edit_cell(root/XLSX,s,f"{c}{r}",remove_formula=True),"R004"))
    for row in range(2,5):cases.append((f"formula-leadership-B{row}",lambda root,r=row:edit_cell(root/XLSX,"Leadership_Summary",f"B{r}",remove_formula=True),"R004"))
    for row in range(2,28):cases.append((f"formula-control-B{row}",lambda root,r=row:edit_cell(root/XLSX,"Control_Summary",f"B{r}",remove_formula=True),"R004"))

    route_edits=[
        (2,"M","HRBP"),(2,"N","HRBP"),(2,"O","hr-bp@company.example.com"),(2,"Q","Routed to HRBP"),
        (3,"M","MANAGER"),(3,"O","lisa.wang@company.example.com"),(3,"Q","Routed to Lisa Wang"),
        (4,"K","NO"),(4,"M","MANAGER"),(4,"O","derek.huang@company.example.com"),(4,"P",""),(4,"R","Handbook Section 4 Level 1"),
        (5,"M","HRBP"),(5,"O","hr-bp@company.example.com"),(5,"P",""),(5,"R","Handbook Section 4 Level 2"),
        (6,"L","NO"),(6,"M","MANAGER"),(6,"N","Robert Kim"),(6,"O","robert.kim@company.example.com"),(6,"Q","Routed to Robert Kim"),(6,"R","Handbook Section 4 Level 1"),
        (7,"M","MANAGER"),(7,"O","lisa.wang@company.example.com"),(7,"Q","Routed to Lisa Wang"),
        (8,"K","NO"),(8,"M","HRBP"),(8,"O","hr-bp@company.example.com"),(8,"P",""),(8,"R","Handbook Section 4 Level 2"),
        (9,"M","MANAGER"),(9,"O","mira.patel@company.example.com"),(9,"Q","Routed to Mira Patel"),
    ]
    row_criterion={2:"R005",3:"R006",4:"R008",5:"R009",6:"R007",7:"R006",8:"R008",9:"R006"}
    for row,column,value in route_edits:cases.append((f"decision-{CASES[row-2]}-{column}",lambda root,r=row,c=column,v=value:edit_cell(root/XLSX,"Routing_Decisions",f"{c}{r}",value=v),row_criterion[row]))
    for row in range(2,10):cases.append((f"decision-slack-allowed-{row}",lambda root,r=row:edit_cell(root/XLSX,"Routing_Decisions",f"T{r}",value="YES"),"R012"))
    for row in range(2,10):cases.append((f"decision-confidentiality-{row}",lambda root,r=row:edit_cell(root/XLSX,"Routing_Decisions",f"U{r}",value="PUBLIC"),row_criterion[row]))

    queue_edits=[("manager-case","Manager_Queue","C2","GR-2026-042"),("manager-to","Manager_Queue","F2","hr-bp@company.example.com"),("manager-state","Manager_Queue","J2","SENT"),("hrbp-first","HRBP_Queue","C2","GR-2026-041"),("hrbp-route","HRBP_Queue","D3","MANAGER"),("hrbp-to","HRBP_Queue","F4","wrong@company.example.com"),("hrbp-state","HRBP_Queue","J5","SENT"),("legal-first","Legal_HR_Director_Queue","C2","GR-2026-044"),("legal-cc","Legal_HR_Director_Queue","G3",""),("legal-state","Legal_HR_Director_Queue","J4","SENT")]
    for name,sheet,cell,value in queue_edits:cases.append((f"queue-{name}",lambda root,s=sheet,c=cell,v=value:edit_cell(root/XLSX,s,c,value=v),"R010"))
    leadership_edits=[("manager-count","B2",2),("manager-cases","C2","GR-2026-041, GR-2026-042"),("hrbp-count","B3",3),("hrbp-cases","C3","GR-2026-042"),("legal-count","B4",4),("legal-cases","C4","GR-2026-043"),("confidentiality","G3","PUBLIC")]
    for name,cell,value in leadership_edits:cases.append((f"leadership-{name}",lambda root,c=cell,v=value:edit_cell(root/XLSX,"Leadership_Summary",c,value=v),"R016" if name!="confidentiality" else "R012"))
    for name in list(EMPLOYEES)[:4]:cases.append((f"leadership-name-{name.replace(' ','-').lower()}",lambda root,n=name:edit_cell(root/XLSX,"Leadership_Summary","G2",value=n),"R012"))
    summary_edits=[("cases","B2",7),("level1","B3",2),("level2","B4",5),("level3","B5",2),("overrides","B6",1),("manager-exceptions","B7",0),("manager-routes","B8",2),("hrbp-routes","B9",3),("legal-routes","B10",4),("email-drafts","B11",7),("emails-sent","B12",1),("updates","B13",7),("applied","B14",1),("slack-drafts","B15",1),("slack-posts","B16",1),("source-modified","B17",1),("side-effects","B18",1),("first-case","B19","GR-2026-042"),("last-case","B20","GR-2026-047"),("manager-queue","B21",2),("hrbp-queue","B22",3),("legal-queue","B23",4),("confidentiality","B24","PUBLIC"),("slack-details","B25",1),("policy-messages","B26",1),("evidence","B27",11)]
    for name,cell,value in summary_edits:cases.append((f"summary-{name}",lambda root,c=cell,v=value:edit_cell(root/XLSX,"Control_Summary",c,value=v),"R004"))
    cases.extend([("xlsx-native-table-removed",lambda root:openpyxl_mutation(root/XLSX,lambda wb:wb["Restricted_Source"].tables.clear()),"R017"),("xlsx-native-chart-removed",lambda root:openpyxl_mutation(root/XLSX,lambda wb:setattr(wb["Leadership_Summary"],"_charts",[])),"R017"),("xlsx-conditional-formatting-removed",lambda root:openpyxl_mutation(root/XLSX,lambda wb:setattr(wb["Routing_Decisions"],"conditional_formatting",type(wb["Routing_Decisions"].conditional_formatting)())),"R017")])

    cases.extend([("mbox-empty",lambda root:write_mbox_states(root/MBOX,[]),"R010"),("mbox-reordered",lambda root:reorder_messages(root/MBOX),"R010"),("mbox-dropped",lambda root:write_mbox_states(root/MBOX,mbox_states(root/MBOX)[:-1]),"R010")])
    for index,(case_id,criterion) in enumerate(zip(CASES,ROUTE_CRITERIA)):
        fields=[("from","ceo@company.example.com",criterion),("to","wrong@company.example.com",criterion),("bcc","audit@company.example.com",criterion),("subject",f"Wrong subject {case_id}","R010"),("date","Thu, 16 Apr 2026 09:00:00 +0000","R010"),("message_id",f"<wrong-{index}@example.com>","R010"),("status","SENT","R010"),("route","WRONG_ROUTE","R010"),("order","99","R010"),("confidentiality","PUBLIC","R010")]
        for field,value,target in fields:cases.append((f"mbox-{case_id}-{field}",lambda root,i=index,f=field,v=value:mutate_message(root/MBOX,i,lambda state:state.update({f:v})),target))
        cases.append((f"mbox-{case_id}-body-case",lambda root,i=index,c=case_id:mutate_message(root/MBOX,i,lambda state:state.update({"body":state["body"].replace(c,"GR-WRONG",1)})),criterion))
        cases.append((f"mbox-{case_id}-body-status",lambda root,i=index:mutate_message(root/MBOX,i,lambda state:state.update({"body":state["body"].replace("DRAFT_NOT_SENT","SENT",1)})),criterion))
    for index in (2,3,6):cases.append((f"mbox-{CASES[index]}-drop-cc",lambda root,i=index:mutate_message(root/MBOX,i,lambda state:state.update({"cc":""})),ROUTE_CRITERIA[index]))
    cases.append(("mbox-manager-add-cc",lambda root:mutate_message(root/MBOX,0,lambda state:state.update({"cc":"hr-director@company.example.com"})),"R005"))

    update_cases=[("task-id",lambda p:p.update({"task_id":"wrong"})),("status",lambda p:p.update({"status":"APPLIED"})),("spreadsheet",lambda p:p.update({"spreadsheet_id":"wrong"})),("worksheet",lambda p:p.update({"worksheet_id":"wrong"})),("count",lambda p:p.update({"update_count":7})),("applied-count",lambda p:p.update({"applied_count":1})),("source-modified",lambda p:p.update({"source_modified":True})),("drop-update",lambda p:p["updates"].pop()),("reverse",lambda p:p["updates"].reverse()),("trace",lambda p:p.update({"source_trace":[]}))]
    for name,mutator in update_cases:cases.append((f"updates-{name}",lambda root,m=mutator:mutate_json(root/UPDATES,m),"R011"))
    for index,case_id in enumerate(CASES):
        for field,value in (("source_row",99),("case_id","GR-WRONG"),("route_type","WRONG"),("proposed_status","Applied"),("authority","Wrong authority"),("state","APPLIED"),("applied",True),("source_trace","wrong")):cases.append((f"updates-{case_id}-{field}",lambda root,i=index,f=field,v=value:mutate_json(root/UPDATES,lambda p:p["updates"][i].update({f:v})),"R011"))

    slack_cases=[("task-id",lambda p:p.update({"task_id":"wrong"})),("channel-id",lambda p:p["source_channel"].update({"id":"C_OTHER"})),("channel-private",lambda p:p["source_channel"].update({"is_private":True})),("status",lambda p:p.update({"status":"POSTED"})),("draft-count",lambda p:p.update({"draft_count":1})),("posted-count",lambda p:p.update({"posted_count":1})),("posted",lambda p:p.update({"posted":True})),("case-detail-count",lambda p:p.update({"case_detail_count":1})),("employee-count",lambda p:p.update({"employee_name_count":1})),("manager-count",lambda p:p.update({"manager_name_count":1})),("message",lambda p:p["messages"].append({"text":"update"})),("reason",lambda p:p.update({"reason":"Slack allowed"})),("controls",lambda p:p["controls"].pop()),("trace",lambda p:p.update({"source_trace":[]}))]
    for name,mutator in slack_cases:cases.append((f"slack-{name}",lambda root,m=mutator:mutate_json(root/SLACK,m),"R012"))
    for value in [CASES[0],EMPLOYEES[0],MANAGERS[0],DESCRIPTIONS[0]]:cases.append((f"slack-leak-{value[:12].replace(' ','-').lower()}",lambda root,v=value:mutate_json(root/SLACK,lambda p:p.update({"leaked_detail":v})),"R012"))

    execution_cases=[("task-id",lambda p:p.update({"task_id":"wrong"})),("status",lambda p:p.update({"status":"EXECUTED"})),("email-drafts",lambda p:p["notification_email"].update({"draft_count":7})),("email-sent",lambda p:p["notification_email"].update({"sent_count":1,"state":"SENT"})),("update-drafts",lambda p:p["tracker_updates"].update({"draft_count":7})),("update-applied",lambda p:p["tracker_updates"].update({"applied_count":1,"state":"APPLIED"})),("slack-drafts",lambda p:p["slack"].update({"draft_count":1})),("slack-posted",lambda p:p["slack"].update({"posted_count":1})),("docx",lambda p:p["board_artifacts"].update({"docx_count":0})),("pptx",lambda p:p["board_artifacts"].update({"pptx_count":0})),("source-modified",lambda p:p["source_workbook"].update({"modified_row_count":1,"state":"MODIFIED"})),("authorized-count",lambda p:p["confidentiality"].update({"authorized_email_case_count":7})),("slack-details",lambda p:p["confidentiality"].update({"slack_case_detail_count":1})),("identity-count",lambda p:p["confidentiality"].update({"leadership_identity_count":1})),("side-effect",lambda p:p.update({"real_side_effect_count":1})),("trace",lambda p:p.update({"source_trace":[]}))]
    for name,mutator in execution_cases:cases.append((f"execution-{name}",lambda root,m=mutator:mutate_json(root/EXECUTION,m),"R013"))

    word_cases=[("manager-count","1","2","MANAGER","R014"),("hrbp-count","4","3","HRBP","R014"),("legal-count","3","4","LEGAL_HR_DIRECTOR","R014"),("case041","GR-2026-041","GR-WRONG","GR-2026-041","R014"),("case045-route","HRBP","MANAGER","GR-2026-045","R014"),("case043-override","YES","NO","GR-2026-043","R014"),("email-sent","0","1","Notification emails","R014"),("updates-applied","0","1","Tracker updates","R014"),("slack-posted","0","1","Slack grievance messages","R014"),("source-modified","0","1","Source workbook changes","R014"),("side-effect","0","1","Real side effects","R014")]
    for name,old,new,row_token,criterion in word_cases:cases.append((f"word-{name}",lambda root,o=old,n=new,t=row_token:replace_docx_text(root/DOCX,o,n,row_token=t),criterion))
    for value in [EMPLOYEES[0],MANAGERS[0],"robert.kim@company.example.com",DESCRIPTIONS[0]]:cases.append((f"word-leak-{value[:12].replace(' ','-').lower()}",lambda root,v=value:inject_docx_text(root/DOCX,f"Restricted: {v}"),"R012"))

    cases.append(("ppt-source-count",lambda root:edit_pptx_shape_text(root/PPTX,0,2,"7"),"R015"))
    ppt_text_cases=[("manager-count","1","2"),("hrbp-count","4","3"),("legal-count","3","4"),("slack-count","0","1"),("precedence","discrimination override -> manager-subject exception -> severity route","severity route -> manager-subject exception"),("email-sent","8 drafts · 0 sent","8 drafts · 1 sent"),("update-applied","8 drafts · 0 applied","8 drafts · 1 applied"),("slack-posted","0 drafts · 0 posts","1 draft · 1 post")]
    for name,old,new in ppt_text_cases:cases.append((f"ppt-{name}",lambda root,o=old,n=new:replace_pptx_text(root/PPTX,o,n),"R015"))
    ppt_table_cases=[("case041-route",1,1,5,"HRBP"),("case043-override",1,3,3,"NO"),("case044-route",1,4,5,"HRBP"),("case045-manager",1,5,4,"NO"),("case047-route",1,7,5,"HRBP")]
    for name,slide,row,column,value in ppt_table_cases:cases.append((f"ppt-{name}",lambda root,s=slide,r=row,c=column,v=value:edit_pptx_table_cell(root/PPTX,s,r,c,v),"R015"))
    cases.append(("ppt-drop-slide",lambda root:drop_ppt_slide(root/PPTX),"R015"))
    for value in [EMPLOYEES[0],MANAGERS[0],"robert.kim@company.example.com",DESCRIPTIONS[0]]:cases.append((f"ppt-leak-{value[:12].replace(' ','-').lower()}",lambda root,v=value:inject_pptx_text(root/PPTX,f"Restricted: {v}"),"R012"))
    return cases


def main()->None:
    parser=argparse.ArgumentParser();parser.add_argument("--json-out",type=Path);args=parser.parse_args();judge=load_judge();cases=build_cases();results=[];failures=[]
    for name,mutator,expected_failure in cases:
        with tempfile.TemporaryDirectory(prefix=f"negative-5117-{name}-") as temporary:
            artifact_root=Path(temporary)/"artifacts";shutil.copytree(GOLDEN,artifact_root)
            try:
                mutator(artifact_root);result=judge.evaluate(artifact_root);detected=not result.get("criteria",{}).get(expected_failure,{}).get("passed",True);rejected=not result.get("passed",False) and float(result.get("reward",1.0))<0.8
                results.append({"case":name,"expected_failure":expected_failure,"detected":detected,"rejected":rejected,"reward":result.get("reward"),"hurdle_failures":result.get("hurdle_failures",[])})
                if not detected or not rejected:failures.append(f"{name}: expected={expected_failure}, detected={detected}, reward={result.get('reward')}")
            except Exception as exc:failures.append(f"{name}: mutation/evaluation error: {type(exc).__name__}: {exc}")
    payload={"task_id":TASK_ID if "TASK_ID" in globals() else "cowork-evo/automationbench-5117-grievance-routing-confidentiality-package","case_count":len(results),"declared_case_count":len(cases),"all_rejected":not failures and len(results)==len(cases),"max_reward":max((float(item["reward"]) for item in results),default=None),"failures":failures,"results":results}
    if args.json_out:args.json_out.parent.mkdir(parents=True,exist_ok=True);args.json_out.write_text(json.dumps(payload,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    print(json.dumps(payload,ensure_ascii=False,indent=2))
    if failures or len(results)!=len(cases):raise SystemExit(1)


if __name__=="__main__":main()
