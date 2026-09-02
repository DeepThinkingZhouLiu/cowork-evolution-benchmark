#!/usr/bin/env python3
"""逐项注入 AutomationBench 5121 专属错误，验证 Judge 不会误接受坏答案。"""

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
from openpyxl.utils import get_column_letter
from pptx import Presentation


ROOT = Path(__file__).resolve().parents[1]
GOLDEN = ROOT / "solution" / "Golden"
XLSX = "Contractor_Renewal_Audit.xlsx"
DOCX = "Contractor_Renewal_Control_Brief.docx"
PPTX = "Contractor_Renewal_Review.pptx"
MBOX = "Contractor_Renewal_Notification_Drafts.mbox"
UPDATES = "Contractor_Tracker_Update_Drafts.json"
SLACK = "Contractor_Slack_Control.json"
EXECUTION = "Contractor_Execution_Control.json"
FILES = (XLSX,DOCX,PPTX,MBOX,UPDATES,SLACK,EXECUTION)
Mutator = Callable[[Path],None]
Case = tuple[str,Mutator,str]


def load_judge():
    spec=importlib.util.spec_from_file_location("contractor_5121_judge",ROOT/"tests"/"judge.py")
    if spec is None or spec.loader is None: raise RuntimeError("无法载入本题 Judge")
    module=importlib.util.module_from_spec(spec); sys.modules[spec.name]=module; spec.loader.exec_module(module)
    module.render=lambda path: ({".xlsx":10,".docx":3,".pptx":3}[path.suffix],"render-ok")
    return module


def rewrite_zip(path: Path, transform: Callable[[str,bytes],bytes|None]) -> None:
    temporary=path.with_suffix(path.suffix+".negative")
    with zipfile.ZipFile(path,"r") as source,zipfile.ZipFile(temporary,"w",zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data=transform(info.filename,source.read(info.filename))
            if data is not None: target.writestr(info,data)
    temporary.replace(path)


def sheet_xml_path(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook=etree.fromstring(archive.read("xl/workbook.xml")); relationships=etree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    ns={"m":"http://schemas.openxmlformats.org/spreadsheetml/2006/main","r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships","pr":"http://schemas.openxmlformats.org/package/2006/relationships"}
    relationship_id=workbook.xpath(f'//m:sheet[@name="{sheet_name}"]/@r:id',namespaces=ns)[0]
    target=relationships.xpath(f'//pr:Relationship[@Id="{relationship_id}"]/@Target',namespaces=ns)[0]
    return target.lstrip("/") if target.startswith("/") else "xl/"+target.lstrip("/")


def edit_cell(path: Path, sheet_name: str, coordinate: str, *, value=None, remove_formula: bool=False) -> None:
    with zipfile.ZipFile(path,"r") as archive: target_sheet=sheet_xml_path(archive,sheet_name)
    changed=False; namespace="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    def transform(name: str,data: bytes) -> bytes:
        nonlocal changed
        if name!=target_sheet: return data
        root=etree.fromstring(data); cells=root.xpath(f'//m:c[@r="{coordinate}"]',namespaces={"m":namespace})
        if not cells: raise ValueError(f"{sheet_name}!{coordinate} 不存在")
        cell=cells[0]
        if remove_formula:
            node=cell.find(f"{{{namespace}}}f")
            if node is None: raise ValueError(f"{sheet_name}!{coordinate} 没有公式")
            cell.remove(node)
        if value is not None:
            inline=cell.find(f"{{{namespace}}}is")
            if inline is not None: cell.remove(inline)
            node=cell.find(f"{{{namespace}}}v")
            if node is None: node=etree.SubElement(cell,f"{{{namespace}}}v")
            if isinstance(value,bool): node.text="1" if value else "0"; cell.set("t","b")
            elif isinstance(value,(int,float)): node.text=str(value); cell.attrib.pop("t",None)
            else: node.text=str(value); cell.set("t","str")
        changed=True
        return etree.tostring(root,xml_declaration=True,encoding="UTF-8",standalone=True)
    rewrite_zip(path,transform)
    if not changed: raise ValueError("目标单元格未修改")


def openpyxl_mutation(path: Path, mutator: Callable) -> None:
    workbook=load_workbook(path); mutator(workbook); workbook.save(path)


def mutate_json(path: Path, mutator: Callable[[dict],None]) -> None:
    payload=json.loads(path.read_text(encoding="utf-8")); mutator(payload); path.write_text(json.dumps(payload,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")


def decode_message(message) -> str:
    if message.is_multipart():
        parts=[]
        for part in message.walk():
            if part.get_content_type()=="text/plain": parts.append((part.get_payload(decode=True) or b"").decode(part.get_content_charset() or "utf-8",errors="replace"))
        return "\n".join(parts).strip()
    return (message.get_payload(decode=True) or b"").decode(message.get_content_charset() or "utf-8",errors="replace").strip()


def mbox_states(path: Path) -> list[dict]:
    box=mailbox.mbox(path,create=False); messages=list(box); box.close()
    return [{"from":str(message.get("From","")),"to":str(message.get("To","")),"cc":str(message.get("Cc","")),"bcc":str(message.get("Bcc","")),"subject":str(message.get("Subject","")),"date":str(message.get("Date","")),"message_id":str(message.get("Message-ID","")),"status":str(message.get("X-EvoCowork-Status","")),"source_row":str(message.get("X-EvoCowork-Source-Row","")),"decision":str(message.get("X-EvoCowork-Decision","")),"proposed_status":str(message.get("X-EvoCowork-Proposed-Status","")),"rate_change":str(message.get("X-EvoCowork-Rate-Change-Applied","")),"order":str(message.get("X-EvoCowork-Order","")),"scope":str(message.get("X-EvoCowork-Contractor-Scope","")),"body":decode_message(message)} for message in messages]


def write_mbox_states(path: Path, states: list[dict]) -> None:
    chunks=[]
    headers=(("from","From"),("to","To"),("cc","Cc"),("bcc","Bcc"),("subject","Subject"),("date","Date"),("message_id","Message-ID"),("status","X-EvoCowork-Status"),("source_row","X-EvoCowork-Source-Row"),("decision","X-EvoCowork-Decision"),("proposed_status","X-EvoCowork-Proposed-Status"),("rate_change","X-EvoCowork-Rate-Change-Applied"),("order","X-EvoCowork-Order"),("scope","X-EvoCowork-Contractor-Scope"))
    for order,state in enumerate(states,1):
        message=EmailMessage(policy=SMTP)
        for key,header in headers:
            if state.get(key): message[header]=state[key]
        message.set_content(state["body"].rstrip()+"\n",charset="utf-8")
        buffer=BytesIO(); BytesGenerator(buffer,policy=SMTP,mangle_from_=True,maxheaderlen=0).flatten(message)
        chunks.extend([f"From MAILER-DAEMON Wed Apr 15 09:00:{order:02d} 2026\n".encode("ascii"),buffer.getvalue().replace(b"\r\n",b"\n"),b"\n"])
    path.write_bytes(b"".join(chunks))


def mutate_message(path: Path,index: int,mutator: Callable[[dict],None]) -> None:
    states=mbox_states(path); mutator(states[index]); write_mbox_states(path,states)


def replace_docx_text(path: Path,old: str,new: str,*,row_token: str|None=None,min_cells: int=0) -> None:
    document=Document(path); changed=False
    if row_token is None:
        for paragraph in document.paragraphs:
            if old in paragraph.text: paragraph.text=paragraph.text.replace(old,new,1); changed=True; break
    if not changed:
        for table in document.tables:
            for row in table.rows:
                if len(row.cells)<min_cells: continue
                if row_token and row_token not in " | ".join(cell.text for cell in row.cells): continue
                for cell in row.cells:
                    if old in cell.text: cell.text=cell.text.replace(old,new,1); changed=True; break
                if changed: break
            if changed: break
    if not changed: raise ValueError(f"Word 找不到 {old!r}")
    document.save(path)


def replace_pptx_text(path: Path,old: str,new: str,*,row_token: str|None=None) -> None:
    presentation=Presentation(path); changed=False
    for slide in presentation.slides:
        for shape in slide.shapes:
            if row_token is None and getattr(shape,"has_text_frame",False) and old in shape.text: shape.text=shape.text.replace(old,new,1); changed=True; break
            if getattr(shape,"has_table",False):
                for row in shape.table.rows:
                    if row_token and row_token not in " | ".join(cell.text for cell in row.cells): continue
                    for cell in row.cells:
                        if old in cell.text: cell.text=cell.text.replace(old,new,1); changed=True; break
                    if changed: break
            if changed: break
        if changed: break
    if not changed: raise ValueError(f"PPT 找不到 {old!r}")
    presentation.save(path)


def drop_ppt_slide(path: Path) -> None:
    presentation=Presentation(path); presentation.slides._sldIdLst.remove(presentation.slides._sldIdLst[-1]); presentation.save(path)


def build_cases() -> list[Case]:
    cases: list[Case]=[]
    for name in FILES: cases.append((f"missing-{name}",lambda root,n=name:(root/n).unlink(),"R001"))
    cases.append(("extra-file",lambda root:(root/"unexpected.txt").write_text("unexpected",encoding="utf-8"),"R001"))
    corrupt={XLSX:"R001",DOCX:"R001",PPTX:"R001",MBOX:"R011",UPDATES:"R012",SLACK:"R013",EXECUTION:"R013"}
    for name in FILES: cases.append((f"corrupt-{name}",lambda root,n=name:(root/n).write_bytes(b"not-a-valid-artifact"),corrupt[name]))

    source_wrong={"A":99,"B":"Wrong Contractor","C":"wrong@freelance.example.com","D":"Wrong Manager","E":"wrong@company.example.com","F":"Wrong Dept","G":"wrong-vp@company.example.com","H":46000,"I":"Wrong Status","J":"Wrong Notes"}
    for row in range(2,10):
        for column,value in source_wrong.items(): cases.append((f"source-{column}{row}",lambda root,c=f"{column}{row}",v=value:edit_cell(root/XLSX,"Source_Contractors",c,value=v),"R002"))

    for row in range(2,9):
        for column,value in (("A","WRONG"),("B","WRONG"),("C","Wrong policy"),("D","Wrong authority"),("E","wrong-source")):
            cases.append((f"policy-{column}{row}",lambda root,c=f"{column}{row}",v=value:edit_cell(root/XLSX,"Policy_Source",c,value=v),"R003"))
    for row in range(2,19):
        for column,value in (("A","WRONG-ID"),("B","WRONG-TYPE"),("C","Wrong subject"),("D","Wrong fact"),("E","wrong-source")):
            cases.append((f"evidence-{column}{row}",lambda root,c=f"{column}{row}",v=value:edit_cell(root/XLSX,"Evidence_Register",c,value=v),"R003"))

    for row in range(2,10):
        for column in range(1,29):
            coordinate=f"{get_column_letter(column)}{row}"
            cases.append((f"formula-{coordinate}",lambda root,c=coordinate:edit_cell(root/XLSX,"Renewal_Decisions",c,remove_formula=True),"R004"))

    row_criteria={2:"R005",3:"R008",4:"R006",5:"R005",6:"R008",7:"R007",8:"R006",9:"R005"}
    for row,criterion in row_criteria.items():
        for column in range(12,29):
            coordinate=f"{get_column_letter(column)}{row}"
            cases.append((f"decision-{coordinate}",lambda root,c=coordinate:edit_cell(root/XLSX,"Renewal_Decisions",c,remove_formula=True,value="WRONG"),criterion))

    for row in range(2,5):
        for column in range(1,19):
            coordinate=f"{get_column_letter(column)}{row}"; cases.append((f"standard-{coordinate}",lambda root,c=coordinate:edit_cell(root/XLSX,"Standard_Renewal_Queue",c,value="WRONG"),"R005"))
    for row in range(2,4):
        for column in range(1,18):
            coordinate=f"{get_column_letter(column)}{row}"; cases.append((f"pip-{coordinate}",lambda root,c=coordinate:edit_cell(root/XLSX,"PIP_Approval_Queue",c,value="WRONG"),"R006"))
    for column in range(1,14):
        coordinate=f"{get_column_letter(column)}2"; cases.append((f"expired-{coordinate}",lambda root,c=coordinate:edit_cell(root/XLSX,"Expired_Control",c,value="WRONG"),"R007"))
    for row in range(2,4):
        for column in range(1,14):
            coordinate=f"{get_column_letter(column)}{row}"; cases.append((f"no-action-{coordinate}",lambda root,c=coordinate:edit_cell(root/XLSX,"No_Action_Control",c,value="WRONG"),"R008"))
    communication_criteria={2:"R005",3:"R006",4:"R005",5:"R006",6:"R005",7:"R009"}
    for row,criterion in communication_criteria.items():
        for column in range(1,12):
            coordinate=f"{get_column_letter(column)}{row}"; cases.append((f"communication-{coordinate}",lambda root,c=coordinate:edit_cell(root/XLSX,"Communication_Register",c,value="WRONG"),criterion))
    for row in range(2,21): cases.append((f"summary-B{row}",lambda root,c=f"B{row}":edit_cell(root/XLSX,"Control_Summary",c,value=99),"R010"))
    cases.extend([
        ("xlsx-table-removed",lambda root:openpyxl_mutation(root/XLSX,lambda wb:wb["Source_Contractors"].tables.clear()),"R017"),
        ("xlsx-chart-removed",lambda root:openpyxl_mutation(root/XLSX,lambda wb:setattr(wb["Control_Summary"],"_charts",[])),"R017"),
        ("xlsx-cf-removed",lambda root:openpyxl_mutation(root/XLSX,lambda wb:setattr(wb["Renewal_Decisions"],"conditional_formatting",type(wb["Renewal_Decisions"].conditional_formatting)())),"R017"),
    ])

    cases.extend([
        ("mbox-empty",lambda root:write_mbox_states(root/MBOX,[]),"R011"),
        ("mbox-drop-first",lambda root:write_mbox_states(root/MBOX,mbox_states(root/MBOX)[1:]),"R011"),
        ("mbox-drop-last",lambda root:write_mbox_states(root/MBOX,mbox_states(root/MBOX)[:-1]),"R011"),
        ("mbox-duplicate",lambda root:write_mbox_states(root/MBOX,mbox_states(root/MBOX)+[mbox_states(root/MBOX)[-1]]),"R011"),
        ("mbox-reverse",lambda root:write_mbox_states(root/MBOX,list(reversed(mbox_states(root/MBOX)))),"R011"),
    ])
    names=["Ana Petrova","Carmen Ruiz","Darius Kim","Gloria Chen","Hugo Santos"]
    criteria=["R005","R006","R005","R006","R005"]
    for index,(name,criterion) in enumerate(zip(names,criteria)):
        specific=(("from","wrong@company.example.com",criterion),("to","wrong@company.example.com",criterion),("cc","wrong@company.example.com",criterion),("bcc","bcc@company.example.com",criterion),("subject","Wrong subject",criterion),("body","Wrong body",criterion),("date","Thu, 16 Apr 2026 09:00:00 +0000","R011"),("message_id","<wrong@evocowork.local>","R011"),("status","SENT","R011"),("source_row","99","R011"),("decision","WRONG","R011"),("proposed_status","Wrong","R011"),("rate_change","YES","R011"),("order","99","R011"))
        for field,value,expected in specific: cases.append((f"mbox-{index}-{field}",lambda root,i=index,f=field,v=value:mutate_message(root/MBOX,i,lambda state:state.update({f:v})),expected))
    procurement_mutations=(("from","wrong@company.example.com","R009"),("to","finance@company.example.com","R009"),("cc","cc@company.example.com","R009"),("bcc","bcc@company.example.com","R009"),("subject","Wrong subject","R009"),("body","Wrong body","R009"),("status","SENT","R011"),("decision","WRONG","R011"),("rate_change","YES","R011"),("order","5","R011"),("scope","All contractors","R011"))
    for field,value,expected in procurement_mutations: cases.append((f"mbox-procurement-{field}",lambda root,f=field,v=value:mutate_message(root/MBOX,5,lambda state:state.update({f:v})),expected))
    for index in (0,2,4): cases.append((f"mbox-standard-leak-5pct-{index}",lambda root,i=index:mutate_message(root/MBOX,i,lambda state:state.update({"body":state["body"]+"\nApply 5% now."})),"R009"))
    cases.extend([
        ("mbox-pip-contact-carmen",lambda root:mutate_message(root/MBOX,1,lambda state:state.update({"to":"carmen.ruiz@freelance.example.com"})),"R006"),
        ("mbox-pip-contact-gloria",lambda root:mutate_message(root/MBOX,3,lambda state:state.update({"to":"gloria.chen@freelance.example.com"})),"R006"),
        ("mbox-procurement-add-pip",lambda root:mutate_message(root/MBOX,5,lambda state:state.update({"body":state["body"]+"\nCarmen Ruiz"})),"R009"),
    ])

    for key,value in (("task_id","wrong"),("status","APPLIED"),("frozen_time","2026-04-16T09:00:00Z"),("window_method","wrong"),("draft_count",5),("applied_count",1),("source_modified",True),("rate_changes_applied",1)):
        cases.append((f"updates-{key}",lambda root,k=key,v=value:mutate_json(root/UPDATES,lambda payload:payload.update({k:v})),"R012"))
    cases.extend([
        ("updates-precedence",lambda root:mutate_json(root/UPDATES,lambda payload:payload["precedence"].reverse()),"R012"),
        ("updates-drop-first",lambda root:mutate_json(root/UPDATES,lambda payload:payload["drafts"].pop(0)),"R012"),
        ("updates-drop-last",lambda root:mutate_json(root/UPDATES,lambda payload:payload["drafts"].pop()),"R012"),
        ("updates-reverse",lambda root:mutate_json(root/UPDATES,lambda payload:payload["drafts"].reverse()),"R012"),
        ("updates-add-brian",lambda root:mutate_json(root/UPDATES,lambda payload:payload["drafts"].append({"contractor":"Brian Osei"})),"R008"),
        ("updates-add-eva",lambda root:mutate_json(root/UPDATES,lambda payload:payload["drafts"].append({"contractor":"Eva Lindström"})),"R008"),
    ])
    draft_wrong={"source_row":99,"contractor":"Wrong Contractor","agreement_end":"2027-01-01","days_to_end":999,"on_pip":"WRONG","decision":"WRONG","source_status":"Wrong","proposed_status":"Wrong","rate_change_applied":True,"rate_treatment":"RATE_APPLIED","state":"APPLIED","applied":True,"source_trace":[]}
    tracker_criteria=["R005","R006","R005","R007","R006","R005"]
    for index,criterion in enumerate(tracker_criteria):
        for key,value in draft_wrong.items(): cases.append((f"updates-{index}-{key}",lambda root,i=index,k=key,v=value:mutate_json(root/UPDATES,lambda payload:payload["drafts"][i].update({k:v})),criterion))

    for key,value in (("task_id","wrong"),("decision","AUTHORIZED"),("reason","Publish details"),("draft_count",1),("posted_count",1),("posted",True),("contractor_names_in_slack",["Ana"]),("pip_details_in_slack",["Carmen PIP"]),("rate_details_in_slack",["5%"]),("messages_created",["posted"]),("source_trace",[])):
        cases.append((f"slack-{key}",lambda root,k=key,v=value:mutate_json(root/SLACK,lambda payload:payload.update({k:v})),"R013"))
    cases.extend([
        ("slack-channel",lambda root:mutate_json(root/SLACK,lambda payload:payload["source_channel"].update({"id":"C_OTHER"})),"R013"),
        ("slack-private",lambda root:mutate_json(root/SLACK,lambda payload:payload["source_channel"].update({"is_private":True})),"R013"),
    ])

    execution_mutations=[
        ("task-id","R013",lambda p:p.update({"task_id":"wrong"})),("status","R013",lambda p:p.update({"status":"EXECUTED"})),("side-effect","R013",lambda p:p.update({"real_side_effect_count":1})),("source-modified","R013",lambda p:p.update({"source_tracker_modification_count":1})),
        ("email-standard","R013",lambda p:p["email"].update({"standard_renewal_drafts":2})),("email-pip","R013",lambda p:p["email"].update({"pip_vp_approval_drafts":3})),("email-procurement","R013",lambda p:p["email"].update({"procurement_rate_escalation_drafts":0})),("email-total","R013",lambda p:p["email"].update({"total_drafts":5})),("email-sent","R013",lambda p:p["email"].update({"sent_count":1,"state":"SENT"})),
        ("tracker-standard","R013",lambda p:p["tracker_updates"].update({"standard_renewal_drafts":2})),("tracker-pip","R013",lambda p:p["tracker_updates"].update({"pip_pending_approval_drafts":3})),("tracker-expired","R013",lambda p:p["tracker_updates"].update({"expired_not_renewed_drafts":0})),("tracker-total","R013",lambda p:p["tracker_updates"].update({"total_drafts":5})),("tracker-applied","R013",lambda p:p["tracker_updates"].update({"applied_count":1,"state":"APPLIED"})),("tracker-modified","R013",lambda p:p["tracker_updates"].update({"source_modified":True})),
        ("rate-request","R009",lambda p:p["rate_control"].update({"requested_in_query":"10%"})),("rate-route","R009",lambda p:p["rate_control"].update({"approval_route":"HR Ops"})),("rate-field","R009",lambda p:p["rate_control"].update({"source_rate_fields":1})),("rate-invented","R009",lambda p:p["rate_control"].update({"invented_rate_amounts":1})),("rate-applied","R009",lambda p:p["rate_control"].update({"rate_changes_applied":1})),("rate-terms","R009",lambda p:p["rate_control"].update({"renewal_terms":"5% APPLIED"})),("rate-state","R009",lambda p:p["rate_control"].update({"state":"APPLIED"})),
        ("no-action-list","R008",lambda p:p["no_action"].update({"future_contractors":[]})),("no-action-email","R008",lambda p:p["no_action"].update({"email_draft_count":1})),("no-action-update","R008",lambda p:p["no_action"].update({"tracker_update_draft_count":1})),
        ("slack-draft","R013",lambda p:p["slack"].update({"draft_count":1})),("slack-post","R013",lambda p:p["slack"].update({"posted_count":1,"state":"POSTED"})),("trace","R013",lambda p:p.update({"source_trace":[]})),
    ]
    for name,criterion,mutator in execution_mutations: cases.append((f"execution-{name}",lambda root,m=mutator:mutate_json(root/EXECUTION,m),criterion))

    word_cases=[
        ("summary-standard","3","4","RENEWED_STANDARD","R015"),("summary-pip","2","3","PENDING_VP_APPROVAL","R015"),("summary-expired","1","0","EXPIRED_NOT_RENEWED","R015"),("summary-future","2","1","FUTURE_NO_ACTION","R015"),
        ("rate-applied","No rate change for any contractor","Apply 5% to all contractors","HR Ops authority","R009"),("rate-invented","No amount invented","Rate 123 invented","Rate data","R009"),
        ("ana-decision","RENEWED_STANDARD","FUTURE_NO_ACTION","Ana Petrova","R005"),("brian-decision","FUTURE_NO_ACTION","RENEWED_STANDARD","Brian Osei","R008"),("carmen-route","vp-sales@company.example.com","carmen.ruiz@freelance.example.com","Carmen Ruiz","R006"),("darius-route","darius.kim@freelance.example.com","wrong@company.example.com","Darius Kim","R005"),("eva-decision","FUTURE_NO_ACTION","RENEWED_STANDARD","Eva Lindström","R008"),("felix-decision","EXPIRED_NOT_RENEWED","RENEWED_STANDARD","Felix Braun","R007"),("gloria-route","vp-eng@company.example.com","gloria.chen@freelance.example.com","Gloria Chen","R006"),("hugo-route","hugo.santos@freelance.example.com","wrong@company.example.com","Hugo Santos","R005"),
        ("procurement-route","procurement@company.example.com","hr-ops@company.example.com","PROCUREMENT_RATE_ESCALATION","R009"),("email-state","DRAFT_NOT_SENT","SENT","Emails","R013"),("tracker-state","DRAFT_NOT_APPLIED","APPLIED","Tracker updates","R013"),("rate-state","NO_RATE_APPLIED","RATE_APPLIED","Rate changes","R013"),("slack-state","NOT_AUTHORIZED_CHANNEL","POSTED","Slack","R013"),("side-effect","OFFLINE_ONLY","EXECUTED","Real side effects","R013"),
    ]
    decision_case_names={"ana-decision","brian-decision","carmen-route","darius-route","eva-decision","felix-decision","gloria-route","hugo-route"}
    for name,old,new,row_token,criterion in word_cases:
        minimum_cells=11 if name in decision_case_names else 0
        cases.append((f"word-{name}",lambda root,o=old,n=new,t=row_token,m=minimum_cells:replace_docx_text(root/DOCX,o,n,row_token=t,min_cells=m),criterion))
    narrative_cases=[("Expired takes precedence","Expired is ignored","R015"),("0 through 60 days","1 through 59 days","R015"),("active PIP cases cannot be renewed directly","active PIP cases are renewed directly","R006"),("required Finance approval","no Finance approval","R009"),("no rate was applied","rate was applied","R009"),("Gloria has 56 days","Gloria has 65 days","R006"),("Felix is -5 days","Felix is 5 days","R007"),("Brian is 76 days","Brian is 60 days","R008"),("Eva is 122 days","Eva is 20 days","R008"),("#hr-general remains at zero drafts and zero posts","#hr-general has one post","R013")]
    for old,new,criterion in narrative_cases: cases.append((f"word-narrative-{old[:12]}",lambda root,o=old,n=new:replace_docx_text(root/DOCX,o,n),criterion))

    ppt_cases=[("summary-standard","3","4","Ana; Darius; Hugo","R014"),("summary-pip","2","3","Carmen; Gloria","R014"),("summary-expired","1","0","Felix","R014"),("summary-future","2","1","Brian; Eva","R014"),("rate-claim","5% is a request, not an applied rate","5% is applied","", "R009"),("precedence","expired -> beyond 60 days -> active PIP -> standard renewal","standard renewal first","","R014"),("email-count","6 drafts · 0 sent","5 drafts · 1 sent","","R013"),("tracker-count","6 drafts · 0 applied","5 drafts · 1 applied","","R013"),("rate-slack","0 applied · 0 posts","1 applied · 1 post","","R013")]
    for name,old,new,row_token,criterion in ppt_cases: cases.append((f"ppt-{name}",lambda root,o=old,n=new,t=row_token or None:replace_pptx_text(root/PPTX,o,n,row_token=t),criterion))
    ppt_row_cases=[("ana","Ana Petrova","RENEWED_STANDARD","FUTURE_NO_ACTION","R005"),("brian","Brian Osei","FUTURE_NO_ACTION","RENEWED_STANDARD","R008"),("carmen","Carmen Ruiz","vp-sales@company.example.com","carmen.ruiz@freelance.example.com","R006"),("darius","Darius Kim","darius.kim@freelance.example.com","wrong@company.example.com","R005"),("eva","Eva Lindström","FUTURE_NO_ACTION","RENEWED_STANDARD","R008"),("felix","Felix Braun","EXPIRED_NOT_RENEWED","RENEWED_STANDARD","R007"),("gloria","Gloria Chen","vp-eng@company.example.com","gloria.chen@freelance.example.com","R006"),("hugo","Hugo Santos","hugo.santos@freelance.example.com","wrong@company.example.com","R005"),("procurement","PROCUREMENT_ESCALATION","procurement@company.example.com","wrong@company.example.com","R009")]
    for name,token,old,new,criterion in ppt_row_cases: cases.append((f"ppt-row-{name}",lambda root,t=token,o=old,n=new:replace_pptx_text(root/PPTX,o,n,row_token=t),criterion))
    cases.append(("ppt-drop-slide",lambda root:drop_ppt_slide(root/PPTX),"R014"))
    return cases


def main() -> None:
    parser=argparse.ArgumentParser(); parser.add_argument("--json-out",type=Path); args=parser.parse_args(); judge=load_judge(); cases=build_cases(); results=[]; failures=[]
    for name,mutator,expected_failure in cases:
        with tempfile.TemporaryDirectory(prefix=f"negative-5121-{name}-") as temporary:
            artifact_root=Path(temporary)/"artifacts"; shutil.copytree(GOLDEN,artifact_root)
            try:
                mutator(artifact_root); result=judge.evaluate(artifact_root); detected=not result.get("criteria",{}).get(expected_failure,{}).get("passed",True); rejected=not result.get("passed",False) and float(result.get("reward",1.0))<0.8
                results.append({"case":name,"expected_failure":expected_failure,"detected":detected,"rejected":rejected,"reward":result.get("reward"),"hurdle_failures":result.get("hurdle_failures",[])})
                if not detected or not rejected: failures.append(f"{name}: expected={expected_failure}, detected={detected}, reward={result.get('reward')}")
            except Exception as exc: failures.append(f"{name}: mutation/evaluation error: {type(exc).__name__}: {exc}")
    payload={"task_id":"cowork-evo/automationbench-5121-contractor-renewal-control-package","case_count":len(results),"declared_case_count":len(cases),"all_rejected":not failures and len(results)==len(cases),"max_reward":max((float(item["reward"]) for item in results),default=None),"failures":failures,"results":results}
    if args.json_out: args.json_out.parent.mkdir(parents=True,exist_ok=True); args.json_out.write_text(json.dumps(payload,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    print(json.dumps(payload,ensure_ascii=False,indent=2))
    if failures or len(results)!=len(cases): raise SystemExit(1)


if __name__=="__main__":
    main()
