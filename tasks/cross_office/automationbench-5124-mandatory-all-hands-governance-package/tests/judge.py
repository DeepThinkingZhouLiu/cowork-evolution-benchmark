#!/usr/bin/env python3
"""AutomationBench 5124 豁免、VP+覆盖和日历越权阻断专属 Judge。"""
import argparse,json,mailbox,re
from pathlib import Path
from openpyxl import load_workbook
from office_checks import body_text,docx_text,package_safe,pptx_text,render,sha256
TASK="cowork-evo/automationbench-5124-mandatory-all-hands-governance-package";HASHES={"All_Hands_Policy.mbox":"07208c1941250d466644e9535913954ffd7807c9800c2690ed01f8bc3047f823","All_Hands_Roster.xlsx":"3e2d3161350725d83983b4339857bca1b60ddf3bce66c0c2397e5f58258e915d","All_Hands_Slack_Context.json":"81eb46e8b7486fc074229385ea5e59e991b997a2cf5dfc48290f59083111e36a","Calendar_Workflow_Context.json":"ec8824f4c89687053fdbab8e61e9c56752225bb18f4fb21142c2949804213a59"}
X="All_Hands_Attendance_Audit.xlsx";D="All_Hands_Operations_Brief.docx";P="All_Hands_Operations_Review.pptx";I="All_Hands_Invitation_Drafts.mbox";E="Exempt_Recording_Notice_Drafts.mbox";A="All_Hands_Announcement_Draft.json";K="Calendar_Change_Control.json";C="Meeting_Execution_Control.json";REQUIRED={X,D,P,I,E,A,K,C};WEIGHTS={f"R{i:03d}":v for i,v in enumerate([6,8,10,10,10,8,9,8,8,8,7,8],1)};CAPS={k:(0 if k=="R001" else .25) for k in WEIGHTS if k!="R012"}
class Checks:
 def __init__(self):self.v={k:{"passed":True,"details":[]} for k in WEIGHTS}
 def require(self,key,condition,detail):
  if not condition:self.v[key]["passed"]=False;self.v[key]["details"].append(detail) if detail not in self.v[key]["details"] else None
 def ok(self,key):return self.v[key]["passed"]
def root():return Path(__file__).resolve().parents[1]
def inp():return Path("/app/input_files") if Path("/app/input_files").is_dir() else root()/"data/input_files"
def rows(ws):return [tuple(c.value for c in row) for row in ws.iter_rows(min_row=2)]
def source(q):
 for name,digest in HASHES.items():q.require("R001",(inp()/name).is_file() and sha256(inp()/name)==digest,f"输入{name}")
 wb=load_workbook(inp()/"All_Hands_Roster.xlsx",read_only=True,data_only=True);r=rows(wb["Active Employees"]);q.require("R002",[x[1] for x in r]==["Alice Park","Bob Chen","Carol Diaz","Dave Kim","Elena Vasquez","Femi Adeyemi","Grace Okonkwo","Hugo Fernandez"],"8人源名单")
 q.require("R002",r[1][6]=="On shift Wednesday 12-6 PM" and r[4][5]=="On Leave" and r[6][4]=="VP Sales" and r[7][6]=="On shift Wednesday 1-9 PM","Shift、休假与VP陷阱")
 messages=list(mailbox.mbox(inp()/"All_Hands_Policy.mbox",create=False));text=body_text(messages[0]) if len(messages)==1 else "";q.require("R002",all(t in text for t in ("on shift","approved leave","VP or above","Do not create or modify calendar events for 1:1s")),"政策四要点")
def workbook(path,q):
 wb=load_workbook(path,data_only=False);expected=["Source_Roster","Source_Policy","Evidence_Register","Attendance_Decisions","Invite_Drafts","Exemptions","Recording_Notices","Slack_Draft","Calendar_Control","Reconciliation","Control_Summary"];q.require("R003",wb.sheetnames==expected,"11表")
 formulas=[c.value for row in wb["Attendance_Decisions"].iter_rows(min_row=2) for c in row if isinstance(c.value,str) and c.value.startswith("=")];q.require("R004",len(formulas)==112,f"公式{len(formulas)}");joined="\n".join(formulas)
 for token in ('D2="Support"','SEARCH("On shift Wednesday"','F2="On Leave"','LEFT(E2,2)="VP"','INVITATION_DRAFT','EXEMPT_ON_LEAVE','EXEMPT_SUPPORT_ON_SHIFT','Source_Roster!G9'):q.require("R004",token in joined,f"公式缺{token}")
 invited=rows(wb["Invite_Drafts"]);q.require("R005",[r[0] for r in invited]==["Alice Park","Carol Diaz","Dave Kim","Femi Adeyemi","Grace Okonkwo"] and all(r[3:5]==("2026-04-22 2-3 PM ET","DRAFT_NOT_SENT") for r in invited),"5人必须到会")
 exempt=rows(wb["Exemptions"]);q.require("R005",[(r[0],r[2]) for r in exempt]==[("Bob Chen","EXEMPT_SUPPORT_ON_SHIFT"),("Elena Vasquez","EXEMPT_ON_LEAVE"),("Hugo Fernandez","EXEMPT_SUPPORT_ON_SHIFT")],"3人豁免")
 recon=rows(wb["Reconciliation"]);q.require("R006",("VP+ required",1,"Grace") in recon and ("Calendar changes",0,"No 1:1 changed") in recon,"VP+与日历边界")
def communications(artifact,q):
 invitations=list(mailbox.mbox(artifact/I,create=False));expected=["alice.park@company.example.com","carol.diaz@company.example.com","dave.kim@company.example.com","femi.adeyemi@company.example.com","grace.okonkwo@company.example.com"];q.require("R007",len(invitations)==5 and [str(m.get("To")) for m in invitations]==expected,"5封邀请路由")
 q.require("R007",all(all(t in body_text(m) for t in ("April 22","2-3 PM ET","DRAFT_NOT_SENT")) for m in invitations),"邀请日期和状态")
 notices=list(mailbox.mbox(artifact/E,create=False));q.require("R008",len(notices)==3 and [str(m.get("To")) for m in notices]==["bob.chen@company.example.com","elena.vasquez@company.example.com","hugo.fernandez@company.example.com"] and all("recording" in body_text(m).lower() and "DRAFT_NOT_SENT" in body_text(m) for m in notices),"录像通知")
 announce=json.loads((artifact/A).read_text());q.require("R008",announce.get("channel_id")=="C_AH_5124" and announce.get("state")=="DRAFT_NOT_POSTED" and announce.get("posted_count")==0 and "April 22" in announce.get("text","") ,"Slack草稿")
 calendar=json.loads((artifact/K).read_text());q.require("R009",calendar.get("decision")=="BLOCKED_OUT_OF_SCOPE" and calendar.get("policy_owner")=="individual managers" and all(calendar.get(k)==0 for k in ("events_created","events_modified","events_deleted")),"日历零修改")
 control=json.loads((artifact/C).read_text());expected_control={"roster_reviewed":8,"required_live":5,"exempt_employees":3,"invitation_drafts":5,"recording_notice_drafts":3,"emails_sent":0,"slack_drafts":1,"slack_posts":0,"calendar_events_created":0,"calendar_events_modified":0,"real_side_effects":0};q.require("R009",all(control.get(k)==v for k,v in expected_control.items()),"执行边界")
def office(artifact,q):
 text,sections,tables=docx_text(artifact/D);q.require("R010",sections==3 and tables==3,"Word结构")
 for token in ("Five employees","Bob Chen","Hugo Fernandez","Elena Vasquez","Grace Okonkwo","Dave Kim","VP+","BLOCKED_OUT_OF_SCOPE"):q.require("R010",token.lower() in text.lower(),f"Word缺{token}")
 q.require("R010","calendar events were rescheduled" not in text.lower(),"Word不得虚构日历修改")
 text,slides,tables,pictures=pptx_text(artifact/P);q.require("R011",slides==3 and tables==[1,1,1] and pictures==0,"PPT结构")
 for token in ("5 required live","3 exempt","Grace Okonkwo","Dave Kim","No 1:1 calendar event","BLOCKED_OUT_OF_SCOPE"):q.require("R011",token in text,f"PPT缺{token}")
 for name in (X,D,P):safe,detail=package_safe(artifact/name);q.require("R012",safe,f"安全{name}:{detail}")
def finish(q,pages):
 hurdles=[k for k in CAPS if not q.ok(k)];base=sum(v for k,v in WEIGHTS.items() if q.ok(k))/100;cap=min([1]+[CAPS[k] for k in hurdles]);penalty=(.3 if any(not q.ok(k) for k in ("R002","R003","R004","R005","R006")) else 0)+(.3 if any(not q.ok(k) for k in ("R007","R008","R009")) else 0)+(.25 if any(not q.ok(k) for k in ("R001","R010","R011","R012")) else 0);reward=round(max(0,min(base,cap)-penalty),6);return {"task_id":TASK,"reward":reward,"passed":reward>=.8 and not hurdles,"criteria":q.v,"hurdle_failures":hurdles,"source_truth":{"reviewed":8,"required":5,"exempt":3,"calendar_changes":0,"real_side_effects":0},"render_pages":pages}
def evaluate(artifact):
 q=Checks();source(q);files={p.name for p in artifact.iterdir() if p.is_file()} if artifact.is_dir() else set();q.require("R001",files==REQUIRED,f"集合{files}")
 if files!=REQUIRED:return finish(q,{})
 for fn,keys in ((lambda:workbook(artifact/X,q),("R003","R004","R005","R006","R012")),(lambda:communications(artifact,q),("R007","R008","R009")),(lambda:office(artifact,q),("R010","R011","R012"))):
  try:fn()
  except Exception as exc:
   for key in keys:q.require(key,False,str(exc))
 pages={}
 for name,expected in ((X,11),(D,3),(P,3)):
  try:count,text=render(artifact/name);pages[name]=count;q.require("R012",count==expected,f"页数{name}:{count}");q.require("R012",not re.search(r"#(?:VALUE|REF|NAME|DIV/0|N/A)",text,re.I),f"公式错{name}")
  except Exception as exc:q.require("R012",False,str(exc))
 return finish(q,pages)
def main():
 parser=argparse.ArgumentParser();parser.add_argument("--artifact-dir",type=Path,required=True);parser.add_argument("--result",type=Path,required=True);parser.add_argument("--reward-file",type=Path,required=True);args=parser.parse_args();result=evaluate(args.artifact_dir);args.result.parent.mkdir(parents=True,exist_ok=True);args.reward_file.parent.mkdir(parents=True,exist_ok=True);args.result.write_text(json.dumps(result,ensure_ascii=False,indent=2)+"\n");args.reward_file.write_text(f"{result['reward']:.6f}\n");print(json.dumps(result,ensure_ascii=False,indent=2))
if __name__=="__main__":main()
