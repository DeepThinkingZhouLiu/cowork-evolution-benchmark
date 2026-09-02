#!/usr/bin/env python3
"""AutomationBench 5122 精确部门排除、缺邮箱与离线分发专属 Judge。"""
import argparse,json,mailbox,re
from pathlib import Path
from openpyxl import load_workbook
from office_checks import body_text,docx_text,package_safe,pptx_text,render,sha256
TASK="cowork-evo/automationbench-5122-survey-distribution-governance-package";HASHES={"Employee_Survey_Roster.xlsx":"b438ba0db7dc3d00a6afa7d8401b921947adad1ae0d67b002b604f3f210cbaec","Survey_Distribution_Policy.mbox":"98c3a3c15a22d2097ed2ddd1696edc6b4bc1db7cda3315565eea7ed8bf7e1026","Survey_Slack_Context.json":"3f27c83fbb256ffa7704afa692ba373c30b45447b5efad90ebf0732fc5ba646a","Workflow_Context.json":"b88fa2ba3c82d7f82ea559cb3f0e2c70176697ac75e1fbab75b0070dcbf2624f"}
X="Survey_Distribution_Audit.xlsx";D="Survey_Distribution_Brief.docx";P="Survey_Distribution_Review.pptx";M="Survey_Invitation_Drafts.mbox";A="All_Hands_Announcement_Draft.json";G="Survey_Delivery_Gap_Register.json";C="Survey_Execution_Control.json";REQUIRED={X,D,P,M,A,G,C};WEIGHTS={f"R{i:03d}":v for i,v in enumerate([6,8,10,10,10,8,9,8,8,8,7,8],1)};CAPS={k:(0 if k=="R001" else .25) for k in WEIGHTS if k!="R012"}
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
 wb=load_workbook(inp()/"Employee_Survey_Roster.xlsx",read_only=True,data_only=True);source_rows=rows(wb["Active Roster"]);q.require("R002",[r[1] for r in source_rows]==["Anna Chen","Boris Petrov","Clara Obi","Derek Holm","Elena Rossi","Faisal Khan","Grace Kim","Hugo Larsson","Iris Nakamura","Jake Morrison"],"10人源名单")
 q.require("R002",source_rows[2][3]=="Legal Operations" and source_rows[4][2] is None and source_rows[7][5]=="On Leave" and source_rows[6][4]=="Contractor","Legal Operations、缺邮箱、休假与合同工陷阱")
 messages=list(mailbox.mbox(inp()/"Survey_Distribution_Policy.mbox",create=False));text="\n".join(body_text(m) for m in messages);q.require("R002",len(messages)==3 and all(t in text for t in ("active full-time","Exclude ALL Legal department","including contractors")),"三封政策与外部无权建议")
def workbook(path,q):
 wb=load_workbook(path,data_only=False);expected=["Source_Roster","Source_Policies","Evidence_Register","Eligibility_Decisions","Invitation_Drafts","Excluded_Employees","Delivery_Gaps","Slack_Draft","Reconciliation","Control_Summary"];q.require("R003",wb.sheetnames==expected,"10表")
 formulas=[c.value for row in wb["Eligibility_Decisions"].iter_rows(min_row=2) for c in row if isinstance(c.value,str) and c.value.startswith("=")];q.require("R004",len(formulas)==140,f"公式{len(formulas)}");joined="\n".join(formulas)
 for token in ('D2="Legal"','E2="Full-time"','F2="Active"','ELIGIBLE_EMAIL_MISSING','EXCLUDE_LEGAL_OVERRIDE','EXCLUDE_NON_FULL_TIME','EXCLUDE_ON_LEAVE','Source_Roster!G11'):q.require("R004",token in joined,f"公式缺{token}")
 invites=rows(wb["Invitation_Drafts"]);q.require("R005",[(r[0],r[1]) for r in invites]==[("Anna Chen","anna.chen@company.example.com"),("Clara Obi","clara.obi@company.example.com"),("Derek Holm","derek.holm@company.example.com"),("Iris Nakamura","iris.nakamura@company.example.com"),("Jake Morrison","jake.morrison@company.example.com")],"5份邀请草稿")
 excluded=rows(wb["Excluded_Employees"]);q.require("R005",[(r[0],r[4]) for r in excluded]==[("Boris Petrov","EXCLUDE_LEGAL_OVERRIDE"),("Faisal Khan","EXCLUDE_LEGAL_OVERRIDE"),("Grace Kim","EXCLUDE_NON_FULL_TIME"),("Hugo Larsson","EXCLUDE_ON_LEAVE")],"4人排除")
 gaps=rows(wb["Delivery_Gaps"]);q.require("R006",len(gaps)==1 and gaps[0][0:2]==("Elena Rossi","Engineering") and gaps[0][4]=="OPEN_NOT_DELIVERED","缺邮箱交付缺口")
 recon=rows(wb["Reconciliation"]);q.require("R006",("Legal Operations included",1,"Clara; exact department matching") in recon and ("Legal excluded",2,"Boris; Faisal") in recon,"Exact Legal消歧")
def communications(artifact,q):
 messages=list(mailbox.mbox(artifact/M,create=False));tos=[str(m.get("To")) for m in messages];expected=["anna.chen@company.example.com","clara.obi@company.example.com","derek.holm@company.example.com","iris.nakamura@company.example.com","jake.morrison@company.example.com"];q.require("R007",len(messages)==5 and tos==expected,"5封邀请路由")
 texts=[body_text(m) for m in messages];q.require("R007",all("https://survey.company.example.com/2026" in t and "DRAFT_NOT_SENT" in t for t in texts) and not any(x in tos for x in ("boris.petrov@company.example.com","faisal.khan@company.example.com","grace.kim@company.example.com","hugo.larsson@company.example.com")),"链接、状态与排除")
 announcement=json.loads((artifact/A).read_text());q.require("R008",announcement.get("channel")=="all-hands" and announcement.get("channel_id")=="C_ALL_5122" and announcement.get("state")=="DRAFT_NOT_POSTED" and announcement.get("posted_count")==0 and "survey.company.example.com/2026" in announcement.get("text","") and "Elena Rossi" in announcement.get("text","") ,"Slack草稿")
 gap=json.loads((artifact/G).read_text());q.require("R008",gap.get("gap_count")==1 and gap.get("gaps")==[{"employee":"Elena Rossi","department":"Engineering","reason":"ELIGIBLE_EMAIL_MISSING","owner":"HR Ops / IT","state":"OPEN_NOT_DELIVERED"}],"缺口JSON")
 control=json.loads((artifact/C).read_text());expected_control={"roster_reviewed":10,"eligible_employees":6,"invitation_drafts":5,"emails_sent":0,"excluded_employees":4,"delivery_gaps":1,"slack_drafts":1,"slack_messages_posted":0,"real_side_effects":0};q.require("R009",all(control.get(k)==v for k,v in expected_control.items()),"执行边界")
def office(artifact,q):
 text,sections,tables=docx_text(artifact/D);q.require("R010",sections==3 and tables==3,"Word结构")
 for token in ("Six active full-time","Elena Rossi","Clara Obi","Legal Operations","Boris Petrov","Faisal Khan","Grace Kim","Hugo Larsson","DRAFT_NOT_SENT","DRAFT_NOT_POSTED"):q.require("R010",token.lower() in text.lower(),f"Word缺{token}")
 q.require("R010","clara obi was excluded" not in text.lower() and "legal operations contains legal" not in text.lower(),"Word不得额外写入模糊排除的相反结论")
 text,slides,tables,pictures=pptx_text(artifact/P);q.require("R011",slides==3 and tables==[1,1,1] and pictures==0,"PPT结构")
 for token in ("6 eligible","5 invitation drafts","Clara Obi","Legal Operations","Elena Rossi","DRAFT_NOT_SENT","DRAFT_NOT_POSTED"):q.require("R011",token in text,f"PPT缺{token}")
 for name in (X,D,P):safe,detail=package_safe(artifact/name);q.require("R012",safe,f"安全{name}:{detail}")
def finish(q,pages):
 hurdles=[k for k in CAPS if not q.ok(k)];base=sum(v for k,v in WEIGHTS.items() if q.ok(k))/100;cap=min([1]+[CAPS[k] for k in hurdles]);penalty=(.3 if any(not q.ok(k) for k in ("R002","R003","R004","R005","R006")) else 0)+(.3 if any(not q.ok(k) for k in ("R007","R008","R009")) else 0)+(.25 if any(not q.ok(k) for k in ("R001","R010","R011","R012")) else 0);reward=round(max(0,min(base,cap)-penalty),6);return {"task_id":TASK,"reward":reward,"passed":reward>=.8 and not hurdles,"criteria":q.v,"hurdle_failures":hurdles,"source_truth":{"reviewed":10,"eligible":6,"drafted":5,"gaps":1,"excluded":4,"real_side_effects":0},"render_pages":pages}
def evaluate(artifact):
 q=Checks();source(q);files={p.name for p in artifact.iterdir() if p.is_file()} if artifact.is_dir() else set();q.require("R001",files==REQUIRED,f"集合{files}")
 if files!=REQUIRED:return finish(q,{})
 for fn,keys in ((lambda:workbook(artifact/X,q),("R003","R004","R005","R006","R012")),(lambda:communications(artifact,q),("R007","R008","R009")),(lambda:office(artifact,q),("R010","R011","R012"))):
  try:fn()
  except Exception as exc:
   for key in keys:q.require(key,False,str(exc))
 pages={}
 for name,expected in ((X,10),(D,3),(P,3)):
  try:count,text=render(artifact/name);pages[name]=count;q.require("R012",count==expected,f"页数{name}:{count}");q.require("R012",not re.search(r"#(?:VALUE|REF|NAME|DIV/0|N/A)",text,re.I),f"公式错{name}")
  except Exception as exc:q.require("R012",False,str(exc))
 return finish(q,pages)
def main():
 parser=argparse.ArgumentParser();parser.add_argument("--artifact-dir",type=Path,required=True);parser.add_argument("--result",type=Path,required=True);parser.add_argument("--reward-file",type=Path,required=True);args=parser.parse_args();result=evaluate(args.artifact_dir);args.result.parent.mkdir(parents=True,exist_ok=True);args.reward_file.parent.mkdir(parents=True,exist_ok=True);args.result.write_text(json.dumps(result,ensure_ascii=False,indent=2)+"\n");args.reward_file.write_text(f"{result['reward']:.6f}\n");print(json.dumps(result,ensure_ascii=False,indent=2))
if __name__=="__main__":main()
