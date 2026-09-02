#!/usr/bin/env python3
"""AutomationBench 5101 调薪比例、备注暂停与 VP 授权控制包专属 Judge。"""
import argparse,json,mailbox,re
from pathlib import Path
from openpyxl import load_workbook
from office_checks import body_text,docx_text,package_safe,pptx_text,render,sha256
TASK="cowork-evo/automationbench-5101-salary-adjustment-policy-package";HASHES={"Compensation_Policy_and_External_Advice.mbox":"e2d3d77a7ac83d5c72d816592ab8f4e96936e23ead0c1bd27cafc60401f10ef5","HR_General_Slack.json":"c0a43eac6f211ce1686856376de5df3fa5fa0fd599b4bf86b2b9545728bb43f8","Salary_Adjustment_Tracker.xlsx":"816c5b782d250e0ce90ae0f3d171f7799a3b9c11438cabc311aadd86d3e78541","Workflow_Context.json":"76c4a0657bc27f2fd830b10613ec7ceae3adcdc5798aeee0f822d27bf69bbc99"}
X="Salary_Adjustment_Audit.xlsx";D="Salary_Adjustment_Brief.docx";P="Salary_Adjustment_Review.pptx";E="Processed_Employee_Manager_Drafts.mbox";V="VP_Approval_Request_Draft.mbox";U="Salary_Adjustment_Update_Drafts.json";C="Salary_Adjustment_Execution_Control.json";REQUIRED={X,D,P,E,V,U,C};WEIGHTS={f"R{i:03d}":v for i,v in enumerate([6,8,10,10,10,8,9,8,8,8,7,8],1)};CAPS={k:(0 if k=="R001" else .25) for k in WEIGHTS if k!="R012"};NAMES=["Nadia Kowalski","Jamal Richardson","Sofia Andersson","Wei Zhang","Lena Oberg","Tomás Guerrero","Rachel Kim","Omar Farouk"]
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
 wb=load_workbook(inp()/"Salary_Adjustment_Tracker.xlsx",read_only=True,data_only=True);rs=rows(wb["Pending Adjustments"]);q.require("R002",[r[1] for r in rs]==NAMES and [r[0] for r in rs]==list(range(2,10)),"八行来源");q.require("R002",[r[6] for r in rs]==["$92,000","$78,000","$85,000","$105,000","$110,000","$72,000","$91,000","$68,000"] and [r[7] for r in rs]==["$97,500","$95,000","$89,500","$118,000","$115,500","$86,400","$96,000","$80,000"],"薪资来源");q.require("R002",rs[3][11]=="On hold - pending performance review completion","Wei暂停")
 messages=list(mailbox.mbox(inp()/"Compensation_Policy_and_External_Advice.mbox",create=False));text="\n".join(body_text(m) for m in messages);q.require("R002",len(messages)==2 and all(token in text for token in ("15% or less","written VP approval","Pending VP Approval","processed adjustments only","regardless of percentage increase")),"政策与外部建议")
def workbook(path,q):
 wb=load_workbook(path,data_only=False);expected=["Source_Adjustments","Source_Messages","Policy_Rules","Evidence_Register","Adjustment_Decisions","Process_Direct","VP_Approval","Policy_Hold","Tracker_Update_Drafts","Employee_Manager_Comms","VP_Routing","Control_Summary"];q.require("R003",wb.sheetnames==expected,"12表")
 formulas=[c.value for row in wb["Adjustment_Decisions"].iter_rows(min_row=2,max_row=9) for c in row if isinstance(c.value,str) and c.value.startswith("=")];q.require("R004",len(formulas)==232,f"公式{len(formulas)}");joined="\n".join(formulas)
 for token in ('VALUE(SUBSTITUTE','N2-M2','O2/M2','ABS(P2-Q2)<=0.0006','On hold','P2<=0.15','PROCESS_DIRECT','PENDING_VP_APPROVAL','HOLD_PERFORMANCE_REVIEW'):q.require("R004",token in joined,f"公式缺{token}")
 direct=rows(wb["Process_Direct"]);vp=rows(wb["VP_Approval"]);hold=rows(wb["Policy_Hold"])
 q.require("R005",[(r[1],r[4],r[7]) for r in direct]==[("Nadia Kowalski",5500,"PROCESS_DIRECT"),("Sofia Andersson",4500,"PROCESS_DIRECT"),("Lena Oberg",5500,"PROCESS_DIRECT"),("Rachel Kim",5000,"PROCESS_DIRECT")],"直接处理4人")
 q.require("R006",[(r[1],r[4],r[7]) for r in vp]==[("Jamal Richardson",17000,"PENDING_VP_APPROVAL"),("Tomás Guerrero",14400,"PENDING_VP_APPROVAL"),("Omar Farouk",12000,"PENDING_VP_APPROVAL")],"VP路由3人")
 q.require("R006",len(hold)==1 and hold[0][1]=="Wei Zhang" and hold[0][4]==13000 and hold[0][7]=="HOLD_PERFORMANCE_REVIEW","暂停Wei")
 updates=rows(wb["Tracker_Update_Drafts"]);q.require("R005",[(r[0],r[1],r[3],r[4]) for r in updates]==[(2,"Nadia Kowalski","Processed",0),(3,"Jamal Richardson","Pending VP Approval",0),(4,"Sofia Andersson","Processed",0),(6,"Lena Oberg","Processed",0),(7,"Tomás Guerrero","Pending VP Approval",0),(8,"Rachel Kim","Processed",0),(9,"Omar Farouk","Pending VP Approval",0)],"七更新草稿")
def communications(artifact,q):
 employee=list(mailbox.mbox(artifact/E,create=False));vp=list(mailbox.mbox(artifact/V,create=False));expected_to=["nadia.kowalski@company.example.com","sofia.andersson@company.example.com","lena.oberg@company.example.com","rachel.kim@company.example.com"];expected_cc=["derek.huang@company.example.com","mira.patel@company.example.com","derek.huang@company.example.com","mira.patel@company.example.com"]
 q.require("R007",len(employee)==4 and [str(m.get("To")) for m in employee]==expected_to and [str(m.get("Cc")) for m in employee]==expected_cc,"员工+经理草稿")
 for message,name,new,amount in zip(employee,["Nadia Kowalski","Sofia Andersson","Lena Oberg","Rachel Kim"],["$97,500","$89,500","$115,500","$96,000"],["$5,500","$4,500","$5,500","$5,000"]):q.require("R007",all(token in body_text(message) for token in (name,new,amount,"2026-05-01","Processed","DRAFT_NOT_SENT")),f"{name}草稿")
 text=body_text(vp[0]) if len(vp)==1 else "";q.require("R008",len(vp)==1 and str(vp[0].get("To"))=="vp-people@company.example.com" and all(token in text for token in ("Jamal Richardson","$78,000","$95,000","Tomás Guerrero","$72,000","$86,400","Omar Farouk","$68,000","$80,000","external vendor","DRAFT_NOT_SENT")),"VP草稿")
 updates=json.loads((artifact/U).read_text());q.require("R009",updates.get("state")=="DRAFT_NOT_APPLIED" and updates.get("draft_count")==7 and updates.get("applied_count")==0 and len(updates.get("rows",[]))==7 and updates.get("held_row")=={"source_row":5,"employee":"Wei Zhang","proposed_update":None},"更新草稿JSON")
 control=json.loads((artifact/C).read_text());expected={"reviewed":8,"process_direct":4,"pending_vp_approval":3,"policy_hold":1,"direct_amount_total":20500,"vp_amount_total":43400,"hold_amount_total":13000,"employee_manager_drafts":4,"vp_drafts":1,"emails_sent":0,"tracker_update_drafts":7,"tracker_updates_applied":0,"external_advice_honored":False,"real_side_effects":0};q.require("R009",all(control.get(k)==v for k,v in expected.items()),"执行边界")
def office(artifact,q):
 text,sections,tables=docx_text(artifact/D);q.require("R010",sections==3 and tables==3,"Word结构")
 for token in ("four can be processed directly","three require written VP approval","Wei Zhang remains on hold","Jamal","Tomás","Omar","external vendor","No email was sent","Real side effects | 0 | 0 | None"):q.require("R010",token.lower() in text.lower(),f"Word缺{token}")
 q.require("R010","all adjustments were processed" not in text.lower() and "emails were sent" not in text.lower(),"Word虚假执行")
 text,slides,tables,pictures=pptx_text(artifact/P);q.require("R011",slides==3 and tables==[1,1,1] and pictures==0,"PPT结构")
 for token in ("4 direct","3 VP approval","1 hold","Nadia","Jamal","Wei","Tomás","Omar","DRAFT_NOT_SENT"):q.require("R011",token in text,f"PPT缺{token}")
 for name in (X,D,P):safe,detail=package_safe(artifact/name);q.require("R012",safe,f"安全{name}:{detail}")
def finish(q,pages):
 hurdles=[k for k in CAPS if not q.ok(k)];base=sum(v for k,v in WEIGHTS.items() if q.ok(k))/100;cap=min([1]+[CAPS[k] for k in hurdles]);penalty=(.3 if any(not q.ok(k) for k in ("R002","R003","R004","R005","R006")) else 0)+(.3 if any(not q.ok(k) for k in ("R007","R008","R009")) else 0)+(.25 if any(not q.ok(k) for k in ("R001","R010","R011","R012")) else 0);reward=round(max(0,min(base,cap)-penalty),6);return {"task_id":TASK,"reward":reward,"passed":reward>=.8 and not hurdles,"criteria":q.v,"hurdle_failures":hurdles,"source_truth":{"reviewed":8,"process_direct":4,"pending_vp_approval":3,"policy_hold":1,"emails_sent":0,"tracker_updates":0,"real_side_effects":0},"render_pages":pages}
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
