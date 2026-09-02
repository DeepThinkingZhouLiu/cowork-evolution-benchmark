#!/usr/bin/env python3
"""AutomationBench 5057 四应用入职一一对账专属 Judge。"""
import argparse,json,mailbox,re
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook
from office_checks import body_text,docx_text,package_safe,pptx_text,render,sha256
TASK="cowork-evo/automationbench-5057-multi-app-new-hire-onboarding-package";HASHES={"Full_Onboarding_Checklist.xlsx":"741c44f29a6a88d724e239a999d722d6f2e34befb165ceffab5d0ac4bd75c6c1","Initial_Application_State.json":"3404ee64b25349823923e4973ae6e4f5ab568e1076df3a0530e3836d68172a02","Workflow_Context.json":"a8a2321ff9af256293a203a899bd5e0140b101f1590d4cd295b34a8c98d1757a"};X="New_Hire_Onboarding_Audit.xlsx";D="New_Hire_Onboarding_Brief.docx";P="New_Hire_Onboarding_Review.pptx";B="BambooHR_Profile_Create_Drafts.json";F="Drive_Folder_Create_Drafts.json";J="Jira_Onboarding_Ticket_Drafts.json";M="New_Hire_Welcome_Drafts.mbox";C="Onboarding_Execution_Control.json";REQUIRED={X,D,P,B,F,J,M,C};WEIGHTS={f"R{i:03d}":v for i,v in enumerate([6,8,10,10,10,8,9,8,8,8,7,8],1)};CAPS={k:(0 if k=="R001" else .25) for k in WEIGHTS if k!="R012"}
class Checks:
 def __init__(self):self.v={k:{"passed":True,"details":[]} for k in WEIGHTS}
 def require(self,k,c,d):
  if not c:self.v[k]["passed"]=False;self.v[k]["details"].append(d) if d not in self.v[k]["details"] else None
 def ok(self,k):return self.v[k]["passed"]
def root():return Path(__file__).resolve().parents[1]
def inp():return Path("/app/input_files") if Path("/app/input_files").is_dir() else root()/"data/input_files"
def rows(ws):return [tuple(c.value for c in r) for r in ws.iter_rows(min_row=2)]
def source(q):
 for n,h in HASHES.items():q.require("R001",(inp()/n).is_file() and sha256(inp()/n)==h,f"输入{n}")
 wb=load_workbook(inp()/"Full_Onboarding_Checklist.xlsx",read_only=True,data_only=True);r=rows(wb["This Week"]);q.require("R002",[(x[1],x[2],x[3],x[4],x[5].strftime("%Y-%m-%d"),x[6],x[7]) for x in r]==[("Alicia Fernandez","alicia.fernandez@company.example.com","Engineering","David Okonkwo","2026-03-24","Austin","Frontend Developer"),("Tyrone Washington","tyrone.washington@company.example.com","Sales","Lisa Wang","2026-03-25","Remote","Account Executive")],"2人全字段")
 s=json.loads((inp()/"Initial_Application_State.json").read_text());q.require("R002",s["bamboohr"]["actions"]=={} and s["jira"]["actions"]=={} and s["google_drive"]["folders"]==[{"id":"folder_employees","name":"Employee Documents"}],"应用初始状态")
def workbook(p,q):
 wb=load_workbook(p,data_only=False);q.require("R003",wb.sheetnames==["Source_Checklist","Source_System_State","Evidence_Register","Hire_Decisions","BambooHR_Drafts","Drive_Drafts","Jira_Drafts","Welcome_Drafts","Manager_Exclusion","Cross_System_Reconciliation","Readiness_Matrix","Control_Summary"],"12表");f=[c.value for r in wb["Hire_Decisions"].iter_rows(min_row=2) for c in r if isinstance(c.value,str) and c.value.startswith("=")];q.require("R004",len(f)==32,f"公式{len(f)}");j="\n".join(f)
 for t in ("Source_Checklist!H3","folder_employees/",'"Onboarding: "','L2=C2','david.okonkwo@company.example.com','lisa.wang@company.example.com','READY_DRAFT_NOT_EXECUTED'):q.require("R004",t in j,f"公式缺{t}")
 rec=rows(wb["Cross_System_Reconciliation"]);q.require("R005",rec==[("Alicia Fernandez",1,1,1,1,"YES"),("Tyrone Washington",1,1,1,1,"YES")],"四系统一一对账")
 mgr=rows(wb["Manager_Exclusion"]);q.require("R006",mgr==[("David Okonkwo",0,0,"Manager field; not a new hire"),("Lisa Wang",0,0,"Manager field; not a new hire")],"经理排除")
def artifacts(a,q):
 b=json.loads((a/B).read_text());p=b.get("profiles",[]);q.require("R007",b.get("draft_count")==2 and b.get("created_count")==0 and [(x["name"],x["email"],x["department"],x["manager"],x["start_date"],x["location"],x["role"],x["state"]) for x in p]==[("Alicia Fernandez","alicia.fernandez@company.example.com","Engineering","David Okonkwo","2026-03-24","Austin","Frontend Developer","DRAFT_NOT_CREATED"),("Tyrone Washington","tyrone.washington@company.example.com","Sales","Lisa Wang","2026-03-25","Remote","Account Executive","DRAFT_NOT_CREATED")],"Bamboo草稿")
 f=json.loads((a/F).read_text());q.require("R007",f.get("draft_count")==2 and f.get("created_count")==0 and [(x["name"],x["parent_id"],x["state"]) for x in f.get("folders",[])]==[("Alicia Fernandez","folder_employees","DRAFT_NOT_CREATED"),("Tyrone Washington","folder_employees","DRAFT_NOT_CREATED")],"Drive草稿")
 j=json.loads((a/J).read_text());ts=j.get("tickets",[]);q.require("R008",j.get("draft_count")==2 and j.get("created_count")==0 and [x["summary"] for x in ts]==["Onboarding: Alicia Fernandez","Onboarding: Tyrone Washington"] and all(x["state"]=="DRAFT_NOT_CREATED" for x in ts) and all(t in ts[0]["description"] for t in ("Frontend Developer","David Okonkwo","2026-03-24","Austin")) and all(t in ts[1]["description"] for t in ("Account Executive","Lisa Wang","2026-03-25","Remote")),"Jira草稿")
 ms=list(mailbox.mbox(a/M,create=False));tos=[str(x.get("To")) for x in ms];texts=[body_text(x) for x in ms];q.require("R008",len(ms)==2 and tos==["alicia.fernandez@company.example.com","tyrone.washington@company.example.com"] and all(t in texts[0] for t in ("Alicia Fernandez","Engineering","Frontend Developer","2026-03-24","Austin","David Okonkwo","DRAFT_NOT_SENT")) and all(t in texts[1] for t in ("Tyrone Washington","Sales","Account Executive","2026-03-25","Remote","Lisa Wang","DRAFT_NOT_SENT")) and not any(x in tos for x in ("david.okonkwo@company.example.com","lisa.wang@company.example.com")),"欢迎邮件只给新人")
 c=json.loads((a/C).read_text());exp={"new_hires":2,"bamboo_profile_drafts":2,"bamboo_profiles_created":0,"drive_folder_drafts":2,"drive_folders_created":0,"jira_ticket_drafts":2,"jira_tickets_created":0,"welcome_email_drafts":2,"welcome_emails_sent":0,"manager_profiles_created":0,"manager_welcome_emails_sent":0,"external_actions_executed":0,"real_side_effects":0};q.require("R009",all(c.get(k)==v for k,v in exp.items()),"执行边界")
def office(a,q):
 text,s,t=docx_text(a/D);q.require("R010",s==3 and t==3,"Word结构")
 for x in ("Alicia Fernandez","Tyrone Washington","David Okonkwo","Lisa Wang","BambooHR","Google Drive","Jira","Gmail","DRAFT_NOT_CREATED"):q.require("R010",x.lower() in text.lower(),f"Word缺{x}")
 q.require("R010","david okonkwo is a new hire" not in text.lower(),"Word不得误分经理")
 text,s,t,pic=pptx_text(a/P);q.require("R011",s==3 and t==[1,1,1] and pic==0,"PPT结构")
 for x in ("Two hires, four systems each","Alicia Fernandez","Tyrone Washington","David Okonkwo","Lisa Wang","DRAFT_NOT_CREATED","DRAFT_NOT_SENT"):q.require("R011",x in text,f"PPT缺{x}")
 for n in (X,D,P):safe,d=package_safe(a/n);q.require("R012",safe,f"安全{n}:{d}")
def finish(q,pages):
 h=[k for k in CAPS if not q.ok(k)];base=sum(v for k,v in WEIGHTS.items() if q.ok(k))/100;cap=min([1]+[CAPS[k] for k in h]);pen=(.3 if any(not q.ok(k) for k in ("R002","R003","R004","R005","R006")) else 0)+(.3 if any(not q.ok(k) for k in ("R007","R008","R009")) else 0)+(.25 if any(not q.ok(k) for k in ("R001","R010","R011","R012")) else 0);reward=round(max(0,min(base,cap)-pen),6);return {"task_id":TASK,"reward":reward,"passed":reward>=.8 and not h,"criteria":q.v,"hurdle_failures":h,"source_truth":{"new_hires":2,"drafts_per_system":2,"manager_actions":0,"external_actions":0},"render_pages":pages}
def evaluate(a):
 q=Checks();source(q);files={p.name for p in a.iterdir() if p.is_file()} if a.is_dir() else set();q.require("R001",files==REQUIRED,f"集合{files}")
 if files!=REQUIRED:return finish(q,{})
 for fn,ks in ((lambda:workbook(a/X,q),("R003","R004","R005","R006","R012")),(lambda:artifacts(a,q),("R007","R008","R009")),(lambda:office(a,q),("R010","R011","R012"))):
  try:fn()
  except Exception as e:
   for k in ks:q.require(k,False,str(e))
 pages={}
 for n,e in ((X,12),(D,3),(P,3)):
  try:c,t=render(a/n);pages[n]=c;q.require("R012",c==e,f"页数{n}:{c}");q.require("R012",not re.search(r"#(?:VALUE|REF|NAME|DIV/0|N/A)",t,re.I),f"公式错{n}")
  except Exception as x:q.require("R012",False,str(x))
 return finish(q,pages)
def main():
 p=argparse.ArgumentParser();p.add_argument("--artifact-dir",type=Path,required=True);p.add_argument("--result",type=Path,required=True);p.add_argument("--reward-file",type=Path,required=True);a=p.parse_args();r=evaluate(a.artifact_dir);a.result.parent.mkdir(parents=True,exist_ok=True);a.reward_file.parent.mkdir(parents=True,exist_ok=True);a.result.write_text(json.dumps(r,ensure_ascii=False,indent=2)+"\n");a.reward_file.write_text(f"{r['reward']:.6f}\n");print(json.dumps(r,ensure_ascii=False,indent=2))
if __name__=="__main__":main()
