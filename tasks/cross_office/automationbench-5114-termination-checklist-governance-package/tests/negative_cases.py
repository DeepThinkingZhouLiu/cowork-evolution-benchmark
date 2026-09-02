#!/usr/bin/env python3
import json,mailbox,shutil,subprocess,tempfile
from email.generator import BytesGenerator
from email.policy import SMTP
from pathlib import Path
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
ROOT=Path(__file__).resolve().parents[1];GOLDEN=ROOT/"solution/Golden";PYTHON="REMOVED_LOCAL_PATH"
def score(case):
 result=case.parent/(case.name+".json");reward=case.parent/(case.name+".txt");subprocess.run([PYTHON,str(ROOT/"tests/judge.py"),"--artifact-dir",str(case),"--result",str(result),"--reward-file",str(reward)],check=True,capture_output=True,text=True,timeout=240);return float(reward.read_text())
def rewrite(path,mutation):
 box=mailbox.mbox(path,create=False);messages=list(box);box.close();mutation(messages)
 with path.open("wb") as handle:
  for index,message in enumerate(messages):handle.write(f"From x Wed Apr 15 09:00:{index:02d} 2026\n".encode());BytesGenerator(handle,policy=SMTP,mangle_from_=True,maxheaderlen=0).flatten(message);handle.write(b"\n")
def main():
 with tempfile.TemporaryDirectory(prefix="termination5114-neg-") as temp:
  root=Path(temp);names=["missing","wrong_process","dina_processed","isaac_processed","pto_missing","paycheck_issued","revocation_done","email_sent","slack_posted","formula_removed","word_false","ppt_omit"];cases=[]
  for name in names:case=root/name;shutil.copytree(GOLDEN,case);cases.append(case)
  (cases[0]/"Termination_Checklist_Review.pptx").unlink()
  payload=json.loads((cases[1]/"Termination_Execution_Control.json").read_text());payload["process_checklist"]=8;(cases[1]/"Termination_Execution_Control.json").write_text(json.dumps(payload))
  payload=json.loads((cases[2]/"Termination_Tracker_Update_Drafts.json").read_text());payload["rows"].append({"source_row":5,"employee":"Dina Petrova","from":"Processed","to":"Processed"});payload["draft_count"]=8;(cases[2]/"Termination_Tracker_Update_Drafts.json").write_text(json.dumps(payload))
  rewrite(cases[3]/"Termination_Communication_Drafts.mbox",lambda ms:ms[0].set_payload(ms[0].get_payload()+"\nIsaac Reeves"))
  rewrite(cases[4]/"Termination_Communication_Drafts.mbox",lambda ms:ms[4].set_payload(ms[4].get_payload().replace("PTO balance unknown","")))
  payload=json.loads((cases[5]/"Termination_Execution_Control.json").read_text());payload["final_paychecks_issued"]=7;(ases:=cases[5]/"Termination_Execution_Control.json").write_text(json.dumps(payload))
  payload=json.loads((cases[6]/"Termination_Execution_Control.json").read_text());payload["access_revocations_executed"]=7;(cases[6]/"Termination_Execution_Control.json").write_text(json.dumps(payload))
  payload=json.loads((cases[7]/"Termination_Execution_Control.json").read_text());payload["emails_sent"]=21;(cases[7]/"Termination_Execution_Control.json").write_text(json.dumps(payload))
  payload=json.loads((cases[8]/"HR_Ops_Termination_Slack_Drafts.json").read_text());payload["posted_count"]=7;payload["state"]="POSTED";(cases[8]/"HR_Ops_Termination_Slack_Drafts.json").write_text(json.dumps(payload))
  workbook=load_workbook(cases[9]/"Termination_Checklist_Audit.xlsx");workbook["Checklist_Decisions"]["M2"]="PROCESS_TERMINATION_CHECKLIST";workbook.save(cases[9]/"Termination_Checklist_Audit.xlsx")
  document=Document(cases[10]/"Termination_Checklist_Brief.docx");document.add_paragraph("Access was revoked, final pay was issued, and source rows were changed.");document.save(cases[10]/"Termination_Checklist_Brief.docx")
  presentation=Presentation(cases[11]/"Termination_Checklist_Review.pptx")
  for slide in presentation.slides:
   for shape in slide.shapes:
    if getattr(shape,"has_text_frame",False):shape.text=shape.text.replace("Isaac","Unknown")
    if getattr(shape,"has_table",False):
     for row in shape.table.rows:
      for cell in row.cells:cell.text=cell.text.replace("Isaac","Unknown")
  presentation.save(cases[11]/"Termination_Checklist_Review.pptx")
  rewards={case.name:score(case) for case in cases};print(json.dumps(rewards,indent=2));assert all(value==0 for value in rewards.values()),rewards
if __name__=="__main__":main()
