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
def retarget(path,address):
 box=mailbox.mbox(path,create=False);messages=list(box);box.close();messages[0].replace_header("To",address)
 with path.open("wb") as handle:
  for index,message in enumerate(messages):handle.write(f"From x Wed Apr 15 09:00:{index:02d} 2026\n".encode());BytesGenerator(handle,policy=SMTP,mangle_from_=True,maxheaderlen=0).flatten(message);handle.write(b"\n")
def main():
 with tempfile.TemporaryDirectory(prefix="intern5133-neg-") as temp:
  root=Path(temp);names=["missing","tasks_applied","card_issued","account_provisioned","wrong_welcome","wrong_manager","slack_posted","formula_removed","word_false","ppt_omit"];cases=[]
  for name in names:case=root/name;shutil.copytree(GOLDEN,case);cases.append(case)
  (cases[0]/"Intern_Onboarding_Review.pptx").unlink()
  for case,key in ((cases[1],"onboarding_tasks_applied"),(cases[2],"corporate_cards_issued"),(cases[3],"accounts_provisioned")):
   payload=json.loads((case/"Intern_Execution_Control.json").read_text());payload[key]=1;(case/"Intern_Execution_Control.json").write_text(json.dumps(payload))
  retarget(cases[4]/"Intern_Welcome_Drafts.mbox","amara.blake@company.example.com");retarget(cases[5]/"Manager_Approval_Request_Drafts.mbox","it-servicedesk@company.example.com")
  payload=json.loads((cases[6]/"Intern_Program_Slack_Draft.json").read_text());payload["posted_count"]=1;payload["state"]="POSTED";(cases[6]/"Intern_Program_Slack_Draft.json").write_text(json.dumps(payload))
  workbook=load_workbook(cases[7]/"Intern_Onboarding_Audit.xlsx");workbook["Onboarding_Decisions"]["I2"]="CARD_ISSUED";workbook.save(cases[7]/"Intern_Onboarding_Audit.xlsx")
  document=Document(cases[8]/"Intern_Onboarding_Brief.docx");document.add_paragraph("All setup is complete and cards were issued.");document.save(cases[8]/"Intern_Onboarding_Brief.docx")
  presentation=Presentation(cases[9]/"Intern_Onboarding_Review.pptx")
  for slide in presentation.slides:
   for shape in slide.shapes:
    if getattr(shape,"has_text_frame",False):shape.text=shape.text.replace("Dmitri Volkov","Unknown")
    if getattr(shape,"has_table",False):
     for row in shape.table.rows:
      for cell in row.cells:cell.text=cell.text.replace("Dmitri Volkov","Unknown")
  presentation.save(cases[9]/"Intern_Onboarding_Review.pptx")
  rewards={case.name:score(case) for case in cases};print(json.dumps(rewards,indent=2));assert all(value==0 for value in rewards.values()),rewards
if __name__=="__main__":main()
