#!/usr/bin/env python3
import json,mailbox,shutil,subprocess,tempfile
from email.generator import BytesGenerator
from email.policy import SMTP
from pathlib import Path
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
R=Path(__file__).resolve().parents[1];G=R/"solution/Golden";PY="REMOVED_LOCAL_PATH"
def score(c):
 o=c.parent/(c.name+'.json');w=c.parent/(c.name+'.txt');subprocess.run([PY,str(R/'tests/judge.py'),'--artifact-dir',str(c),'--result',str(o),'--reward-file',str(w)],check=True,capture_output=True,text=True,timeout=240);return float(w.read_text())
def retarget(p):
 b=mailbox.mbox(p,create=False);ms=list(b);b.close();ms[0].replace_header('To','alice.park@company.example.com')
 with p.open('wb') as h:
  for i,m in enumerate(ms):h.write(f'From x Wed Apr 15 09:00:{i:02d} 2026\n'.encode());BytesGenerator(h,policy=SMTP,mangle_from_=True,maxheaderlen=0).flatten(m);h.write(b'\n')
def main():
 with tempfile.TemporaryDirectory(prefix='hb5131-neg-') as t:
  root=Path(t);names=['missing','wrong_counts','elena_escalated','wrong_employee','wrong_hr','slack_posted','formula_removed','word_false','ppt_omit'];cs=[]
  for n in names:c=root/n;shutil.copytree(G,c);cs.append(c)
  (cs[0]/'Handbook_Review.pptx').unlink()
  p=json.loads((cs[1]/'Handbook_Execution_Control.json').read_text());p['escalated']=8;(cs[1]/'Handbook_Execution_Control.json').write_text(json.dumps(p))
  p=json.loads((cs[2]/'Handbook_Execution_Control.json').read_text());p['leave_exempt']=0;p['escalated']=5;(cs[2]/'Handbook_Execution_Control.json').write_text(json.dumps(p))
  retarget(cs[3]/'Handbook_Employee_Manager_Drafts.mbox');retarget(cs[4]/'Handbook_HR_Director_Escalation_Draft.mbox')
  p=json.loads((cs[5]/'Handbook_Slack_Summary_Draft.json').read_text());p['posted_count']=1;p['state']='POSTED';(cs[5]/'Handbook_Slack_Summary_Draft.json').write_text(json.dumps(p))
  w=load_workbook(cs[6]/'Handbook_Compliance_Audit.xlsx');w['Compliance_Decisions']['N3']='COMPLIANT_ON_TIME';w.save(cs[6]/'Handbook_Compliance_Audit.xlsx')
  d=Document(cs[7]/'Handbook_Control_Brief.docx');d.add_paragraph('All eight employees are compliant and messages were sent.');d.save(cs[7]/'Handbook_Control_Brief.docx')
  p=Presentation(cs[8]/'Handbook_Review.pptx')
  for s in p.slides:
   for sh in s.shapes:
    if getattr(sh,'has_text_frame',False):sh.text=sh.text.replace('Femi Adeyemi','Unknown')
    if getattr(sh,'has_table',False):
     for row in sh.table.rows:
      for cell in row.cells:cell.text=cell.text.replace('Femi Adeyemi','Unknown')
  p.save(cs[8]/'Handbook_Review.pptx')
  x={c.name:score(c) for c in cs};print(json.dumps(x,indent=2));assert all(v==0 for v in x.values()),x
if __name__=='__main__':main()
