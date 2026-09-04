#!/usr/bin/env python3
import argparse, hashlib, json, re, zipfile
from datetime import datetime
from pathlib import Path
from openpyxl import load_workbook

def digest(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def crit(i,ok,e): return {'criterion_id':f'R{i:03d}','score':1.0 if ok else 0.0,'evidence':e,'method':'deterministic'}
def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--artifact-dir',type=Path,required=True); ap.add_argument('--result',type=Path,required=True); ap.add_argument('--reward-file',type=Path,required=True); a=ap.parse_args()
    root=Path(__file__).parents[1]; inp=root/'data/input_files'; out=a.artifact_dir/'event_plan_v2.xlsx'; rs=[]
    native=out.is_file() and zipfile.is_zipfile(out); rs.append(crit(1,native,'有效 XLSX'))
    if native:
        try:
            wb=load_workbook(out,data_only=True); ws=wb['Event Plan']; vals=list(ws.iter_rows(min_row=2,values_only=True)); ids=[r[1] for r in vals]; headers=[c.value for c in ws[1]]
            md=list(inp.glob('event_plan_*.md')); expected={}
            for p in md:
                t=p.read_text(encoding='utf-8'); n=int(re.search(r'_(\d+)\.md$',p.name).group(1));
                def f(label):
                    m=re.search(rf'\*\*{re.escape(label)}\*\*:\s*(.+)',t); return m.group(1).strip() if m else ''
                lead=re.search(r'Lead generation \((\d+) target\)',t,re.I); mgr=re.search(r'- Event Manager:\s*(.+)',t); mkt=re.search(r'- Marketing Lead:\s*(.+)',t); st=re.search(r'\*\*Status\*\*:\s*(.+)',t)
                expected[n]=[p.name,f'Event {n}',f('Event Name'),datetime.strptime(f('Date'),'%Y-%m-%d'),f('Location'),int(f('Expected Attendees')),int(lead.group(1)),mgr.group(1).strip(),mkt.group(1).strip(),st.group(1).strip()]
            rs += [crit(2,len(vals)==60,f'rows={len(vals)}'),crit(3,headers==['Filename','Event ID','Event Name','Date','Location','Expected Attendees','Lead Generation Target','Event Manager','Marketing Lead','Status'],str(headers)),crit(4,ids==[f'Event {i}' for i in range(1,61)],'Event IDs连续'),crit(5,all(vals[i-1][0:10]==tuple(expected[i]) for i in range(1,61)),'逐行重算与Markdown一致'),crit(6,all(isinstance(r[3],datetime) for r in vals),'日期为日期值'),crit(7,all(isinstance(r[5],int) and r[5]>0 and isinstance(r[6],int) and r[6]>0 for r in vals),'金额字段为正整数'),crit(8,all(r[4]==f'Virtual / City {i}' for i,r in enumerate(vals,1)),'Location格式'),crit(9,all(r[7]==f'Team Lead {i}' and r[8]==f'Marketing Manager {i}' for i,r in enumerate(vals,1)),'负责人格式'),crit(10,all(r[9]=='Active Planning' for r in vals),'Status一致'),crit(11,'Source Notes' in wb.sheetnames,'包含来源说明'),crit(12,all(x.name=='event_plan_v2.xlsx' for x in a.artifact_dir.iterdir()),'无额外输出')]
        except Exception as e: rs += [crit(i,False,f'解析失败: {e}') for i in range(2,13)]
    else: rs += [crit(i,False,'输出缺失或损坏') for i in range(2,13)]
    reward=sum(x['score'] for x in rs)/len(rs); a.result.parent.mkdir(parents=True,exist_ok=True); a.result.write_text(json.dumps({'task_id':'cowork-evo/workspace-event-plan','rubric_version':'workspace-event-plan-v0.1','criteria':rs,'reward':reward},ensure_ascii=False,indent=2)); a.reward_file.parent.mkdir(parents=True,exist_ok=True); a.reward_file.write_text(f'{reward:.6f}\n')
if __name__=='__main__': main()
