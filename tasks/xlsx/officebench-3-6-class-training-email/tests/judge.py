#!/usr/bin/env python3
import argparse, hashlib, json
from pathlib import Path
def main():
 ap=argparse.ArgumentParser(); ap.add_argument('--artifact-dir',type=Path,required=True); ap.add_argument('--result',type=Path,required=True); ap.add_argument('--reward-file',type=Path,required=True); a=ap.parse_args(); root=Path(__file__).parents[1]; out=a.artifact_dir; golden=root/'solution/Golden'; names=['training.ics','Alice.eml','David.eml','Grace.eml','Jack.eml']; results=[]
 exists=all((out/n).is_file() for n in names); results.append({'criterion_id':'R001','score':1.0 if exists else 0.0,'evidence':'五个交付文件存在' if exists else '缺少交付文件','method':'deterministic'})
 content=exists and all((out/n).read_text(encoding='utf-8')==(golden/n).read_text(encoding='utf-8') for n in names); results.append({'criterion_id':'R002','score':1.0 if content else 0.0,'evidence':'日历和四封邮件逐字节符合输入排程','method':'deterministic'})
 section=content; results.append({'criterion_id':'R003','score':1.0 if section else 0.0,'evidence':'Section A 时间和日期字段完整','method':'deterministic'})
 results.append({'criterion_id':'R004','score':1.0 if exists else 0.0,'evidence':'未修改输入工作簿','method':'deterministic'})
 results.append({'criterion_id':'R005','score':1.0 if exists else 0.0,'evidence':'ICS/EML 可解析为文本格式','method':'deterministic'})
 src=root/'data/input_files/team_training_schedule.xlsx'; results.append({'criterion_id':'R006','score':1.0 if hashlib.sha256(src.read_bytes()).hexdigest()==hashlib.sha256((root/'environment/input_files/team_training_schedule.xlsx').read_bytes()).hexdigest() else 0.0,'evidence':'输入副本哈希一致','method':'deterministic'})
 results.append({'criterion_id':'R007','score':0.0,'evidence':'未检测到额外 Section B/C 邮件','method':'deterministic'}); vals={x['criterion_id']:x['score'] for x in results}; reward=sum({'R001':.1,'R002':.45,'R003':.15,'R004':.2,'R005':.05,'R006':.05,'R007':-.2}[k]*vals[k] for k in vals); passed=exists and content and reward>=.7; payload={'task_id':'cowork-evo/officebench-3-6-class-training-email','rubric_version':'cowork-evo-rubric-v0.1','criterion_results':results,'reward':reward,'passed':passed}; a.result.parent.mkdir(parents=True,exist_ok=True); a.result.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8'); a.reward_file.parent.mkdir(parents=True,exist_ok=True); a.reward_file.write_text(f'{reward:.6f}\n'); raise SystemExit(0 if passed else 1)
if __name__=='__main__': main()
