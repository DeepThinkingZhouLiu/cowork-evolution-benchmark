"""只新测七道升级Excel，哈希核验后复用其余53道，逐题串行。"""
import argparse,fcntl,json,math,subprocess,sys
from datetime import datetime,timezone
from pathlib import Path
import freeze_benchmark_release as release
from xlsx_hardening_results import save,summarize

def main():
    p=argparse.ArgumentParser();p.add_argument('--runner-root',type=Path,required=True);p.add_argument('--old-results',type=Path,required=True);p.add_argument('--output-root',type=Path,required=True);p.add_argument('--prepare-only',action='store_true');a=p.parse_args()
    root=Path(__file__).resolve().parents[1];out=a.output_root.resolve();out.mkdir(parents=True,exist_ok=True)
    lock=(out/'worker.lock').open('a');fcntl.flock(lock,fcntl.LOCK_EX|fcntl.LOCK_NB)
    sys.path.insert(0,str(a.runner_root.resolve()))
    import run_office_module60 as office
    from xlsx_verifier_recovery import ensure_image,offline_judge
    frozen=office.frozen;office.configure_snapshot(root,'xlsx');frozen.SMOKE_TASK='workspace-payroll-reconciliation'
    prior=json.loads((a.old_results/'results.json').read_text());old={r['task']:r for r in prior['tasks']}
    manifest=json.loads((a.old_results/'experiment-manifest.json').read_text())
    if frozen.sha256(frozen.CLAUDE_BINARY)!=manifest['runtime_sha256']:raise RuntimeError('CLI二进制不等同历史实验')
    if (manifest['model'],manifest['max_turns'],manifest['reasoning_effort'])!=('claude-sonnet-5',80,'xhigh'):raise RuntimeError('历史配置不匹配')
    release.validate_release(root,root/'benchmarks/cowork-evo-240-v0.4')
    old_tasks={r['instance_id']:r for r in json.loads((root/'benchmarks/cowork-evo-240-v0.3/dataset.lock').read_text())['tasks']}
    tasks=release._discover_tasks(root);current={t.slug:t for t in tasks if t.domain=='xlsx'}
    mapping=json.loads((root/'docs/xlsx-hardening-7-replacements.json').read_text())['replacements']
    changed={r['new_instance_id'].split('/',1)[1] for r in mapping};baseline=[]
    for slug,t in current.items():
        if slug in changed:continue
        if old_tasks[t.instance_id]['content_sha256']!=t.content_sha256:raise RuntimeError('未改题哈希变化：'+slug)
        row=old[slug]
        if row['status']!='scored' or not math.isfinite(row['reward']):raise RuntimeError('历史结果不可用：'+slug)
        for field in ('result_path','judge_result_path'):
            if not Path(row[field]).is_file():raise RuntimeError('历史证据缺失：'+slug)
        score=json.loads(Path(row['judge_result_path']).read_text()).get('reward')
        if score is None or abs(score-row['reward'])>1e-6:raise RuntimeError('历史评分不一致：'+slug)
        baseline.append({**row,'origin':'reused_identical_task_hash','content_sha256':t.content_sha256})
    if len(baseline)!=53 or len(changed)!=7:raise RuntimeError('必须是53份复用+7次新答题')
    tree=subprocess.check_output(['git','rev-parse','HEAD:tasks'],cwd=root,text=True).strip()
    def assert_snapshot():
        if subprocess.check_output(['git','rev-parse','HEAD:tasks'],cwd=root,text=True).strip()!=tree or subprocess.check_output(['git','status','--porcelain','--','tasks'],cwd=root,text=True).strip():raise RuntimeError('冻结题目变化，停止付费试验')
    assert_snapshot()
    plan={'task_tree':tree,'release':'cowork-evo-240-v0.4','fresh_tasks':sorted(changed),'reused_count':53,
          'model':'claude-sonnet-5','cli_version':'2.1.251','max_turns':80,'reasoning_effort':'xhigh',
          'runtime_sha256':manifest['runtime_sha256'],'provider':frozen.PROVIDER_BASE_URL,
          'credential_source':str(frozen.SETTINGS),'concurrent':1,'difficulty_counts':[12,18,18,12],
          'policy':'每题首次有效作答；基础设施失败最多3次尝试；能力失败不重采样；不裁剪负分。'}
    if (out/'evaluation-plan.json').exists() and json.loads((out/'evaluation-plan.json').read_text())!=plan:raise RuntimeError('已有计划不匹配')
    save(out/'evaluation-plan.json',plan);frozen.write_manifest(out,'claude');office.amend_manifest(out,'xlsx',warm_build_cache=False,build_proxy=None)
    if a.prepare_only:
        summarize(root,out,baseline,[]);print('预检通过：53份可复用，仅7次新答题',flush=True);return
    from docx_panel_results import active_harbor_pids
    if active_harbor_pids(out):raise RuntimeError('已有本实验Harbor进程，禁止重复作答')
    environment=office.build_environment(frozen.load_provider_secret(frozen.SETTINGS));fresh=[]
    def record(job):
        rec=frozen.attempt_record(job);p=frozen.result_file(job)
        if p is None:return None
        payload=json.loads(p.read_text());kind=(payload.get('exception_info') or {}).get('exception_type')
        if not payload.get('finished_at') or kind not in (None,'AgentTimeoutError'):return None
        events=[]
        for log in p.parent.glob('agent/claude-code.txt'):
            for line in log.read_text(errors='replace').splitlines():
                try:event=json.loads(line)
                except json.JSONDecodeError:continue
                if isinstance(event,dict):events.append(event)
        finals=[e for e in events if e.get('type')=='result']
        if finals and finals[-1].get('is_error') and finals[-1].get('subtype') not in ('error_max_turns','error_max_budget_usd'):return None
        if rec['infrastructure_failure']:
            actual=any(e.get('type')=='assistant' for e in events)
            if kind!='AgentTimeoutError' or not actual or rec['reward'] is None:return None
        return rec
    for index,task in frozen.ordered_tasks():
        if task.name not in changed:continue
        assert_snapshot()
        save(out/'status.json',{'state':'running','current_task':task.name,'fresh_completed':len(fresh),'updated_at':datetime.now(timezone.utc).isoformat()})
        candidate_record=None
        for job in sorted((out/'trials').glob(f'{index:02d}-{task.name}-a*')):
            candidate_record=record(job)
            if candidate_record is not None:break
        while candidate_record is None and frozen.latest_attempt(out,index,task.name)<3:
            frozen.run_one('claude',out,index,task,environment,force=True)
            job=out/'trials'/f'{index:02d}-{task.name}-a{frozen.latest_attempt(out,index,task.name):02d}'
            candidate_record=record(job)
        if candidate_record is None:
            summarize(root,out,baseline,fresh);save(out/'status.json',{'state':'infrastructure_failure','task':task.name,'fresh_completed':len(fresh)});raise RuntimeError('三次基础设施尝试未完成，禁止把缺项记0')
        source=Path(candidate_record['result_path']);evidence=out/'rejudge'/task.name
        image=ensure_image(task,evidence)
        judged=offline_judge(task,source.parent/'artifacts/app/output',evidence/'candidate',image)
        if abs(judged['reward']-candidate_record['reward'])>1e-6:raise RuntimeError('Harbor与独立重判分数不一致：'+task.name)
        fresh.append({'task':task.name,'origin':'fresh_v0.4_rollout','reward':judged['reward'],
                      'content_sha256':current[task.name].content_sha256,'result_path':str(source),
                      'judge_result_path':str(evidence/'candidate/judge-result.json'),'agent_exception':candidate_record['exception'],
                      'input_tokens':candidate_record['input_tokens'],'output_tokens':candidate_record['output_tokens']})
        summary=summarize(root,out,baseline,fresh);print('PROGRESS',len(fresh),'/7','score',judged['reward'],flush=True)
    assert_snapshot();save(out/'status.json',{'state':'completed','fresh_completed':7,'summary':summary});print(json.dumps(summary),flush=True)

if __name__=='__main__':main()
