"""冻结七道Excel替换后的v0.4，全部题保持原分区，拒绝来源跨区。"""
import json,subprocess,tempfile
from pathlib import Path
import freeze_benchmark_release as release

def main():
    root=Path(__file__).resolve().parents[1];target=root/'benchmarks/cowork-evo-240-v0.4'
    if target.exists():raise FileExistsError('禁止覆盖已有Release')
    if subprocess.check_output(['git','status','--porcelain','--','tasks'],cwd=root,text=True).strip():raise RuntimeError('先提交任务，再冻结内容锁')
    mapping=json.loads((root/'docs/xlsx-hardening-7-replacements.json').read_text())['replacements']
    previous=release._load_frozen_assignments(root/'benchmarks/cowork-evo-240-v0.3')
    for pair in mapping:previous[pair['new_instance_id']]=previous.pop(pair['old_instance_id'])
    tasks=release._discover_tasks(root);components=release._build_components(tasks)
    if set(previous)!={t.instance_id for t in tasks}:raise ValueError('替换清单与240题不一致')
    for component in components:
        if len({previous[tasks[i].instance_id] for i in component})!=1:raise ValueError('新来源出现跨分区泄漏，不能自动发布')
    with tempfile.TemporaryDirectory(prefix='xlsx-release-mapping-') as temporary:
        staging=Path(temporary);(staging/'splits').mkdir()
        for partition,filename in release.SPLIT_FILENAMES.items():
            release._write_jsonl(staging/'splits'/filename,({'instance_id':k} for k in sorted(previous) if previous[k]==partition))
        release.generate_release(root,target,target.name,source_release_root=staging)
    release._write_json(target/'migration.json',{'source_release':'cowork-evo-240-v0.3','replacements':mapping,'moved_unchanged_tasks':0,'policy':'七题一一继承旧题分区；其他233题不变；不使用Agent分数划分训练/验证/测试。','old_scores_reusable_for_replacements':False})
    release.validate_release(root,target)

if __name__=='__main__':main()
