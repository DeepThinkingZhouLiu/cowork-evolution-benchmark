"""Excel七题重测的原始分数汇总；不裁剪负分，不改正式难度标签。"""
import csv,json,math,tomllib
from pathlib import Path

def save(path,value):
    path.parent.mkdir(parents=True,exist_ok=True)
    tmp=path.with_suffix(path.suffix+'.tmp');tmp.write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n');tmp.replace(path)

def summarize(root,output,baseline,fresh):
    rows=baseline+fresh
    values=[r['reward'] for r in rows]
    summary={'valid_count':len(rows),'expected_count':60,'fresh_count':len(fresh),'reused_count':len(baseline),
             'reward_sum':sum(values),'mean_reward':sum(values)/len(values) if values else None,
             'full_score_count':sum(abs(x-1)<1e-8 for x in values),'nonzero_count':sum(x>0 for x in values),
             'negative_count':sum(x<0 for x in values),'pending_count':60-len(rows)}
    save(output/'results.json',{'summary':summary,'tasks':sorted(rows,key=lambda r:r['task'])})
    if len(rows)!=60:return summary
    if len({r['task'] for r in rows})!=60:raise ValueError('题目重复')
    for r in rows:
        if not math.isfinite(r['reward']):raise ValueError('非有限分数')
        m=tomllib.loads((root/'tasks/xlsx'/r['task']/'task.toml').read_text())['metadata']
        r.update(estimated_minutes=m['estimated_minutes'],task_family=m['task_family'])
    rows.sort(key=lambda r:(-round(r['reward'],12),r['estimated_minutes'],r['task']))
    tiers=[];offset=0
    for tier,count in [('Easy',12),('Medium',18),('Hard',18),('Expert',12)]:
        selected=rows[offset:offset+count]
        for rank,row in enumerate(selected,offset+1):row.update(rank=rank,pilot_tier=tier)
        tiers.append({'tier':tier,'count':count,'mean_reward':sum(r['reward'] for r in selected)/count,
                      'min_reward':min(r['reward'] for r in selected),'max_reward':max(r['reward'] for r in selected),
                      'full_score_count':sum(abs(r['reward']-1)<1e-8 for r in selected)})
        offset+=count
    ties=[k for k in (12,30,48) if abs(rows[k-1]['reward']-rows[k]['reward'])<1e-12]
    report={'method':'Claude单轮原始Reward降序，固定12/18/18/12；同分按已有预计工作量升序、task ID稳定排序。',
            'scope':'单Harness试分档，不等于三Harness Panel或重复成功率；不修改正式task.toml标签。',
            'negative_policy':'保留历史Judge的负分，不裁剪、不归一化；不把1-Reward解释为失败概率。',
            'boundary_ties':ties,'summary':summary,'tiers':tiers,'tasks':rows}
    save(output/'difficulty.json',report)
    fields=['rank','task','pilot_tier','reward','origin','estimated_minutes','task_family','result_path','judge_result_path']
    with (output/'difficulty.csv').open('w',newline='',encoding='utf-8-sig') as f:
        writer=csv.DictWriter(f,fields,extrasaction='ignore');writer.writeheader();writer.writerows(rows)
    lines=['Excel 60题：Claude Code 2.1.251 / Sonnet 5 单轮试分档','',report['method'],'',
           '| 难度   | 题数 | Claude 平均得分 | 最低得分 | 最高得分 | 满分题 |',
           '| ------ | ---: | --------------: | -------: | -------: | -----: |']
    for t in tiers:lines.append(f"| {t['tier']:<6} | {t['count']:4} | {t['mean_reward']*100:13.2f}% | {t['min_reward']*100:7.2f}% | {t['max_reward']*100:7.2f}% | {t['full_score_count']:6} |")
    lines.extend([f"| 全部   |   60 | {summary['mean_reward']*100:13.2f}% | {min(values)*100:7.2f}% | {max(values)*100:7.2f}% | {summary['full_score_count']:6} |",'',
                  '7道新题重新作答，53道以相同任务内容哈希复用已核验结果。负分按原Judge保留；不是裁剪后的成功率。',
                  f'同分跨档边界：{ties}；此处不同标签不代表实测差异。'])
    (output/'difficulty.md').write_text('\n'.join(lines)+'\n')
    return summary
