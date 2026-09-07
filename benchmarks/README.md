# Benchmark Release 目录

本目录保存已经冻结的 Benchmark Release。Release 只决定评哪些题、题目版本和可见性，不保存 Agent 运行产物，也不实现 Judge。

一个正式 Release 至少包含：

```text
benchmarks/<release-id>/
├── benchmark.json
├── dataset.lock
├── leakage-groups.jsonl
└── splits/
    ├── feedback.jsonl
    ├── selection.jsonl
    └── final.sealed.jsonl
```

Release 发布后不可原地修改。题目、分区、Rubric、Judge 或阈值发生变化时必须创建新版本，并保留旧版本的结果解释能力。

当前 Release：

- [`cowork-evo-240-v0.2`](cowork-evo-240-v0.2/README.md)：当前版本；保持 Feedback 90、Selection 30、Final 120 的原分区归属，收紧 DOCX 确定性 Judge 并更新任务内容锁。
- [`cowork-evo-240-v0.1`](cowork-evo-240-v0.1/README.md)：历史版本；首次冻结 240 道题的三段式划分。

生成与完整性复核入口：

```bash
python tools/freeze_benchmark_release.py --check
```
