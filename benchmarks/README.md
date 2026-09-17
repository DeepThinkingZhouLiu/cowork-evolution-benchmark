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

- [`cowork-evo-240-v0.5`](cowork-evo-240-v0.5/README.md)：当前版本；四模块240道题全部回填难度，统一重建内容锁和split清单，逐题分区继承v0.4。
- [`cowork-evo-240-v0.4`](cowork-evo-240-v0.4/README.md)：历史Excel升级版本。其后一次元数据更新只更新了内容锁、未同步split哈希；保留作历史记录，当前请使用v0.5。
- [`cowork-evo-240-v0.3`](cowork-evo-240-v0.3/README.md)：历史Word升级版本。
- [`cowork-evo-240-v0.2`](cowork-evo-240-v0.2/README.md)：历史Judge收紧版本；保持Feedback 90、Selection 30、Final 120的数量，更新DOCX确定性Judge与任务内容锁。
- [`cowork-evo-240-v0.1`](cowork-evo-240-v0.1/README.md)：历史版本；首次冻结 240 道题的三段式划分。

统一逐题难度、三Harness得分与来源SHA256见 [difficulty-labels-v0.5.csv](difficulty-labels-v0.5.csv)。状态为单轮Panel校准；标签已回填，不等同于独立跨模型验证已经完成。

完整性复核入口（默认v0.5，只读）：

```bash
python tools/freeze_benchmark_release.py --check
python tools/sync_difficulty_labels.py
```
