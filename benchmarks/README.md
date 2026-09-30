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

Release 发布后不可原地修改。题目、分区、Rubric、Judge 或阈值发生变化时必须创建新版本，并在外部实验归档中保留旧版本的结果解释能力。

当前 Release：

- [`cowork-evo-240-v0.5`](cowork-evo-240-v0.5/README.md)：当前公开版本；四模块 240 道题全部回填难度，统一重建内容锁和 split 清单。

旧版本结果保存在项目外部的实验归档中，不随公开仓库发布。

统一逐题难度、三Harness得分与来源SHA256见 [difficulty-labels-v0.5.csv](difficulty-labels-v0.5.csv)。状态为单轮Panel校准；标签已回填，不等同于独立跨模型验证已经完成。

完整性复核入口（默认v0.5，只读）：

```bash
python tools/freeze_benchmark_release.py --check
python tools/sync_difficulty_labels.py
```
