# 数据资产目录

本目录只保存数据清单、版本锁、来源记录、许可状态、隐私状态、转换报告和泄漏组，不保存真实企业大文件。

建议的数据生命周期为：

```text
原始数据 -> 候选数据 -> 制题数据 -> 校准数据 -> 正式 Release
```

每个数据版本至少应包含：

- `manifest.jsonl`：逐任务来源、文件哈希和当前状态。
- `dataset.lock`：不可变版本、清单哈希和外部存储标识。
- `leakage-groups.jsonl`：重复文件、近重复指令和同源模板的分组。
- `conversion-report.json`：转换工具版本、输入输出数量和失败原因。
- `privacy-report.json`：隐私审查范围、结论和发布限制。

私有数据默认放在被 Git 忽略的 `datasets/private/`，正式代码只能通过配置的数据根目录读取，不能写死个人机器的绝对路径。

当前公开参考源的许可证与可用性审计见 [`reference-source-audit.md`](reference-source-audit.md)。Excel 数据分成两个互不混淆的集合：

- [`xlsx-baseline-audit.csv`](xlsx-baseline-audit.csv)：现有 14 道 Excel 题的人工审计基线。
- [`xlsx-candidate-inventory.csv`](xlsx-candidate-inventory.csv)：本轮新筛出的 97 道上游候选，其中 `ready=70`、`conditional=27`；其中 27 道已使用制题 Skill 转换为 `pending_human_review` 的 Harbor 包，进度见 [`xlsx-rollout-status.md`](xlsx-rollout-status.md)。

按 42 个任务家族整理的全 Office 候选和逐题进度仍见 [`task-candidate-map.md`](task-candidate-map.md)。所有候选表都只是人工工作队列，不代表任务已经下载、改造或进入正式 Release。
