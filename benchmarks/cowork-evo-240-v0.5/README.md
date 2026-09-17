# CoworkEvoBench 240 题三段式划分 v0.5

本目录冻结 240 道 Harbor Task 的分区归属，不复制题目文件。`tasks/` 是唯一题目源；本目录只保存 Release、任务哈希、泄漏组件和 split 清单。

本版继承v0.4全部逐题分区，只统一回填难度元数据并重建内容哈希；题目输入、环境、Golden、Rubric、Judge和历史分数均不变。统一难度清单见 [difficulty-labels-v0.5.csv](../difficulty-labels-v0.5.csv)，变更记录见 [migration.json](migration.json)。四模块各为easy 12、medium 18、hard 18、expert 12，校准状态为Sonnet 5三Harness单轮Panel，不代表已完成独立跨模型验证。

| 模块         | Feedback / 训练反馈 | Selection / 验证 | Final / 测试 | 用途 |
|--------------|--------------------:|-----------------:|-------------:|------|
| PPT          |                  30 |               10 |           20 | 域内训练反馈、Champion 选择与最终测试 |
| DOCX         |                  30 |               10 |           20 | 域内训练反馈、Champion 选择与最终测试 |
| XLSX         |                  30 |               10 |           20 | 域内训练反馈、Champion 选择与最终测试 |
| Cross-Office |                   0 |                0 |           60 | Compositional OOD / Cross-Office Transfer |
| 合计         |                  90 |               30 |          120 | Final 包含 60 道域内题和 60 道组合式 OOD 题 |

## 划分原则

- Feedback 返回逐题轨迹、Rubric 得分和失败证据，用于 Harness Evolution。
- Selection 只返回聚合指标，用于选择或回退 Candidate，不能用于逐题修补。
- Final 在 Harness 锁定后一次性运行，用于论文最终报告。
- PPT 按单轮试校准的 12/18/18/12 档位等比例分层；DOCX/XLSX 沿用按任务家族、来源、依赖范围和预计耗时得到的原分区。
- 难度以本 Release 的逐题 `difficulty_*` 字段为准；回填难度不改变已冻结的任务分区。
- 相同 `leakage_group_id`、相同 `provenance_id` 或共享任一输入 SHA256 的任务合并为同一泄漏组件，整个组件只能进入一个分区。
- 与 Cross-Office 共享来源或输入的单模态任务一并进入 Final，避免 OOD 测试材料从 Feedback 或 Selection 泄漏。

## 文件责任

- `benchmark.json`：实验入口和三个分区的完整任务编号。
- `dataset.lock`：240 道题的不可变内容哈希与 Git task-tree 标识。
- `leakage-groups.jsonl`：泄漏组件、组成任务与分区。
- `splits/*.jsonl`：逐题路径、来源、家族、难度、角色、哈希与分区。
- 已标注难度：240/240；校准状态和来源见对应任务元数据。

## 边界

本仓库公开的是分区清单，因此 `sealed` 表示实验协议中的访问边界，而不是 Git 权限隔离。正式论文运行仍需由评测 Runner 在 Candidate 锁定前禁止读取 Final 的题目内容、Golden、Rubric 和 Judge。
