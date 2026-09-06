# CoworkEvoBench 240 题三段式划分 v0.1

本目录冻结 240 道 Harbor Task 的分区归属，不复制题目文件。`tasks/` 是唯一题目源；本目录只保存 Release、任务哈希、泄漏组件和 split 清单。

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
- PPT 按现有单轮试校准的 12/18/18/12 档位等比例分层；正式 `difficulty_*` 字段仍保持 `unassigned`。
- DOCX 和 XLSX 尚无冻结经验难度，因此只按任务家族、来源、依赖范围、预计耗时分层，不伪造难度标签。
- 相同 `leakage_group_id`、相同 `provenance_id` 或共享任一输入 SHA256 的任务合并为同一泄漏组件，整个组件只能进入一个分区。
- 与 Cross-Office 共享来源或输入的单模态任务一并进入 Final，避免 OOD 测试材料从 Feedback 或 Selection 泄漏。

## 文件责任

- `benchmark.json`：实验入口和三个分区的完整任务编号。
- `dataset.lock`：240 道题的不可变内容哈希与 Git task-tree 标识。
- `leakage-groups.jsonl`：泄漏组件、组成任务与分区。
- `splits/*.jsonl`：逐题路径、来源、家族、角色、哈希与分区。

## 边界

本仓库公开的是分区清单，因此 `sealed` 表示实验协议中的访问边界，而不是 Git 权限隔离。正式论文运行仍需由评测 Runner 在 Candidate 锁定前禁止读取 Final 的题目内容、Golden、Rubric 和 Judge。
