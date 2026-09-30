# CoworkEvoBench 240 题三段式划分 v0.5

本目录冻结 240 道 Harbor Task 的分区归属，不复制题目文件。`tasks/` 是唯一题目源；本目录只保存 Release、任务哈希、泄漏组件和 split 清单。

## 当前公开锁定

公开分支使用任务快照 `5e0690bcdbebabb829e776b1aba7c32d079d27c2`，任务树为 `f502e04190fdee10a0ecef25678c0fbfebdc8240`。这是对历史冻结提交 `7eb448bd6fa25cf973c35b11868b9afbb7f55e89` 做的可移植路径重锁：只替换机器相关路径解析，不改变任务语义、输入、Golden、Rubric 或 Judge。原始冻结提交保留在 Git 历史，完整实验基线保留在归档工作树中；今后任何任务、Judge、Rubric 或 split 变化都必须创建新的 Release。

| 模块         | Feedback / 诊断反馈 | Selection / 验证 | Final / 测试 | 用途 |
|--------------|--------------------:|-----------------:|-------------:|------|
| PPT          |                  30 |               10 |           20 | 域内诊断、聚合验证与最终测试 |
| DOCX         |                  30 |               10 |           20 | 域内诊断、聚合验证与最终测试 |
| XLSX         |                  30 |               10 |           20 | 域内诊断、聚合验证与最终测试 |
| Cross-Office |                   0 |                0 |           60 | Compositional OOD / Cross-Office Transfer |
| 合计         |                  90 |               30 |          120 | Final 包含 60 道域内题和 60 道组合式 OOD 题 |

## 划分原则

- Feedback 返回逐题轨迹、Rubric 得分和失败证据，用于误差分析、评测调试和训练记录。
- Selection 只返回聚合指标，用于模型与 Harness 的选择性验证，不能用于逐题修补。
- Final 在评测配置冻结后一次性运行，用于论文最终报告。
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

本仓库公开的是分区清单，因此 `sealed` 表示实验协议中的访问边界，而不是 Git 权限隔离。正式运行仍需由评测 Runner 在运行前禁止受测 Agent 读取 Final 的题目内容、Golden、Rubric 和 Judge。
