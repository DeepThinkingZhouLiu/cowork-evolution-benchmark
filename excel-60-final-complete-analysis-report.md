# Excel 60 题 Sonnet 5 Rollout 最终分析

**日期**：2026-09-01
**Harness/模型**：Claude Code orchestrated / `claude-sonnet-5`
**范围**：`tasks/xlsx` 活动集 60 题，每题 3 次，共 180 trials
**原始汇总**：`.coworkevo/runs/xlsx-60-final-complete-summary.json`
**题级清单**：[`datasets/xlsx-rollout-20260901/manifest.json`](datasets/xlsx-rollout-20260901/manifest.json) 及 [`task-results.csv`](datasets/xlsx-rollout-20260901/task-results.csv)

## 结论先行

这批结果可以作为 **Sonnet 5 的难度校准证据**，但不能直接作为任务发布验收。60 个题包和 180 次 trial 均有记录，整体平均 reward 为 **0.438649（43.86/100）**。L1 基础操作可靠，L2 已进入能力边界；L3/L4 的长链、多文件和专业建模任务对当前配置普遍不可解或不稳定。所有题包仍保持 `pending_human_review`。

## 核心指标

| 指标 | 数值 | 口径 |
|---|---:|---|
| 完成题数 | 60/60 | 每题都有 3 次已完成 trial |
| 总 trials | 180 | 60 × 3 |
| 平均 Score | 43.86/100 | 180 个 reward 的平均值 |
| 满分题 | 13 | 题级均值等于 1.0 |
| Reward ≥ 0.8 题 | 19 | 三次中至少一次达到 0.8（best-of-3） |
| 非零分题 | 47 | 题级均值大于 0 |
| 零分题 | 13 | 三次均为 0，题级均值等于 0 |
| 单次通过 trials | 56/180（31.1%） | 单次 reward ≥ 0.7 |

“19 题高分”是 best-of-3 指标；若要求三次均达到 0.8，则为 17 题。题级指标与 trial 级指标在清单中分列，避免把 55/180 误写成 55 道题。

## 难度分层

| 档位 | 题数 | trials | 平均分 | 通过 trials（≥0.7） | 通过率 | 零分题 |
|---|---:|---:|---:|---:|---:|---:|
| L1 basic | 15 | 45 | 84.67 | 36/45 | 80.0% | 1 |
| L2 intermediate | 16 | 48 | 52.40 | 19/48 | 39.6% | 2 |
| L3 advanced | 17 | 51 | 19.53 | 1/51 | 2.0% | 6 |
| L4 expert | 12 | 36 | 15.96 | 0/36 | 0.0% | 4 |

相邻档位平均分下降：L1→L2 **32.27** 分，L2→L3 **32.87** 分，L3→L4 **3.57** 分。主要问题不是 L3 与 L4 的区分，而是 L2 到 L3 出现明显断层；L3/L4 题目需要先检查可解性、指令链长度和交付物一致性，再决定是否保留当前难度假设。

## 题库构成

最终活动集由以下部分组成：

| 来源 | 题数 | 说明 |
|---|---:|---|
| 原始 60×3 中保留 | 42 | 原始 rollout 中已完成的活动题 |
| 原始环境失败题重测 | 12 | 修复环境后重新执行，纳入最终 60 题 |
| 新增 L1 题 | 6 | 六题均为 3/3、1.0 |
| 合计 | 60 | 42 + 12 + 6 |

另有 6 个结构重复候选移入 `datasets/xlsx-archive/`，不计入本次活动集：`officebench-2-13-section-calendar`、`officebench-2-15-calendar-to-xlsx`、`officebench-3-1-task-pdf-email`、`officebench-3-6-class-training-email`、`sb2-template-14-09-concession-revenue`、`sb2-visualization-129-mars-valuation`。归档清单见 [`datasets/xlsx-archive-manifest.csv`](datasets/xlsx-archive-manifest.csv)。

## 主要发现

1. **L1 设计有效**：15 题平均 84.67 分，新增 6 题全部满分；适合作为基础能力和回归集。
2. **L2 是当前边界**：平均 52.40 分，39.6% trial 通过；简单多阶段公式、格式保持和双工作簿任务之间已有明显差异。
3. **L3/L4 需要审计而不是直接加难**：40/60 题没有任何一次达到 0.7，长链交付、多文件一致性、图表/邮件/日历和财务模型是集中失败点。
4. **稳定性分层清楚**：17 题三次均通过，3 题出现部分通过，40 题三次均未通过。best-of-3 高分但均值不高的两个题是 `agentcompany-budget-variance-analysis` 和 `officebench-1-6-class-schedule`，应按波动性单独复核。
5. **成本记录需保留来源说明**：本次运行报告给出的总量为 **120,256,921 tokens / $113.73**；原始题级 summary 不含 token 字段，因此清单将其标为 run-level reported accounting，而非从题级 reward 反推。

## 后续处理建议

- L1：保留为基础回归集，人工确认重复度和来源许可后再进入切分。
- L2：优先检查阈值边界、输出格式与 Judge 颗粒度，保留通过率能区分模型的题目。
- L3/L4：逐题做输入闭环、专业规则、交付物链和失败产物审计；必要时拆分任务、降级或暂缓，不用统一脚本批量改写。
- 正式发布前：完成至少另一套无辅助 Harness/模型的交叉 rollout、人工复核、泄漏组切分和 `pending_human_review` 决策；本报告不替代这些步骤。

## 可复现检查

```bash
python scripts/build_xlsx_rollout_manifest.py
PYTHONPATH=src python -m cowork_evo_bench.cli validate-public-tasks tasks
```

manifest 脚本会检查 60 个题目录、180 个 trial，并从每题 `metadata/task.json` 读取任务分类、来源和 leakage group；生成的 JSON 记录原始汇总 SHA256，CSV 提供逐题可筛选结果。
