# Cross-Office 逐题扩展与验收台账 v0.1

## 建设目标

Cross-Office 最终目标为 60 道。当前仓库原有 10 道，本轮已逐题新增 3 道，现为 13/60；剩余 47 道继续按本文件的逐题流程建设。这里不使用通用脚本批量制造输入、Query、Golden 或统一 Judge。

## 每题必须完成的流程

1. 逐文件阅读上游 Query、全部 Agent 输入、预处理、Evaluator、Golden、数据卡和许可证；把实际读过的文件记录到 `adaptation/README.md`。
2. 冻结可公开输入并记录 SHA256。不得新增无法由上游证据支持的订单、金额、人物、结论或引用。
3. 明确跨文件因果链。至少两个应用或交付物必须共享关键事实，且上游错误需要传播到下游，不能只是把互不相关的文件放进同一目录。
4. 人工建立逐项真值台账，再设计本题专属 Golden。参考解法可以用代码落实已确认的变换，但不能用通用生成器替代逐题理解与判断。
5. 按 Harbor 目录协议提供 `task.toml`、离线 Docker 环境、`/app/input_files`、单一 `/app/output`、标准 Reward 和隐藏侧解法与评测。
6. Rubric 必须覆盖来源事实、原生 Office 结构、跨文件传播、可用性和安全边界；关键缺陷使用 hurdle cap，冲突、损坏、静态化和无来源断言使用 penalty。
7. Judge 必须从输入独立重算核心真值，并检查实际内容与结构；不得只比 Golden 哈希、文件名、关键词或页数。
8. Golden、空目录 fresh replay、空交付和定向负例全部运行；Office 文件逐页渲染目视检查。
9. 使用官方 Harbor Docker Oracle 至少独立回放两次，确认 Reward、异常数、真实 `/app/output` 和重复产物稳定。
10. 在冻结多 Harness Rollout 前，`difficulty_level` 和 `difficulty_tier` 保持 `unassigned`；L4 / Expert 等表述只能作为结构设计目标。

## 本轮进度

| 序号 | 任务目录                                                   | 上游来源                                      | 跨文件链路                         | 负例 | Harbor 回放       | 状态              |
|-----:|------------------------------------------------------------|-----------------------------------------------|------------------------------------|-----:|-------------------|-------------------|
|    1 | `cowork-bench-financial-audit-reconciliation-package`      | Cowork-Bench `financial-audit-cross-system-reconciliation` | CSV/XLSX/PDF -> XLSX -> DOCX -> EML/ICS |   11 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|    2 | `cowork-bench-conference-program-committee-package`        | Cowork-Bench `academic-conf-volga-rus-kazan-forms-teamly-excel-word-email` | Forms/Teamly -> XLSX -> DOCX/MD -> MBOX |   12 | 2 次，均 1.0/0 异常 | 已完成完整验收    |
|    3 | `cowork-bench-moex-portfolio-analysis-package`             | Cowork-Bench `moex-portfolio-analysis-excel-word-email` | XLSX/CSV/PDF -> XLSX -> DOCX -> EML |   14 | 2 次，均 1.0/0 异常 | 已完成完整验收    |

第一题的两次成功 trial 为 `cross-office-financial-audit-oracle-rerun-20260901` 和 `cross-office-financial-audit-oracle-repeat-20260901`；第二题为 `cross-office-conference-program-oracle-20260901` 和 `cross-office-conference-program-repeat-20260901`；第三题为 `cross-office-moex-portfolio-oracle-20260901` 和 `cross-office-moex-portfolio-repeat-20260901`。完整来源、真值、反例与回放记录位于各题的 `adaptation/` 与 `metadata/validation.json`。
