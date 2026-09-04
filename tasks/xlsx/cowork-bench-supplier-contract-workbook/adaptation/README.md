# R057 / Cowork-Bench supplier-contract-compliance-audit

状态：`pending_human_review`；`difficulty_tier=expert` 是作者的 L4 假设，不是校准或发布决定。

## 来源和许可

- 来源仓库：Cowork-Bench，Apache-2.0，commit `d943e75bc0fc8e3b27141979300cd8cbcd1e890d`，原路径 `tasks/finalpool/supplier-contract-compliance-audit`。
- 原始任务、原始 evaluator、task config 和 LICENSE 保留在本目录。输入来自原任务的 `initial_workspace`，并逐文件复制到 `data/input_files/` 和 `environment/input_files/`。
- 七个 solver-visible 输入的 SHA256 记录在 `metadata/input-manifest.json` 和 `metadata/provenance.json`；两份目录副本已做字节一致性检查。

## 为什么替换 R057

矩阵原 R057 候选 `budget-variance-analysis-quarterly` 的 `initial_workspace/data.csv` 只有两行 `Item1/Item2` 占位数据，且原任务允许在缺失实际额时自行估算。原 evaluator 也明确不比较 Actual YTD、Variance、月份或预测的绝对真值，因此它不能支持可重算的预算差异任务，已在作者审计中标为 `rejected_no_truth`，不应进入 rollout。

本题改用同一 revision 中事实完整的 `supplier-contract-compliance-audit`。五份供应商合同和合同模板均可离线读取，六个合规维度、分数和风险可从输入独立重算，适合多阶段 XLSX 审计工作流。

## XLSX-only 压缩

原任务要求两个 XLSX、两个 DOCX、邮件和日历。根据 Excel 任务输出契约，保留一个自然的 `supplier_contract_audit.xlsx`，将每个有意义的交付环节映射为工作表：

| 原始环节 | XLSX 工作表 | 处理 |
|---|---|---|
| 合同基础信息 | `Supplier Register` | 直接复制 CSV 字段并加入对应合同文件名 |
| 合规评分卡 | `Compliance Assessment` | 从模板和五份合同 PDF 独立重算六维状态、分数、风险 |
| 整改计划 | `Corrective Actions` | 每个真实缺口一行，保留未来 ISO 日期和责任人字段 |
| 供应商发现报告 | `Audit Findings` | 一供应商一行，包含缺口、证据和建议 |
| 管理层摘要 | `Audit Summary` | 平均分、风险计数、高风险名单使用工作表公式 |
| 邮件草稿 | `Notification Drafts` | 收件人来自 CSV，明确 `Draft - Not Sent` |
| 会议草稿 | `Review Meetings` | 每个 High 风险供应商一行，明确 `Draft - Not Created` |

没有把 DOCX/PDF/EML/ICS 作为伪输出，也没有加入网络合同事实。公式、表格、冻结窗格和自动重算设置保持原生可编辑。

## 独立真值

Judge 只读取 `supplier_list.csv`、`contract_template.pdf` 和五份合同 PDF。对每份合同使用模板六个章节的规则：Scope/SLA 需为 `Complete`，其余四项需为 `Compliant`；分数为六项通过数除以 6；零缺口为 Low、一个缺口为 Medium、至少两个缺口为 High。独立重算得到：TechService 83.33/Medium、GlobalLogistics 66.67/High、ControlQuality 83.33/Medium、ConsultGroup 66.67/High、ServicePodderzhka 100/Low。

## Judge 和验证

Judge 不读取 Golden、不比较工作簿 hash；会检查输入哈希、CSV 字段、六维合同状态、分数/风险、整改覆盖、发现和摘要跨表一致性、通知/会议草稿、公式和自动重算，并调用 LibreOffice 导出 PDF（22 页）确认工作簿可渲染。

已执行：

```bash
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/xlsx/cowork-bench-supplier-contract-workbook
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/xlsx/cowork-bench-supplier-contract-workbook
```

Golden 和 fresh replay 的 Judge reward 均为 `1.000000`。负例：空输出和损坏 XLSX 为 `0.000000`；错误合规状态为 `0.450000`；错误摘要公式为 `0.700000`；遗漏整改为 `0.600000`。这些结果均以 hurdle 失败退出，且保留了局部错误的分数梯度。当前尚未进行无辅助模型 rollout 或人工接受。

## 难度假设和风险

作者暂定 L4：7 个异构输入文件、俄文合同文本抽取、模板对照、六维评分、整改、摘要、通知与会议草稿，以及跨表公式一致性。难度仍为 `uncalibrated`，需用冻结 Harness 和人工复核验证；合同文本来自同一 Cowork-Bench fixture，leakage group 为 `cowork-contract-compliance`，不得与同源任务跨切分。
