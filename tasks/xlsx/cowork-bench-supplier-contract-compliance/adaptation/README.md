# C096 逐题改造记录

状态：`pending_human_review`。这是 pilot 的 L4 作者假设，不是最终难度或发布决定。

## 来源和许可

- 来源：Cowork-Bench，commit `d943e75bc0fc8e3b27141979300cd8cbcd1e890d`，`tasks/finalpool/supplier-contract-compliance-audit`。
- 原始任务说明、评估脚本、配置和 Apache-2.0 LICENSE 已分别保留为 `upstream-task.md`、`upstream-evaluator.py`、`upstream-task-config.json`、`LICENSE-Cowork-Bench`。
- 输入为供应商 CSV、合同模板 PDF 和五份供应商合同 PDF，逐文件 SHA256 记录在 `metadata/input-manifest.json` 和 `metadata/provenance.json`。

## 独立事实和改造

以模板六个章节为基准，从 PDF 文本重算：Scope Clarity、SLA Present、Payment Terms、IP Rights、Confidentiality、Liability Limits。完整/合规维度各计一项，Compliance Score 为六项中通过项占比乘 100：TechService 83.33、GlobalLogistics 66.67、ControlQuality 83.33、ConsultGroup 66.67、ServicePodderzhka 100。按差距数划分风险：High=GlobalLogistics/ConsultGroup，Medium=TechService/ControlQuality，Low=ServicePodderzhka。上述结论未从上游 Golden 倒推。

原题要求两个 Excel、两个 Word，并要求实际发邮件和创建 Google Calendar。固定离线包保留四个主要交付物，另以 `audit_notifications.eml`（收件人和正文草稿）及 `supplier_review_meetings.ics`（高风险供应商会议草稿）表达外部副作用；明确不声称已发送或已创建。整改日期和负责人是作者侧计划值，不是输入事实，且仅要求未来 ISO 日期以便评测可验证。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 五供应商六维状态和分数 | CSV、模板及五份合同 PDF | R002 |
| 每个差距的整改动作 | 六维对照结果；作者侧行动模板 | R003 |
| 五供应商发现与建议 | 独立重算状态 | R004 |
| 平均分和风险分布 | 五行分数重新聚合 | R005/R008 |
| 联系人通知草稿 | CSV 的 Contact_Email | R006 |
| 高风险审查会议草稿 | 重算风险等级 | R007 |

## 验证记录

已执行：`validate-task`、`validate-task-assets` 通过；Golden 与 fresh replay 的 Judge 均为 `1.000000`。空输出/缺交付/损坏 DOCX 为 `0.000000`；错分、摘要跨文件不一致和遗漏整改分别为 `0.900000`、`0.800000`、`0.900000`，但对应核心门槛失败。LibreOffice 7.3.7.2 已打开并导出两个 Excel、两个 Word，四份 PDF 均可提取核心供应商和分数；ICS 含 2 个高风险 VEVENT，EML 收件人和主题结构已检查。Judge 只从 CSV/PDF 重算，不读取 Golden。

## 风险和人工决策

- 上游评估器依赖外部 MCP/ClickHouse 并把邮件发送作为关键检查，固定环境无法复现；本题 Judge 只接受可审计的本地等价物，需人工确认离线化是否足够保留原意。
- 上游 Golden 的分数与六维等权独立重算可能不同；本包以输入合同与模板为真值来源，并记录差异。
- PDF 文本提取依赖 `pdftotext` 和字体环境；Docker 固定安装 poppler 与 Liberation fonts。
- `difficulty_tier=expert` 仅为未校准 L4 结构假设；rollout 和人工复核仍未完成。
