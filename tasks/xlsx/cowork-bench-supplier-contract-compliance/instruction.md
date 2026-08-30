# 供应商合同合规审计

请审计输入目录中的五份供应商合同。使用 `supplier_list.csv` 取得供应商名称、联系人、邮箱、服务类型、合同开始日期和年度合同金额；使用 `contract_template.pdf` 作为标准，并将每份 `contract_supplier*.pdf` 与模板逐项比较。合同文件中的俄文内容应按原意理解，交付文件的表头、列名和状态标识保留英文，说明性文字使用俄文。

请在 `/app/output` 创建以下六个文件：

- `compliance_assessment.xlsx`：工作表 `Compliance Assessment`，列为 `Supplier`、`Scope Clarity`、`SLA Present`、`Payment Terms`、`IP Rights`、`Confidentiality`、`Liability Limits`、`Compliance Score`。每个供应商一行，状态使用 `Complete`、`Incomplete`、`Compliant`、`Non-Compliant` 或 `Missing`，分数为 0–100。
- `corrective_actions.xlsx`：工作表 `Corrective Actions`，列为 `Supplier`、`Issue`、`Required Action`、`Owner`、`Target Date`、`Est. Cost`、`Status`。为每个存在差距的供应商至少列一项整改，`Target Date` 使用未来日期的 ISO 格式 `YYYY-MM-DD`。
- `audit_findings.docx`：逐供应商列出审计事实、具体差距（或未发现差距）、合规分数和建议。
- `audit_summary.docx`：列出平均合规分数以及 `High`、`Medium`、`Low` 风险供应商数量，并与 Excel 评估一致。
- `audit_notifications.eml`：给 `supplier_list.csv` 中的供应商联系人准备一封合同合规审计通知草稿，包含主题、收件人和需要整改的事实；不要实际发送邮件。
- `supplier_review_meetings.ics`：为高风险供应商准备离线日历草稿，至少包含供应商名称、审计改进会议标题、责任人和整改事项；不要声称已创建 Google Calendar 事件。

只使用输入文件能够支持的事实。两个 Excel 文件和两个 Word 文件应保持原生可编辑格式；不要用截图或无法打开的替代文件。
