# 供应商合同合规审计工作簿

请审计输入目录中的五份供应商合同。使用 `supplier_list.csv` 取得供应商名称、联系人、邮箱、服务类型、合同开始日期和年度合同金额；使用 `contract_template.pdf` 作为标准，并将每份 `contract_supplier*.pdf` 与模板的六个指定章节逐项比较。合同中的俄文按原意理解；交付工作簿的工作表名、表头和状态标识保留英文，说明性文字可使用中文或俄文。所有事实必须来自输入文件。

只创建一个文件：`/app/output/supplier_contract_audit.xlsx`。工作簿必须是原生、可编辑的 XLSX，不能提交 DOCX、PDF、图片、CSV、脚本或其他文件。

工作簿应包含以下工作表，并按此顺序排列：

1. `Supplier Register`：每个供应商一行，列为 `Supplier`、`Primary Contact`、`Contact Email`、`Service Type`、`Contract Start Date`、`Annual Contract Value`、`Contract File`。这些字段直接取自 `supplier_list.csv`，合同文件名按供应商对应的输入文件填写。
2. `Compliance Assessment`：每个供应商一行，列为 `Supplier`、`Scope Clarity`、`SLA Present`、`Payment Terms`、`IP Rights`、`Confidentiality`、`Liability Limits`、`Compliance Score`、`Risk`。将合同与模板比较：Scope Clarity 和 SLA Present 使用 `Complete`/`Incomplete`；Payment Terms、IP Rights、Confidentiality、Liability Limits 使用 `Compliant`/`Non-Compliant`/`Missing`。六个维度中满足模板要求的数量除以 6，再乘 100，得到 `Compliance Score`（保留两位小数）。有至少一个缺口但少于两个缺口为 `Medium`，有两个或更多缺口为 `High`，没有缺口为 `Low`。
3. `Corrective Actions`：对每个存在缺口的供应商至少列一项整改，列为 `Supplier`、`Issue`、`Required Action`、`Owner`、`Target Date`、`Est. Cost`、`Status`。Issue 必须指向实际缺口；Target Date 使用未来的 ISO 日期 `YYYY-MM-DD`；Status 可使用 `Planned` 或 `In Progress`。没有缺口的供应商可以不出现在此表。
4. `Audit Findings`：每个供应商一行，列为 `Supplier`、`Compliance Score`、`Risk`、`Findings`、`Evidence`、`Recommendation`。Findings 必须说明具体缺口，或明确写出未发现六个维度的缺口；Evidence 至少引用对应合同章节或说明缺失；Recommendation 必须与缺口一致。
5. `Audit Summary`：建立管理层摘要，列为 `Metric`、`Value`，至少包括 `Average Compliance Score`、`High Risk Suppliers`、`Medium Risk Suppliers`、`Low Risk Suppliers`，并可加入各风险级别的供应商名单。摘要数值必须与 `Compliance Assessment` 一致，并使用可重算的工作表公式或等价的可编辑计算。
6. `Notification Drafts`：为每个供应商联系人准备一行离线通知草稿，列为 `Supplier`、`To`、`Subject`、`Draft Body`、`Status`。To 必须取 `supplier_list.csv` 的 `Contact_Email`；正文包含该供应商分数、风险和整改事项（若有）；Status 明确标记为 `Draft - Not Sent`。不要实际发送邮件。
7. `Review Meetings`：为每个 `High` 风险供应商准备一行离线会议草稿，列为 `Supplier`、`Meeting Title`、`Owner`、`Corrective Actions`、`Target Date`、`Status`。标题必须明确为合同合规改进审查，Owner 和整改事项不能为空；Status 标记为 `Draft - Not Created`。不要声称已经创建日历事件。

保留输入事实的来源关系，不要根据网络搜索补充合同条款、供应商信息或价格。完成后确认 `/app/output` 中只有 `supplier_contract_audit.xlsx`。
