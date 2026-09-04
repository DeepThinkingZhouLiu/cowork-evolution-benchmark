# 供应商对账（离线 XLSX）

你是应付账款分析师。请打开 `vendor_reconciliation_inputs.xlsx`，把供应商对账结果整理到一个可编辑工作簿，并保存为 `/app/output/vendor_reconciliation.xlsx`。

请保留输入工作表 `Vendor Statements Received` 和 `Our AP Balances`，在同一工作簿新增：

- `Reconciliation`：按供应商声明顺序联结两张输入表，列出声明余额、AP 余额、差额（声明减 AP）、状态和是否需要联系。差额不为零的供应商标为 `Unresolved discrepancy` 并需要联系；相等的标为 `Matched`。AP 的 Notes 含 `exclude from recon` 的供应商必须标为 `Excluded` 且不产生对账差额或联系要求。
- `Email Draft`：用表格模拟邮件。给 `controller@company.example.com` 一封汇总邮件；只给有未解决差异的供应商发送草稿，正文包含声明余额、AP 余额和差额。不要为已匹配或明确排除的供应商创建邮件草稿。

金额应为可编辑数值并使用千位分隔符；保留清晰的表头、筛选和冻结窗格。此题为离线改造，不需要联网，也不要实际发送邮件。
