# 30 天现金流预测（离线 XLSX）

你是财务分析师。请打开 `cash_flow_inputs.xlsx`，在一个可编辑的工作簿中完成未来 30 天的单期现金流预测，并保存为 `/app/output/cash_flow_forecast.xlsx`。

请保留输入工作表 `Bank Balance`、`AR Due` 和 `AP Due` 的内容，在同一工作簿新增：

- `Cash Flow Forecast`：仅处理 `Status` 为 `Include` 的 AR/AP 行。AR 的 Expected Inflow = Amount × Collection Probability / 100；AP 的 Outflow 为完整 Amount。列出开始余额、预计流入总额、AP 流出总额和期末现金余额。
- `Email Draft`：用表格模拟发给 `cfo@company.example.com` 的邮件，列出每个已纳入 AR 客户的预计流入、每个已纳入 AP 供应商的金额，以及期末余额。被标记 `Do Not Include` 的行不得出现在草稿中。

金额应为可编辑数值并使用千位分隔符；保留清晰的表头、筛选和冻结窗格。此题为离线改造，不需要联网，也不要实际发送邮件。
