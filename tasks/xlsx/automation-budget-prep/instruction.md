# 2026 年度部门预算

请使用输入目录中的 `budget_preparation_source.xlsx`，按 `Rate Card` 表的类别增长率编制 2026 年预算，并将唯一交付保存为 `/app/output/Annual_Budget_2026.xlsx`。

输入工作簿包含：

- `Rate Card`：`Salaries`、`Travel` 和 `Software` 三类增长率。
- `2025 Actuals`：各部门 2025 年三类实际金额及 `Scope`。只处理 `Scope` 为 `Include` 的部门；`Exclude` 部门不得出现在 2026 预算或邮件正文中。
- 空白的 `2026 Budget` 模板。

对每个纳入范围的部门，分别将 2025 金额乘以 `1 +` 对应类别增长率，得到 2026 Salaries、Travel 和 Software。金额按整美元显示并使用千位分隔符，不得使用统一增长率替代类别增长率。保留可审计的 Excel 公式，使源数据或增长率变化后结果可重算。

输出工作簿必须包含四张可编辑工作表：

1. `Rate Card`：保留输入增长率。
2. `2025 Actuals`：保留输入实际数和 Scope。
3. `2026 Budget`：列为 `Department`、`2026 Salaries`、`2026 Travel`、`2026 Software`，每个纳入部门一行，按输入顺序排列，并包含末行 `Total` 汇总。
4. `CFO Email Draft`：列为 `Field`、`Value`，作为原题邮件动作的离线草稿。必须包含 `To`（`cfo@company.example.com`）、说明 2026 预算汇总及各纳入部门金额的 `Body`，以及 `Status`=`Draft - Not Sent`。不得实际发送邮件或声称已发送。

所有表应冻结首行、可筛选、列宽合理；不得用截图代替可编辑工作簿。
