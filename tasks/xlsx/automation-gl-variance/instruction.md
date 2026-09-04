# 月末 GL 差异分析

请使用输入目录中的 `gl_variance_source.xlsx` 完成 2026 年 1 月月末总账差异分析，并将唯一交付保存为 `/app/output/GL_Variance_Analysis.xlsx`。

输入工作簿包含两张表：

- `January Actuals`：每个账户的账户类型、2026 年 1 月实际数、2025 年 12 月比较数和预算数。
- `Materiality Policy`：控制器提供的重大性阈值。百分比和金额阈值均为严格“大于”；Revenue 和 Expense 都要分别与上月、预算比较，Balance sheet 只与上月比较。

对每个允许的比较，计算 `Variance_Amount = Jan 2026 - Comparison Amount`，以及 `Variance_Percentage = Variance_Amount / Comparison Amount × 100`。只在满足该账户类型阈值的比较中报告重大差异。金额、方向和百分比必须保留，不要把非重大账户加入报告；对每条重大差异写出简短的阈值解释。

输出工作簿必须包含以下两张可编辑工作表：

1. `Material Variances`：列为 `Account`、`Type`、`January_Actual`、`Comparison_Period`、`Comparison_Amount`、`Variance_Amount`、`Variance_Percentage`、`Materiality_Basis`、`Explanation`。按输入账户顺序、再按比较顺序排列，只包含重大差异。
2. `Controller Email Draft`：作为原题邮件动作的离线草稿，列为 `Field`、`Value`，并包含 `From`、`To`、`Subject`、`Body`、`Status` 五行。收件人必须为 `controller@company.example.com`，主题应说明这是 GL variance analysis，正文逐条列出 `Material Variances` 中的账户、比较期间、金额差异和百分比，并标记为 `Draft - Not Sent`。不要实际发送邮件，也不要声称发送成功。

金额使用美元格式并保留两位小数，百分比清晰显示为百分数；工作表应冻结首行、可筛选、列宽合理，并保持原生可编辑，不能用截图代替表格。
