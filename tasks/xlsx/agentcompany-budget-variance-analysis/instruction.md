# 识别重大预算差异

请比较输入目录中的 `budget.xlsx` 和 `actual_spending.xlsx`。按 `Department`、`Category`、`Month` 精确联结两份数据，并对每个组合计算：

- `Variance_Amount = Actual_Amount - Budgeted_Amount`
- `Variance_Percentage = Variance_Amount / Budgeted_Amount * 100`

仅当差异金额绝对值严格大于 `$5,000`，且差异百分比绝对值严格大于 `10%` 时，标记该组合。两份输入的月度预算已经体现季节变化，因此必须按同一个月的预算与实际比较，不要改用年度平均值。

创建 `/app/output/flagged_variances.xlsx`，工作表命名为 `Flagged_Variances`。每条重大差异按输入顺序一行，至少包含以下列：`Flagged_Variances`、`Department`、`Category`、`Month`、`Budgeted_Amount`、`Actual_Amount`、`Variance_Amount`、`Variance_Percentage`。标识符格式为 `Department_Category_YYYYMM`。

保留正负号，金额保留两位小数，百分比以百分数形式清晰显示；交付原生可编辑的 Excel 工作簿。
