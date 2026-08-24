我是财务总监（CFO），需要为季度管理层会议准备 FY2026 部门“预算与实际”分析。

`/app/data/input_files/department_budgets_fy2026.pdf` 是已批准的 2026 财年部门预算。`/app/data/input_files/HR_ANALYTICS__PUBLIC__EMPLOYEES.copy.sql` 是从企业 HR 分析仓库的 `HR_ANALYTICS.PUBLIC.EMPLOYEES` 表原样抽取的离线 PostgreSQL COPY 快照，首行包含字段声明，数据行以 Tab 分隔，并以 `\.` 结束。它替代原题中的动态 ClickHouse 连接；表内的记录、部门、薪资和其他字段都不得补造或替换。

请从 PDF 读取每个部门的 `Approved Budget`，并从 COPY 快照中按 `DEPARTMENT` 分组，计算 `SALARY` 之和作为 `Actual Expenditure`、员工数作为 `Headcount`、平均薪资作为 `Avg Salary`。两份输入使用相同的俄文部门名：Инженерия、Финансы、Кадры、Операции、НИОКР、Продажи、Поддержка；必须用这些原始名称匹配。

在 `/app/output/` 中交付两个文件：

- `budget_vs_actuals.pptx`：可编辑、可直接演示的 16:9 PowerPoint。第 1 页标题必须是 `FY2026 Budget vs Actuals`，副标题必须是 `Quarterly Leadership Review - March 2026`。第 2 页标题为 `Department Overview`，包含部门总数、总批准预算、总实际薪资支出和总偏差（预算减实际）。之后按俄文部门名的字母序每个部门一页；页标题使用原始俄文部门名，页内至少包含 Approved Budget、Actual Expenditure、Variance、Variance Pct、Under Budget/Over Budget 状态、Headcount 和 Avg Salary。`Variance = Approved Budget - Actual Expenditure`，`Variance Pct = Variance / Approved Budget * 100`，金额和百分比保留两位小数。最后一页标题为 `Summary`，列出偏差最大的 `Most Under Budget` 部门和偏差最小的 `Most Over Budget` 部门。页面必须使用原生文字、图形或表格，不得用整页截图代替可编辑内容。

- `budget_vs_actuals.xlsx`：只包含 `Department Details` 和 `Totals` 两张工作表。`Department Details` 的列顺序必须是 `Department`、`Approved_Budget`、`Actual_Expenditure`、`Variance`、`Variance_Pct`、`Headcount`、`Avg_Salary`，并按俄文部门名字母序排列。`Totals` 为两列结构，依次包含 `Total_Departments`、`Total_Budget`、`Total_Actuals`、`Total_Variance`、`Most_Under_Budget`、`Most_Over_Budget` 及对应值。

不得联网补数据，也不得在演示文稿中加入输入无法支持的原因、预测、招聘计划、裁员计划或节省目标。两个交付物中的部门数字、总计、偏差与排名必须一致，并且能从两份输入独立重算。
