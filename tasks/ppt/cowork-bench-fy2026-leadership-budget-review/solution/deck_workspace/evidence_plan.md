# 证据计划

- Approved Budget：只来自 `department_budgets_fy2026.pdf`。

- Actual Expenditure：只来自 `HR_ANALYTICS__PUBLIC__EMPLOYEES.copy.sql` 中同部门 `SALARY` 之和。

- Headcount：只来自同部门员工记录数。

- Avg Salary：`Actual Expenditure / Headcount`，保留两位小数。

- Variance：`Approved Budget - Actual Expenditure`。

- Variance Pct：`Variance / Approved Budget * 100`，保留两位小数。

- Most Under Budget：Variance 最大的部门。

- Most Over Budget：Variance 最小的部门；即使所有偏差都为正，也使用这一排名口径。

`truth.json` 是本次 Golden 构建的可读计算记录；Judge 不读该文件，而是用独立实现从任务输入重算。
