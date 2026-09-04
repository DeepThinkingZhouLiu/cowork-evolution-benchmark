请根据固定的 2025 年第四季度区域销售快照和 2026 年第一季度经济指标，建立区域销售预测模型，并将唯一交付保存为 `/app/output/Sales_Forecast.xlsx`。

这是 Cowork-Bench `sf-regional-sales-forecast` 的离线静态改造版。请先阅读 `forecast_methodology.md` 和 `assumptions.txt`。

数据文件的责任如下：

- `historical_sales.json`：上游 ClickHouse mock 快照的 15 条聚合结果，即 5 个区域×10月、11月、12月；字段为 `region, month, orders, revenue`。
- `economic_indicators.json`：上游经济指标 API 的原始 JSON；本模型使用各区域 `gdp_growth`，其他指标只作背景。

预测公式为：

- `Trailing_3M_Avg = (Oct_Revenue + Nov_Revenue + Dec_Revenue) / 3`
- `Economic_Adjustment = 1 + GDP_Growth_Pct / 100`
- `Adjusted_Monthly_Forecast = Trailing_3M_Avg * Economic_Adjustment`
- `Next_Quarter_Forecast = Adjusted_Monthly_Forecast * 3`

`Sales_Forecast.xlsx` 必须包含：

- `Historical`：`Region, Month, Orders, Revenue`，完整保留 15 条来源记录；`Orders` 为整数，`Revenue` 保留两位。
- `Forecast`：`Region, Trailing_3M_Avg, GDP_Growth_Pct, Economic_Adjustment, Adjusted_Monthly_Forecast, Next_Quarter_Forecast`，每个区域一行；移动平均与预测金额保留两位，调整系数保留四位。
- `Summary`：`Metric, Value`，包含 `Total_Q4_2025_Revenue, Total_Monthly_Average, Total_Q1_2026_Forecast, QoQ_Growth_Pct, Highest_Forecast_Region, Highest_Forecast_Amount, Lowest_Forecast_Region, Lowest_Forecast_Amount, Highest_GDP_Growth_Region, Regions_Count`。

工作簿还必须包含 `Forecast Notes` 工作表，用俄文写作，至少包含“简要摘要、方法论、区域分析、关键假设、建议”五个清晰章节，并引用五个区域的具体预测数字。

上游还要求把 `Forecast` 数据写入 Google Sheets。本离线版不产生外部副作用，不需要创建在线表格；Excel 的 `Forecast` 工作表是唯一预测数据交付源。
