请基于固定的 MOEX 市场快照和上游内部研究门户，为五只俄罗斯股票建立行业轮动 Excel 看板，保存为：

- `/app/output/sector_rotation_report.xlsx`

这是 Cowork-Bench `moex-sector-rotation-dashboard` 的离线静态改造版。`moex_snapshot.json` 保留上游数据库中 2026-02-25 和 2026-05-26 的 10 个收盘价，以及五只股票最新年度利润表的营收与净利润。`research_portal.html` 是上游原始研究门户页面，包含行业基准、分析师评级和目标价。

五只股票的固定顺序为 `SBER.ME, GAZP.ME, LKOH.ME, MGNT.ME, MTSS.ME`。计算：

- `Return_Pct = (Current_Price - Price_1Y_Ago) / Price_1Y_Ago * 100`
- `Alpha = Return_Pct - Benchmark_Return_Pct`
- `Upside_Pct = (Target_Price - Current_Price) / Current_Price * 100`
- 营收和净利润均除以 1,000,000 后保留两位；`Profit_Margin_Pct = Net_Income_Latest_Q / Revenue_Latest_Q * 100`。

工作簿必须含：

- `Performance`：`Symbol, Sector, Price_1Y_Ago, Current_Price, Return_Pct, Benchmark_Return_Pct, Alpha, Analyst_Rating, Target_Price, Upside_Pct`。
- `Financials`：`Symbol, Revenue_Latest_Q, Net_Income_Latest_Q, Profit_Margin_Pct`。上游列名保留 `Latest_Q`，但数值按原 Query 取最新年度 `income_stmt`。
- `Summary`：A:B 键值对，依次为 `Best_Performer, Worst_Performer, Avg_Alpha, Stocks_Above_Benchmark, Stocks_Below_Benchmark`。

百分比保留两位。所有正文必须保持为可编辑单元格，不得以截图代替。
