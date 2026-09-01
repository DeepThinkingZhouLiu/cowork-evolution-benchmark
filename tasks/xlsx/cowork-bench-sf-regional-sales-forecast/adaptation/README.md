这道题逐题改造自 Cowork-Bench `sf-regional-sales-forecast`，不是本项目自编的预测业务故事。

- 上游原始 Query、方法论、假设、经济指标 JSON、Evaluator、Excel Golden 和 Word Golden 均已保留。Harbor v0.2.0 的唯一 solver 交付为 `Sales_Forecast.xlsx`；原 Word 报告的摘要、方法、区域分析、假设和建议改由 `Forecast Notes` 工作表承载。
- 原题要求从 ClickHouse 查 Q4 2025 的地区月度销售。本改造直接从同一上游数据库快照聚合 2,000 个客户和 20,000 个订单，只发布 15 个“地区×月”结果。
- 15 条历史记录的 Orders 和 Revenue 均与上游 Excel Golden 逐格一致，没有编造或修改销售数值。
- 上游 Google Sheets 只是 `Forecast` 表的在线副本；本离线版不产生外部写入，也不要求 Word 交付。
- 这是上游 mock 数据库题，不是真实公司生产数据，因此 provenance 明确标记 `synthetic=true`。
- Golden 与从四个静态输入重新计算的 replay 均得分 `1.0`；空/损坏工作簿、历史收入错误、区域预测错误、Summary 错误、缺少 `Forecast Notes` 或其中地区数字与预测表冲突均被拒绝。
- 工作簿可由 LibreOffice 打开和重算；`Forecast Notes` 的五个俄文章节与五个区域数字均在同一 XLSX 中可读取。上游未配置列宽的遗留打印限制不构成隐藏美化门槛。

发布前仍需用非参考 Agent 进行难度和异文容忍度校准。
