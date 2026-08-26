这道题逐题改造自 Cowork-Bench `sf-regional-sales-forecast`，不是本项目自编的预测业务故事。

- 上游原始 Query、方法论、假设、经济指标 JSON、Evaluator、Excel Golden 和 Word Golden 均已保留。
- 原题要求从 ClickHouse 查 Q4 2025 的地区月度销售。本改造直接从同一上游数据库快照聚合 2,000 个客户和 20,000 个订单，只发布 15 个“地区×月”结果。
- 15 条历史记录的 Orders 和 Revenue 均与上游 Excel Golden 逐格一致，没有编造或修改销售数值。
- 上游 Google Sheets 只是 `Forecast` 表的在线副本；本离线版不产生外部写入，保留 Excel 和 Word 两个实体交付物。
- 这是上游 mock 数据库题，不是真实公司生产数据，因此 provenance 明确标记 `synthetic=true`。
- Golden 与从四个静态输入重新计算的 replay 均得分 `1.0`；缺 Word、历史收入错误、区域预测错误、Summary 错误、报告过短、报告地区数字冲突和空文件均被拒绝。
- Word Golden 实际渲染为 2 页，章节、数字与分页正常。Excel Golden 可渲染为 3 页，但上游未配置列宽，打印视图会截断俄文地区名和长表头；原 Query 没有规定视觉样式，因此不把未要求的美化写成隐藏门槛。

发布前仍需用非参考 Agent 进行难度和异文容忍度校准。
