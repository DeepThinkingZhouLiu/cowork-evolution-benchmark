这道题逐题改造自 Cowork-Bench `insales-inventory-reorder-forecast`，不是本项目自编的库存业务故事。

- 上游原题的题干、PDF 政策、Excel 模板、交付期 JSON、确定性 Evaluator、事件真值和 Golden 工作簿均已保留。
- 原题通过 InSales MCP 查询本地 mock 商店。本改造直接从同一上游 PostgreSQL 快照的 `wc.products` 和 `wc.orders` 投影必要字段，得到 82 个产品和 150 个订单；没有编造或补齐记录。
- 为避免带出与任务无关的 mock 客户信息，未投影账单、地址、电话、邮箱和 IP 字段。
- 上游的“今天”对应 `2026-03-07`；为使日期可回放，改为明示的 `snapshot_date.txt`。
- 上游要求写 Google Calendar。本离线版不产生外部副作用，以同一工作簿中的 `Reorder Schedule` 作为可重放交付物。
- 从静态快照复算可得：符合状态的 100 个订单覆盖 `2025-03-08` 至 `2026-03-04`，相差 361 天；77 个正库存产品中 18 个需补货。
- 复算后的 `Stock Analysis` 77 行和 `Reorder Schedule` 18 行与上游 Golden 逐单元格一致。Golden 保持上游原文件，SHA256 为 `9cd90d6ff40663c31dc4073bc292d06c690d9d41f56444ce53d24c68772a777f`。
- Golden 和从六个静态输入重新计算的 replay 均得分 `1.0`；空模板、错误日均销量、漏标补货、错误补货日期、伪造产品 ID 和空文件均被拒绝。
- 上游 Golden 没有设置列宽、冻结或打印样式；LibreOffice 可正常打开和重存，但 PDF 打印视图会截断长产品名。原 Query 只要求数据工作簿，所以本题不把上游未要求的美化当作硬门槛；发布前仍需用非参考 Agent 校准可用性。

这份商店数据是上游公开 Benchmark 的 mock fixture，不是真实公司生产数据；在 provenance 中因此标记 `synthetic=true`。
