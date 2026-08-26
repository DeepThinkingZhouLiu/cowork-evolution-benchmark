请阅读 `reorder_policy.pdf`，使用固定商店快照完成库存补货预测，并将最终工作簿保存为：

- `/app/output/Inventory_Forecast.xlsx`

这是 Cowork-Bench `insales-inventory-reorder-forecast` 的离线静态改造版。上游原题需要调用 InSales 和 Google Calendar；本题把同一上游 mock 数据库快照固化为文件，不执行任何外部写操作。

输入文件的责任如下：

- `store_products.json`：上游 `wc.products` 的 82 个产品快照；仅筛选 `status=publish` 且 `current_stock>0` 的产品。
- `store_orders.json`：上游 `wc.orders` 的 150 个订单快照；销量仅统计 `completed` 和 `processing` 订单的 `line_items`。
- `supplier_lead_times.json`：供应商品类交付期，按产品第一个品类名精确匹配。
- `snapshot_date.txt`：固定计算日期 `2026-03-07`，用它代替动态的“今天”。
- `sales_velocity.xlsx`：上游给定的销售速度表头模板，可作为中间计算参考，不是最终交付物。

按原政策执行以下计算：

- 订单历史天数等于符合状态的最晚订单日期减最早订单日期，不额外加 1。
- `Avg_Daily_Sales = Total_Sold / Days_Period`，保留 4 位小数。
- 有效库存为 `max(Current_Stock - 5, 0)`。
- 无销售历史时，`Days_Until_Safety_Stock` 记为 `999`，且不补货；否则用有效库存除以平均日销量，保留 1 位小数。
- 有销售历史且距安全库存不超过 90 天时，`Needs_Reorder=Yes`。
- 补货日期为固定计算日期 + `floor(Days_Until_Safety_Stock)` - 交付期；若结果不晚于计算日期，改为次日。

工作簿必须包含两张可编辑工作表：

- `Stock Analysis`：`Product_ID, Product_Name, Category, Current_Stock, Avg_Daily_Sales, Days_Until_Safety_Stock, Lead_Time, Needs_Reorder, Reorder_Date`；包含全部符合库存条件的产品。
- `Reorder Schedule`：`Product_Name, Category, Current_Stock, Reorder_Date, Lead_Time_Days`；只包含需要补货的产品。

不需要在真实日历中创建事件；`Reorder Schedule` 就是本离线版的可回放日程交付物。
