请以工作区中的 `sales_api_response.json` 和 `inventory_api_response.json` 为唯一事实来源，完成一次可回放的数据清洗与电子表格交付。只保存一个文件：`/app/output/Data_Pipeline_Report.xlsx`。

这是 Cowork-Bench `fetch-terminal-data-pipeline` 的逐题离线改造版。原题、两份 API 响应、清洗规则和 Excel 分析要求保持不变；原 Google Sheet 的内容并入本工作簿的 `Pipeline Dashboard` 工作表，原 memory 写入并入 `Run Summary`。输入记录、Query 和业务数字都来自上游 Apache-2.0 任务，不是本项目新编。

清洗口径：

- 销售按 `order_id` 去重，并严格保留输入顺序中的第一次出现；库存按 `product_id` 去重，也保留第一次出现。
- 只把 `product_name` 做 `strip + title case`。其余字段和值必须保持第一次出现的原值，特别是不得把俄文 `warehouse` 翻译成英文。
- 清洗后的销售和库存记录须在工作簿中可追溯；不需要提交 JSON 中间文件。
- 销售分析按清洗后的 `product_name` 聚合：`Total_Quantity_Sold` 为数量之和，`Total_Revenue` 为逐订单 `quantity * unit_price` 之和，`Avg_Unit_Price` 为该商品清洗后订单的 `unit_price` 算术平均。
- 库存 `Status`：`Current_Stock <= Reorder_Point` 时为 `LOW`，否则为 `OK`。
- `Combined View` 包含至少出现在任一清洗后数据集中的所有商品；列为 `Product_Name, Total_Sold, Current_Stock, Revenue, Stock_Status`。某一侧缺失时，数值使用 `0`、状态使用 `NO INVENTORY`。

Judge 会从输入独立复算，并以变异快照复核去重逻辑：第一次出现的销售数量和库存量变化应传播到聚合和库存状态，后续重复记录即使取极端值也仍须忽略。不得把 `18`、`5` 或各商品汇总写死。

工作簿必须包含 `Run Summary` 工作表，准确记录原始/清洗后记录数、删除的重复数、总销量、总收入和 LOW 商品数。

`Data_Pipeline_Report.xlsx` 必须包含以下五张原生可编辑工作表：

- `Sales Analysis`：列为 `Product_Name, Total_Quantity_Sold, Total_Revenue, Avg_Unit_Price`。
- `Inventory Status`：列为 `Product_Name, Current_Stock, Reorder_Point, Status`；`Status` 必须是引用本行库存和补货点的 Excel 公式。
- `Combined View`：列为 `Product_Name, Total_Sold, Current_Stock, Revenue, Stock_Status`。
- `Pipeline Dashboard`：与 `Combined View` 完全相同的分析视图，用于离线替代原题 Google Sheet。
- `Run Summary`：原始/清洗后记录数、删除重复数、总销量、总收入和 LOW 商品数。

所有工作表须保留可编辑单元格、清晰表头、冻结首行、筛选/原生表格、合理列宽和货币格式。脚本、JSON 中间结果和进度记录留在参考解法与评测侧，不属于 solver 交付物。
