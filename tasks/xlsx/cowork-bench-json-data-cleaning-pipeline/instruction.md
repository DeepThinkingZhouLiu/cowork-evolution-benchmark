请以工作区中的 `sales_api_response.json`、`inventory_api_response.json` 和空白 `memory_seed.json` 为唯一事实来源，完成一次可回放的数据清洗与电子表格交付。请保存：

- `/app/output/clean_data.py`
- `/app/output/cleaned_sales.json`
- `/app/output/cleaned_inventory.json`
- `/app/output/pipeline_results.json`
- `/app/output/Data_Pipeline_Report.xlsx`
- `/app/output/Pipeline_Dashboard.xlsx`
- `/app/output/progress_note.json`

这是 Cowork-Bench `fetch-terminal-data-pipeline` 的逐题离线改造版。原题、两份 API 响应、清洗规则和 Excel 要求保持不变；原 Google Sheet 改为同内容的离线 `Pipeline_Dashboard.xlsx`，原 memory 写入改为 `progress_note.json`。输入记录、Query 和业务数字都来自上游 Apache-2.0 任务，不是本项目新编。

清洗口径：

- 销售按 `order_id` 去重，并严格保留输入顺序中的第一次出现；库存按 `product_id` 去重，也保留第一次出现。
- 只把 `product_name` 做 `strip + title case`。其余字段和值必须保持第一次出现的原值，特别是不得把俄文 `warehouse` 翻译成英文。
- `cleaned_sales.json` 顶层为 `sales` 和 `count`；`cleaned_inventory.json` 顶层为 `inventory` 和 `count`。
- 销售分析按清洗后的 `product_name` 聚合：`Total_Quantity_Sold` 为数量之和，`Total_Revenue` 为逐订单 `quantity * unit_price` 之和，`Avg_Unit_Price` 为该商品清洗后订单的 `unit_price` 算术平均。
- 库存 `Status`：`Current_Stock <= Reorder_Point` 时为 `LOW`，否则为 `OK`。
- `Combined View` 包含至少出现在任一清洗后数据集中的所有商品；列为 `Product_Name, Total_Sold, Current_Stock, Revenue, Stock_Status`。某一侧缺失时，数值使用 `0`、状态使用 `NO INVENTORY`。

`clean_data.py` 必须真实读取输入并支持：

```bash
python clean_data.py \
  --sales /app/data/input_files/sales_api_response.json \
  --inventory /app/data/input_files/inventory_api_response.json \
  --sales-output /app/output/cleaned_sales.json \
  --inventory-output /app/output/cleaned_inventory.json \
  --analysis-output /app/output/pipeline_results.json
```

Judge 会构造一份变异输入：修改第一次出现的销售数量和库存量，同时把两个后续重复记录改成极端值。正确脚本必须让首次记录的变化传播到聚合和库存状态，同时继续忽略后续重复项；不得把 `18`、`5` 或各商品汇总写死。

`pipeline_results.json` 必须包含 `summary`、`sales_analysis`、`inventory_status` 和 `combined_view`。其中 `summary` 至少准确记录原始/清洗后记录数、删除的重复数、总销量、总收入和 LOW 商品数。

`Data_Pipeline_Report.xlsx` 必须包含三张原生可编辑工作表：

- `Sales Analysis`：列为 `Product_Name, Total_Quantity_Sold, Total_Revenue, Avg_Unit_Price`。
- `Inventory Status`：列为 `Product_Name, Current_Stock, Reorder_Point, Status`；`Status` 必须是引用本行库存和补货点的 Excel 公式。
- `Combined View`：列为 `Product_Name, Total_Sold, Current_Stock, Revenue, Stock_Status`。

`Pipeline_Dashboard.xlsx` 只含 `Pipeline Dashboard` 工作表，数据与 `Combined View` 完全相同，用于离线替代原题 Google Sheet。两份工作簿须保留可编辑单元格、清晰表头、冻结首行、筛选/原生表格、合理列宽和货币格式。

`progress_note.json` 沿用 `memory_seed.json` 的 `entities / relations` 结构，记录读入、去重、规范化、聚合、Excel 和离线 Dashboard 的完成情况以及清洗后的销售/库存数量。不得编造成本节省、业务效果、收入增长或输入未提供的质量结论。
