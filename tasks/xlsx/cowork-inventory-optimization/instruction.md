# 库存需求与补货优先级分析

请使用输入目录中的 `store_products.json`、`store_orders.json`、`supplier_catalog.json`、`demand_history.csv` 和 `snapshot_date.txt`，制作一个可编辑的 Excel 工作簿。输入文件是同一固定快照的唯一事实来源，不要联网、不要补造记录，也不要执行邮件、日历或其他外部写操作。

筛选规则和计算规则如下：

- 只分析 `store_products.json` 中 `status` 为 `publish` 且 `current_stock` 大于 0 的商品。
- 只把订单状态为 `completed` 或 `processing` 的订单行计入销量；按 `product_id` 汇总 `line_items.quantity`，没有销量的商品为 0。
- `Daily_Rate = Total_Sales / 90`，显示并保留两位小数。
- `Days_Remaining = Current_Stock / Daily_Rate`；没有销量时填 999；显示并保留一位小数。
- `Needs_Restock` 在 `Days_Remaining < 30` 时为 `Yes`，否则为 `No`。
- 类别历史需求直接复制 `demand_history.csv` 的 Category、Total_Units_Ordered、Order_Count，不要改写俄文类别名称。
- 供应商风险指数为 `Risk_Index = Lead_Time_Days * (5 - Reliability_Score)`，显示两位小数；按风险指数从高到低理解为需要更早复核的供应商。`Risk_Band` 使用：`High`（交付期 >=14 天或可靠性 <4.0）、`Medium`（交付期 >=10 天或可靠性 <4.5）、否则 `Low`。

只交付一个文件：`/app/output/Inventory_Optimization_Report.xlsx`。输出目录中不要放置脚本、日志、PDF、截图或其他格式文件。

工作簿必须按以下顺序包含五个工作表，表头和状态值保留英文：

1. `Stock_Status`：每个符合筛选规则的商品一行，列为 `Product_ID`、`Product`、`Category`、`Current_Stock`、`Total_Sales`、`Daily_Rate`、`Days_Remaining`、`Needs_Restock`。按 Product 名称升序排列。`Daily_Rate`、`Days_Remaining`、`Needs_Restock` 必须是可重算的原生 Excel 公式。
2. `Supplier_Info`：每个供应商一行，列为 `Supplier`、`City`、`Lead_Time_Days`、`Min_Order_Qty`、`Reliability_Score`、`Risk_Index`、`Risk_Band`。前三个事实字段和可靠性直接取供应商目录，风险列使用上述公式和规则。
3. `Demand_History`：列为 `Category`、`Total_Units_Ordered`、`Order_Count`，保留输入 CSV 的全部类别行。
4. `Restock_Plan`：只列出 `Needs_Restock=Yes` 的商品，按 `Days_Remaining` 升序。列为 `Priority_Rank`、`Product_ID`、`Product`、`Category`、`Current_Stock`、`Total_Sales`、`Days_Remaining`、`Recommended_Action`。优先级从 1 开始连续编号；商品和数值应链接或重算自 `Stock_Status`，整改动作写为明确的离线补货建议。
5. `Restock_Summary`：列为 `Metric`、`Value`，至少包含 `Total_Products_Analyzed`、`Products_Need_Restock`、`Products_Healthy`、`Avg_Days_Remaining`、`Snapshot_Date`。前四项必须使用可重算公式并与 `Stock_Status` 一致，日期取 `snapshot_date.txt`。

保留原生单元格、公式、表格结构和冻结窗格；启用自动完整重算。完成后只在指定输出文件中保存结果。
