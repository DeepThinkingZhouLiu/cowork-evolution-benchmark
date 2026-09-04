# 按双键合并两张产品销售表

打开 `product_sales_input.xlsx`。将 `Table1` 和 `Table2` 按 `PRODUCT NAME` 与 `PRODUCT CODE` 两个键执行全外连接，并在新建的 `Output` 工作表中从 A1 输出：

`KEY CODE`、两侧各自的 PRODUCT NAME/PRODUCT CODE/SALES/COGS/PROFIT、`COMBINED PRODUCT NAME` 和 `COMBINED PRODUCT CODE`。

需要保留只出现在任一来源中的产品；同一代码的两侧数据要保持在同一行，合并名称和编码取非空的一侧。使用可更新的原生公式或等价的可编辑查询结果，不修改两个来源表。保存为 `/app/output/product_sales_outer_join.xlsx`。