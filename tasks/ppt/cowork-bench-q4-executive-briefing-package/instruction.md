请根据 `/app/data/input_files/` 中的 Q4 2024 订单快照和客户维表，完成一套给财务总监和管理层使用的季度销售简报包。

这是 Cowork-Bench `ppt-clickhouse-executive` 真实上游任务的离线改造。上游原题要求从 ClickHouse 企业数据仓库读取数据；本题只把这一个动态依赖替换为两份从同一上游 SQL dump 抽取的静态 CSV，分析口径、交付物和文件名不变。

需要连接 `q4_orders_2024.csv` 与 `customer_dimension.csv`，只使用 2024 年 10 月至 12 月的订单，分别按地理地区和客户分群聚合。

在 `/app/output/` 交付下列三个文件：

- `Q4_Executive_Analysis.xlsx`：只包含 `Regional Summary` 和 `Segment Analysis` 两张工作表。

- `Regional Summary` 的列顺序为 `Region`、`Orders`、`Revenue`、`Avg_Order_Value`。`Avg_Order_Value = Revenue / Orders`，保留两位小数；地区按名称字母序排列。

- `Segment Analysis` 的列顺序为 `Segment`、`Orders`、`Revenue`、`Revenue_Share_Pct`。收入占比是该分群收入占季度总收入的百分比，保留一位小数；分群按名称字母序排列。

- `Q4_Executive_Briefing.pptx`：首页标题必须是 `Q4 2024 Sales Executive Briefing`；之后每个地区一页，按地区名字母序排列，至少显示对应地区的订单数和收入；最后一页显示季度总订单数和总收入。演示文稿应为 16:9、面向管理层、可编辑且可直接演示，不得用整页截图代替原生内容。

- `Q4_Executive_Briefing.pdf`：由完成的 PPTX 导出的同版 PDF 归档，页数、顺序和核心数字必须与 PPTX 一致。

不得从外网获取或补造数据。全部数字必须能从两份输入 CSV 重算得到，且三份交付物之间必须保持一致。
