你是质量管理团队的退款调查分析师。请基于工作区中的真实来源快照，完成一次可回放的退款根因分析，并保证 JSON、Excel、Word 和两封邮件草稿中的核心事实一致。

先阅读 `Quality_Policy.pdf` 和 `guide.md`。两份文件定义了严重度、评论关联、根因分类和供应商调查阈值。其他输入为：

- `refunds.csv`：上游 InSales 店铺的全部 13 条退款记录；退款金额已忠实执行上游 preprocess 的全额退款对齐。
- `refund_orders.json`：与这些退款相关的 13 个订单及原始 `line_items`。
- `products.csv`：上游店铺的 82 个商品投影。
- `product_reviews.csv`：上游店铺的全部 396 条商品评论。

请按以下规则分析：

- 每条退款按 Refund_ID 升序，并用订单 `line_items` 中的所有商品名用英文分号和空格 `; ` 连接。
- 退款金额大于 50 为 `Critical`；20 到 50（含 20）为 `Major`；低于 20 为 `Minor`。
- `Date` 取 `date_created` 中 UTC 时间戳的前 10 个字符，即 `YYYY-MM-DD`。
- 一条退款涉及多个商品时，将该退款金额在订单商品之间等分；先对商品累计，再以金融常用的 `ROUND_HALF_UP` 保留两位小数。
- 平均评分使用 `product_reviews.csv` 中该 `product_id` 的全部评分，以 `ROUND_HALF_UP` 保留两位小数。低于 3.0 为 `Quality Issue`，3.0 及以上为 `Service Issue`，无评论为 `No Reviews`。
- 同一商品有 2 条及以上退款时，`Requires_Investigation` 为 `Yes`，否则为 `No`。
- 根因映射必须与 `guide.md` 一致：`Product defective -> Manufacturing Defect`，`Product arrived damaged -> Shipping Damage`，`Wrong item received -> Wrong Item`，`Changed my mind` 或 `Item not as described -> Customer Expectations`，`Duplicate order` 及其他原因 `-> Other`。

请在输出目录交付且仅交付以下六个核心文件：

- `refund_root_cause_processor.py`：可回放处理脚本，只使用命令行指定的输入，并支持 `python refund_root_cause_processor.py --input-dir <输入目录> --output <JSON 路径>`。
- `analysis_results.json`：顶层键为 `refund_details`、`product_impact`、`summary`、`refund_reason_breakdown`、`root_cause_breakdown`、`investigation_products`。其字段须与后述 Excel 口径一致。
- `Refund_Analysis.xlsx`：原生可编辑 Excel，恰好包含 `Refund Details`、`Product Impact`、`Summary` 三张表。
- `Root_Cause_Report.docx`：原生可编辑 Word 报告。
- `quality_team_email.txt`：发给 `quality_team@company.com` 的离线邮件草稿。
- `supplier_relations_email.txt`：发给 `supplier_relations@company.com` 的离线邮件草稿。

Excel 要求：

- `Refund Details` 表头依次为 `Refund_ID`、`Order_ID`、`Product_Name`、`Refund_Amount`、`Severity`、`Refund_Reason`、`Date`，按 Refund_ID 升序。
- `Product Impact` 表头依次为 `Product_Name`、`Refund_Count`、`Total_Refund_Amount`、`Avg_Review_Rating`、`Review_Count`、`Issue_Type`、`Requires_Investigation`，先按 Refund_Count 降序，再按 Product_Name 字母序。
- `Summary` 的 A/B 列为 `Metric`/`Value`，按此顺序给出 `Total_Refunds`、`Total_Refund_Amount`、`Critical_Count`、`Major_Count`、`Minor_Count`、`Products_Affected`、`Quality_Issues`、`Service_Issues`、`Products_Requiring_Investigation`。
- 使用原生单元格、专业表头、冻结行、过滤或原生表格、合理列宽和货币/小数格式；不得把表格压平成图片。

Word 报告必须准确包含以下六个章节标题：`Executive Summary`、`Refund Trend Analysis`、`Product-Level Findings`、`Root Cause Classification`、`Recommendations`、`Action Plan`。报告要写出核心汇总、退款原因计数、根因计数、所有需调查商品的退款数和分摊金额，并给出与所提供证据相匹配的建议与行动计划。

质量团队邮件主题需同时含 `Refund` 和 `Analysis`，正文需包含退款总数、总金额、严重度拆分和关键问题。供应商关系邮件主题需同时含 `Supplier` 和 `Investigation`，正文需逐一列出需调查商品及其退款数和分摊金额。本题是离线 Benchmark，不要实际发送邮件。

所有结论必须能追溯到所给输入。输入不支持收入影响、客户流失、行业基准、供应商责任比例、成本节省、ROI 或时间趋势方向；不得编造这些数字或断言。
