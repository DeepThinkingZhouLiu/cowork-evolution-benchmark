请基于工作区中的两份上游数据，完成客户满意度综合分析，并将四个交付物保存为：

- `/app/output/satisfaction_analysis.xlsx`
- `/app/output/action_plans.xlsx`
- `/app/output/satisfaction_report.docx`
- `/app/output/executive_summary.docx`

这是 Cowork-Bench `customer-satisfaction-analytics-framework` 的逐题离线改造版。上游原题和两份输入文件保持不变，但评测会从输入重算所有数字，不接受没有数据支持的样本量、行业基准、相关性、预算或收入影响。

输入责任：

- `survey_template.xlsx`：`Responses` 表包含已收集的客户问卷；`Survey Template` 和 `Demographics` 提供问题和分群语义。
- `support_data_export.csv`：客服工单快照，包含解决时长、首次解决和客户满意度。

计算口径：

- `Overall Satisfaction Score`：`Responses.Overall_Satisfaction` 的算术平均值。
- `Net Promoter Score (NPS)`：`Recommend_Score >= 9` 为推荐者，`Recommend_Score <= 6` 为贬损者，NPS = 推荐者占比 - 贬损者占比，乘以 100。
- `First Contact Resolution Rate`：使用 `support_data_export.csv` 中 `First_Contact_Resolution=Yes` 的占比。
- `Average Resolution Time (Hours)`：使用 `support_data_export.csv` 中 `Resolution_Time_Hours` 的平均值。
- 维度评分：分别计算 `Product_Quality`、`Customer_Service`、`Value_for_Money`、`Ease_of_Use` 和 `Recommend_Score` 的平均值。
- 客户分群：对 `Enterprise / Mid-Market / SMB` 分别计算问卷数、平均满意度、NPS 和问卷中的首次解决率。

`satisfaction_analysis.xlsx` 只保留一张 `Satisfaction Analysis` 工作表，并按以下三个表格组织：

- `Summary Metrics`：表头为 `Metric, Value, Source`，包含上述四项核心指标。首次解决率使用 Excel 百分比数值，其余数字保留两位小数。
- `Dimension Ratings`：表头为 `Dimension, Average, Scale`，包含五个维度。
- `Segment Analysis`：表头为 `Customer_Type, Respondents, Overall_Satisfaction, NPS, First_Contact_Resolution_Rate`，包含三个分群。

`action_plans.xlsx` 只保留一张 `Action Plans` 工作表，表头必须为 `Initiative, Owner, Target Metric, Baseline, Goal, Timeline, Status`。至少覆盖以下三类行动：

- 提升首次解决率，基线必须来自客服数据，目标高于基线。
- 缩短平均解决时长，基线必须来自客服数据，目标低于基线。
- 改善 SMB 满意度，基线必须来自问卷分群，目标高于基线。

`Owner` 请写职责角色，不要虚构人名。每条行动都必须有可量化基线和目标、时间范围与状态。

两份 Word 文档的正文使用俄文：

- `satisfaction_report.docx`：详细说明数据范围、方法、核心指标、五个维度、三个分群、行动建议和限制。
- `executive_summary.docx`：简洁呈现四项核心指标、SMB 是满意度最低分群的结论、三个优先行动和数据限制。

两份报告必须明确写出真实样本量，并说明两份数据无法按客户逐条关联，因此不做个体层面相关性或因果推断。所有核心数字必须与两份 Excel 一致。
