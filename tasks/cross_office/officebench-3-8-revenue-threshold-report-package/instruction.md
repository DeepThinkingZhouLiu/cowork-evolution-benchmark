你是公司的收入分析专员。请读取 `/app/input_files/company_revenues.xlsx` 中 2004–2023 年的全部收入记录，执行 200 万美元门槛筛选，并交付一个可复算的 Excel 分析工作簿和一个与工作簿完全一致的 Word 管理报告。不得访问网络，不得补充输入中没有的业务背景或原因。

筛选规则如下：

- 收入大于或等于 `$2,000,000` 的年份保留。
- 收入严格低于 `$2,000,000` 的年份移除。
- 年份按升序排列，金额按原始整数精度计算。
- 报告中的趋势只能描述输入数据，不得把变化归因于通胀、市场、产品、管理层或其他未提供因素。

在 `/app/output` 中交付且只交付以下两个文件：

1. `Revenue_Filtered_Analysis.xlsx`，依次包含五个工作表：
   - `Source_Data`：逐行保留输入的 20 条原始记录。
   - `Threshold_Audit`：对全部 20 年列出 Year、Revenue_USD、Threshold_USD、Status、Margin_to_Threshold 和 Revenue_Rank；Status、Margin_to_Threshold 和 Revenue_Rank 必须使用可复制的原生公式。
   - `Retained_Years`：只保留符合门槛的年份，列出 Year、Revenue_USD、Margin_to_Threshold、Change_vs_Previous_Retained 和 Change_Pct_vs_Previous_Retained；除第一条没有前序保留年份外，后三个计算字段使用原生公式。
   - `Removed_Years`：只列出低于门槛的年份及 Shortfall_to_Threshold，Shortfall 使用原生公式。
   - `Summary`：用原生公式给出 Original_Years、Retained_Years、Removed_Years、Retention_Rate、Retained_Total、Retained_Average、Highest_Revenue_Year、Highest_Revenue、Lowest_Retained_Year、Lowest_Retained_Revenue、First_to_Last_Change 和 First_to_Last_Change_Pct，并提供一个可编辑的收入趋势图。
2. `Revenue_Threshold_Report.docx`，标题为 `Revenue Threshold Screening Report`，至少包含 `Screening Rule`、`Executive Summary`、`Removed Years`、`Retained Revenue Findings` 和 `Evidence Boundary` 五节。报告必须用原生表格列出三个被移除年份及其收入和门槛差额，并准确呈现保留数量、保留率、保留收入总额与平均值、最高收入年份、最低保留年份以及首个保留年份到最后一个保留年份的金额和百分比变化。

两个文件中的年份、金额、比例、筛选结论和趋势结论必须相互一致；Excel 公式和 Word 表格都必须可继续编辑，不得用截图替代。
