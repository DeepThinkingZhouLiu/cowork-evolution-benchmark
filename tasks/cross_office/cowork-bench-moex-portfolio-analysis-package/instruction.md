你是一家资产管理公司的投资组合分析师。请根据 `/app/input_files` 中的三份冻结材料，为投资委员会完成月度投资组合分析包。不得访问网络，不得用其他日期或实时行情替换冻结快照。

输入如下：

- `Portfolio.xlsx`：5 个 MOEX 持仓的代码、公司和数量。
- `moex_market_snapshot.csv`：截至 2026-05-26 的发行人信息及 30 个日历日窗口行情。
- `Investment_Policy.pdf`：基金的行业、币种、交易所、Large Cap 和报告要求。

计算规则：

- `Current_Price` 使用每个代码最新的 2026-05-26 收盘价。
- 30 日基准价使用“不晚于 2026-04-26 的最新收盘价”，本快照中应为 2026-04-24。
- `Current_Value = Quantity * Current_Price`。
- `Return_30d_Pct = Current_Price / Baseline_Close - 1`，在 Excel 中以百分比显示。
- `Weight_Pct = Current_Value / Total_Portfolio_Value`。
- 市值严格大于 1000 亿 RUB 时标记 `Large Cap`。不得自行添加政策中没有的持仓或行业上限。

在 `/app/output` 中交付且只交付以下三个文件：

1. `Portfolio_Analysis.xlsx`，包含三个工作表：
   - `Holdings`：每个持仓一行，至少包含 Symbol、Company、Sector、Quantity、Baseline_Date、Baseline_Close、Current_Date、Current_Price、Current_Value、Return_30d_Pct、Weight_Pct、Market_Cap_RUB、Market_Cap_Category、Currency、Exchange 和 Policy_Status。Current_Value、Return_30d_Pct、Weight_Pct、Market_Cap_Category 和 Policy_Status 必须是可复制的原生公式。
   - `Summary`：用原生公式给出 Total_Portfolio_Value、Best_Performer、Worst_Performer、4 个实际持仓行业的市值、Largest_Sector、Largest_Sector_Weight 和 Policy_Compliant_Positions。GAZP.ME 与 LKOH.ME 必须合并到同一个 Energy 行业。
   - `Source_Prices`：完整保留输入 CSV 的 115 行市场快照，便于追溯。
2. `Investment_Memo.docx`，标题为 `Portfolio Analysis Report`，包含 `Holdings Overview`、`Performance Analysis`、`Policy Review` 和 `Recommendations` 四节。用原生表格呈现 5 个持仓；正文需给出总市值、最佳/最差代码及收益率、最大行业暴露和政策检查；至少给出两条只基于快照和政策的后续建议。不得把同期变化写成未被输入支持的因果解释。
3. `Monthly_Portfolio_Analysis.eml`，发件人为 `analyst@fund.example.com`，收件人为 `investment-committee@fund.example.com`，主题为 `Monthly Portfolio Analysis Report`。邮件正文需准确给出总市值、最佳/最差代码和收益率，并真正嵌入本题生成的 Excel 和 Word 两个附件。

三个文件中的数字、代码、行业和结论必须相互一致；所有 Office 内容必须可继续编辑，不得用截图替代表格或公式。
