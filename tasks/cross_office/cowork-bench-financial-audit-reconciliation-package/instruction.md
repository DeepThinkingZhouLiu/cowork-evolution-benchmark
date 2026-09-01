你是 ООО «Сибирская торговая компания» 的财务审计分析师。请依据以下三份冻结输入，完成 Q1 订单跨系统对账并交付一套可继续编辑、可复算、内部事实完全一致的审计包：

- `/app/input_files/authoritative_ledger.csv`：网店订单权威台账。
- `/app/input_files/q1_manual_tracking.xlsx`：财务部门手工台账，数据位于 `Manual Tracking`。
- `/app/input_files/system_mapping_guide.pdf`：字段含义和 Matching Rules。

先阅读规则 PDF，再以 `Order ID` 将手工台账逐行与权威台账关联。对账规则为：订单号精确一致；日期相差不超过 1 天；金额必须完全一致；两侧状态均须是 `completed` 或 `pending`，比较时忽略大小写。任一条件不满足，或权威台账缺少对应订单，均为 `discrepancy`。

请在 `/app/output` 中且仅在该目录交付以下四个文件：

- `Audit_Reconciliation_Report.xlsx`
- `Audit_Report.docx`
- `Financial_Audit_Reconciliation.eml`
- `Audit_Reconciliation_Review.ics`

`Audit_Reconciliation_Report.xlsx` 必须是原生可编辑工作簿，且恰好包含以下三张工作表：

- `Reconciliation`：每个手工台账订单一行，按 `Order ID` 升序；表头严格依次为 `Order ID`、`Manual Date`、`Ledger Date`、`Date Delta Days`、`Manual Amount`、`Ledger Amount`、`Amount Delta`、`Manual Status`、`Ledger Status`、`Presence Check`、`Date Check`、`Amount Check`、`Status Check`、`Match Status`、`Discrepancy Reason`。日期须为真实 Excel 日期，金额使用 RUB 两位小数。权威记录缺失时，对应 Ledger 字段和数值差异留空。`Date Delta Days`、`Amount Delta` 以及所有 Check、`Match Status`、`Discrepancy Reason` 必须使用可向下复制的原生公式，不得把结果硬编码成静态文本。
- 检查列统一使用 `pass` / `fail`；`Match Status` 统一使用 `matched` / `discrepancy`。`Discrepancy Reason` 只使用 `none`、`missing_in_authoritative_ledger`、`date_mismatch`、`amount_mismatch`、`invalid_status`，多个原因用英文分号连接。
- `Exception Register`：只保留异常订单，表头严格依次为 `Order ID`、`Exception Type`、`Manual Evidence`、`Ledger Evidence`、`Impact`、`Required Follow-up`。每行必须写清手工值、权威值、差异或缺失事实，以及不超出输入证据的后续动作；不得把缺失订单直接断言为欺诈、取消或无效交易。
- `Summary`：A/B 列依次为 `Metric` / `Value`，按顺序给出 `Total Manual Orders`、`Matched`、`Discrepancies`、`Manual Total Amount`、`Authoritative Amount Represented`、`Missing Authoritative Amount`、`Absolute Amount Mismatch`、`Date Mismatches`、`Amount Mismatches`、`Missing Records`、`Control Conclusion`。Value 列全部使用引用 `Reconciliation` 的原生公式；结论在无异常时为 `Reconciled`，否则为 `Needs remediation`。
- 三张表均需使用原生表格或筛选、冻结表头、合理列宽、专业数字格式和异常高亮；不得用截图代替单元格。

`Audit_Report.docx` 必须是正式、原生可编辑的审计报告，使用 Word 标题样式和原生表格，并包含以下五个一级标题：

- `Executive Summary`
- `Scope and Methodology`
- `Reconciliation Results`
- `Discrepancies and Required Actions`
- `Evidence Boundary`

报告必须准确说明总订单数、匹配数、异常数、手工台账总金额，以及每个异常订单的类型、两侧证据、差异和合理后续动作；同时明确权威订单台账是订单事实来源、日期容差为 1 天、金额容差为 0、状态比较忽略大小写。报告中的所有数字和订单结论必须与 Excel 一致。不得声称已完成整改、财务报表无重大错报、缺失订单属于欺诈或已经取消，也不得编造收入影响、罚款、客户流失、责任比例或 ROI。

`Financial_Audit_Reconciliation.eml` 是离线邮件草稿，不要实际发送。它必须是合法 MIME 邮件：

- 收件人为 `finance@company.com`，主题严格为 `Financial Audit Reconciliation - Q1 2024`。
- 正文简要写出总订单数、匹配数、异常数、手工台账总金额，并逐一列出异常订单及类型。
- 以 MIME 附件形式同时附上本次交付的 XLSX 和 DOCX；附件内容必须与输出目录中的两个文件逐字节一致。

`Audit_Reconciliation_Review.ics` 是离线会议邀请，不要连接真实日历服务：

- 使用 `Asia/Shanghai`，会议时间为 2024-04-05 10:00–10:45。
- `SUMMARY` 为 `Audit Reconciliation Review`，状态为 `TENTATIVE`，参会人为 `finance@company.com`。
- `DESCRIPTION` 写明总订单数、匹配数、异常数和全部异常订单号，并用相对 URI 引用两份 Office 交付物；不得出现本机绝对路径。

请自行核验四个文件能够正常打开、Excel 公式可以重算、邮件附件可解包、会议邀请可导入，并确保四个文件中不存在互相冲突或仍使用旧值的结论。
