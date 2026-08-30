# 跨系统财务审计对账

你是财务审计分析师。输入目录包含三个来源：

- `authoritative_ledger.csv`：网店订单的权威台账，字段为 `Order ID`、`Date`、`Amount`、`Status`。
- `q1_manual_tracking.xlsx`：财务部门手工登记的并行记录，使用 `Manual Tracking` 工作表。
- `system_mapping_guide.pdf`：系统字段和匹配规则。

请对手工登记表中的每个订单与权威台账进行对账。订单只有在两边都存在并同时满足以下条件时才是 `matched`，否则为 `discrepancy`：`Order ID` 精确匹配；日期差不超过 1 天；金额完全相同；两边状态均以不区分大小写的方式属于 `completed` 或 `pending`。

请在 `/app/output` 中创建以下四个交付物：

1. `Audit_Reconciliation_Report.xlsx`
   - `Reconciliation` 工作表：列名为 `Order ID`、`Match Status`，按手工登记表顺序每个订单一行；状态只能是 `matched` 或 `discrepancy`。
   - `Summary` 工作表：列名为 `Metric`、`Value`，包含 `Total Orders`、`Matched`、`Discrepancies`、`Total Amount` 四行；总金额取手工登记表 `Amount` 的合计并保留两位小数。
2. `Audit_Report.docx`
   - 至少包含 `Executive Summary` 和 `Discrepancies` 两个标题节；摘要说明匹配数和差异数，差异节列出每个差异订单及具体原因。
3. `audit_notification.eml`
   - 作为未发送的离线邮件草稿，收件人为 `finance@company.com`；主题必须包含精确短语 `Financial Audit Reconciliation`；正文简述匹配数、差异数和差异订单，并提及 `Audit_Report.docx`。
4. `audit_review_meeting.ics`
   - 作为可导入日历的单个审计复核会议草稿；标题包含 `Financial Audit Reconciliation Review`，结束时间晚于开始时间，说明中概括匹配数、差异数和差异订单。

所有数字、状态、差异原因和订单 ID 必须在四个交付物中保持一致。Excel 和 Word 应保持原生可编辑，不要以截图替代。
