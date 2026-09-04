# 跨系统财务审计对账工作簿

请使用输入目录中的 `authoritative_ledger.csv`、`q1_manual_tracking.xlsx`（工作表 `Manual Tracking`）和 `system_mapping_guide.pdf` 对账。订单只有在两边都存在且 `Order ID` 精确匹配、日期差不超过 1 天、金额完全相同、两边状态均为不区分大小写的 `completed` 或 `pending` 时才标记为 `matched`，否则标记为 `discrepancy`。所有状态、金额、订单号和差异原因必须从输入重算，不要使用网络信息。

只交付 `/app/output/financial_audit_reconciliation.xlsx`，且输出目录只能包含这个 XLSX 文件。工作簿按以下顺序包含七个可编辑工作表：

1. `Source Ledger`：复制权威台账的四列 `Order ID`、`Date`、`Amount`、`Status`。
2. `Manual Tracking`：复制输入工作表的五列原始记录，不改变顺序和内容。
3. `Reconciliation`：按手工表顺序列出 `Order ID`、`Match Status`、`Reason`。状态只能是 `matched` 或 `discrepancy`；匹配时 Reason 可为空，差异时说明 `missing from authoritative ledger`、`date mismatch`、`amount mismatch` 或 `invalid status`。
4. `Summary`：列出 `Metric`、`Value`，至少包括 `Total Orders`、`Matched`、`Discrepancies`、`Total Amount`。总金额是手工表 Amount 的合计并保留两位小数；摘要使用可重算公式或等价的可编辑计算。
5. `Discrepancy Details`：每个差异订单一行，列为 `Order ID`、`Reason`、`Manual Date`、`Ledger Date`、`Manual Amount`、`Ledger Amount`；缺失台账记录的 Ledger 字段可为空。
6. `Audit Email Draft`：一行离线邮件草稿，列为 `To`、`Subject`、`Body`、`Status`。To 必须为 `finance@company.com`，Subject 必须包含精确短语 `Financial Audit Reconciliation`，正文包含匹配数、差异数和全部差异订单，Status 为 `Draft - Not Sent`。
7. `Review Meeting Draft`：一行离线会议草稿，列为 `Meeting Title`、`Start`、`End`、`Description`、`Status`。标题包含 `Financial Audit Reconciliation Review`，结束时间晚于开始时间，说明包含匹配数、差异数和全部差异订单，Status 为 `Draft - Not Created`。

工作簿应启用自动完整重算，保留原生表格、冻结窗格和可读格式。不要把 DOCX、EML、ICS、PDF 或截图作为输出。
