请使用工作区中的两份静态快照，完成支持中心 SLA 违约审计，并将唯一交付保存为 `/app/output/SLA_Breach_Audit.xlsx`。

这是 Cowork-Bench `sf-sla-breach-audit` 的离线静态改造版。上游原题要求查询 ClickHouse 并发送邮件；本题把同一上游数据库的表逐行固化为 CSV，并以工作簿内的可回放升级草稿替代真实发送。

输入文件的责任如下：

- `support_tickets.csv`：上游 `SUPPORT_CENTER__PUBLIC__TICKETS` 的 31,588 条完整快照。
- `sla_policies.csv`：上游 `SUPPORT_CENTER__PUBLIC__SLA_POLICIES` 的 4 条 SLA 规则。

请按 `PRIORITY` 连接两个数据集。对每个实际存在工单的优先级，若 `RESPONSE_TIME_HOURS` 严格大于对应的 `RESPONSE_TARGET_HOURS`，则计为一次违约。

Excel 必须恰好包含两张可编辑工作表：

- `Breach Analysis`：列为 `Priority, Total_Tickets, Avg_Response_Hours, SLA_Target_Hours, Breached_Count, Breach_Rate`。`Avg_Response_Hours` 和以百分数数值表示的 `Breach_Rate` 均四舍五入到两位小数；按 `SLA_Target_Hours` 升序排列；只包含实际有工单的优先级，优先级名保持英文原值。
- `Summary`：列为 `Metric, Value`，依次包含 `Total_Tickets, Total_Breached, Overall_Breach_Rate, Worst_Priority, Best_Priority`。总违约率也使用百分数数值并四舍五入到两位小数；最差/最佳优先级分别按违约率最高/最低确定。

工作簿应可筛选、可阅读、数字格式清楚，不得把结果压平成图片。`Breach_Rate` 以及 `Summary` 的三个汇总数字应用 Excel 公式从分组结果推导，便于审计者复核。

工作簿还必须包含 `Escalation Draft` 工作表，作为上游邮件动作的离线草稿，必须包含以下四项：

- `From: sla-audit@company.com`
- `To: ops-director@company.com`
- `Subject: SLA Breach Audit Results`
- 正文明确写出 `Overall_Breach_Rate` 和 `Worst_Priority`，并与 Excel 一致。
