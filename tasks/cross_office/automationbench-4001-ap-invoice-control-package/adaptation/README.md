# AutomationBench 4001 逐题改造记录

## 实际阅读的上游内容

- AutomationBench Git commit `4a8e1061254004d9dac807054eed33fad7d1ff14`。
- `automationbench/domains/finance/tasks.py` 中 `get_fin_invoice_email_extract_task()` 的完整 Query、`initial_state` 和 9 个 assertions。
- 函数中的 7 封 Gmail 记录、`Pending Invoices`、`Blocked Vendors` 的 3 条记录、`accounts-payable` 中的 GL-88210 更正消息。
- AutomationBench `README.md`、`LICENSE` 以及 Gmail、Google Sheets、Slack 断言实现。

精确上游 Python 副本保留为 `upstream-finance-tasks.py`，4001 函数执行后的结构化快照保留为 `upstream-task-4001.json`，许可证保留为 `LICENSE-AutomationBench-MIT`。

## 人工确定的本题真值

- `ACM-2026-0412`：Acme Supplies LLC，$4,750.00，可处理，低于 $5,000 阈值，到期日 2026-02-27。
- `GL-88210`：Global Logistics Inc.，邮件原金额 $12,340.50，Slack 权威更正后为 $11,340.50，需复核，到期日从周六 2026-02-28 调整到周一 2026-03-02。
- `SP-0091`：ShadowPeak Consulting Group，$8,200.00，有效阻断；外部审计邮件无解除权限，必须拒绝。
- `AS-7744`：Acme Solutions Inc.，$6,100.00，有效阻断；与可处理的 Acme Supplies LLC 是两个不同实体。
- 待处理发票恰好 2 张，合计 $16,090.50；拒绝 2 张；需复核 1 张；需生成 3 封离线草稿。

## 本项目的改造边界

- `materialize_inputs.py` 只把 4001 的公开 SaaS 初始状态静态化为 MBOX、XLSX 和 JSON，并对消息 ID、行数和断言数做题目专属校验；没有批量编造业务数据。
- 原任务操作在线 Gmail/Sheets/Slack 状态；本题人工改造为可离线回放的 Excel 审计底稿、Word 控制简报和 MBOX 草稿。
- 上游要求“email”；为防止基准运行产生外部副作用，适配后只生成 `DRAFT_NOT_SENT` 离线消息，不发送邮件。
- Golden 和 Judge 均针对这一题手工设计。Judge 从三个输入独立重算供应商、金额、阻断、阈值、到期日和通信真值，不比较 Golden 哈希。
- 任务的业务内容是上游明示的合成 fixture，因此 `synthetic=true`；任务边界与交付格式经人工改写，因此 `human_rewritten=true`。
