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

## 评分口径 v2（2026-09-03）

- 原严格口径保存在 `rubric.v1.json` 与 `tests/judge_v1.py`，当前 `rubric.json`/`tests/judge.py` 为 v2，`tests/test.sh` 显式调用 v2。
- v1 把任一发票、公式、邮件或跨文件局部错误设为 hurdle，并叠加 25–30 分 penalty，真实 Agent 即使完成大部分内容也常被压到 0。v2 只有 XLSX、DOCX、MBOX 任一缺失、为空或损坏才归零，其余按 4 张发票、11 个权限/政策单位、42 个公式、18 个邮件字段、跨文件事实和原生结构分别给部分分。
- v2 对 v1 聚合 detail 使用本题专属原子映射。例如表头或工作簿不可用会扣除整组单位，而单张发票、单条政策、单个公式或单封邮件字段只扣对应单位；不再简单按 detail 条数或统一乘系数。
- 实测 Golden 与 fresh replay 均为 `1.000000`；核心文件缺失、损坏、仅原样输入均为 `0.000000`；25 个有效变体均得到严格位于 0 与 1 的分数，范围 `0.791668–0.997500`。代表性中等样例：删除一封 MBOX 草稿 `0.812500`；局部样例：邮件汇总金额错误 `0.996111`；越权/虚假声明样例落在 `0.898977–0.956667`。
- 当前仓库内未定位到 Claude Code、OpenCode、OpenHands 的 Cross-Office 候选 artifacts，只有 oracle 运行，因此没有人为制造三套 Harness 差异；待候选产物落盘后可直接用 v2 重算。
- 剩余风险：个别 Excel 变体会经 LibreOffice 重新计算而产生真实的级联差异，报告按受影响的原子单位计分，不将其描述为单字段误差。
