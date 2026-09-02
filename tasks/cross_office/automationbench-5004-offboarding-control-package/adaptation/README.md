# AutomationBench 5004 逐题改造记录

## 实际阅读的上游内容

- AutomationBench Git commit `4a8e1061254004d9dac807054eed33fad7d1ff14`。
- `automationbench/domains/hr/tasks.py` 中 `get_employee_offboarding_task()` 的完整 Query、`initial_state` 和 13 个 assertions。
- 函数中的 4 条员工离职记录、4 条控制政策、运行时间和 `departures` Slack 频道。
- AutomationBench `README.md`、`LICENSE` 以及 Gmail、Google Sheets 和 Slack 的断言语义。

精确上游 Python 副本保留为 `upstream-hr-tasks.py`，5004 函数执行后的结构化快照保留为 `upstream-task-5004.json`，许可证保留为 `LICENSE-AutomationBench-MIT`。上游 `answer` 字段为空，没有可直接复用的 Golden。

## 人工确定的本题真值

- 冻结时间为 2026-03-10。Greg Foster 的最后工作日是 2026-03-07，属于当前自愿离职：目标状态为 `PROCESSED`，需要公开告别草稿、按最后工作日撤权草稿和 Payroll 范围复核草稿。
- Diana Reese 的最后工作日是 2026-03-14，属于当前非自愿离职，并带有“HR and legal reviewing this case”备注：目标状态为 `PROCESSED`，不得公开，需立即撤权、通知 HR Director，并只向 Payroll 路由遣散费范围复核。
- Raj Patel 的最后工作日是 2026-04-15，距冻结时间 36 天：目标状态为 `SCHEDULED`，当前不得生成通知。
- Greg Forster 是近似姓名干扰项，原状态已经是 `Processed`：必须保持不动作，不能与 Greg Foster 混淆或重复执行。
- 遣散费由 Payroll 在法律审核后独占处理。HR 只能路由当前两名可操作员工的范围复核，不能计算金额、批准、付款或声称已经完成。
- 最终离线交付恰好包含 4 封邮件草稿和 1 条 Slack 草稿；所有通信均必须明确为未发送或未发布。

## 本项目的改造边界

- `materialize_inputs.py` 只执行并静态化 5004 的公开初始状态，并对 example ID、任务函数、记录规模、政策、时间、频道和 assertions 做题目专属断言；没有使用通用脚本批量编造业务数据。
- 原任务操作在线 Gmail、Google Sheets 和 Slack；本题人工改造为可离线回放的 XLSX 决策台账、DOCX 控制简报、MBOX 邮件草稿和 Slack JSON 草稿。
- Golden 和 Judge 均针对这一题逐项设计。Judge 从三个冻结输入独立重算日期差、目标状态、公开/保密边界、通知对象、遣散费权限和四产物一致性，不比较 Golden 哈希。
- 逐题回放发现并修复了两个真实缺陷：其一，Excel 公式用子串搜索会把 `Involuntary` 误判为 `Voluntary`；其二，Word Judge 的全局关键词检查无法约束 Raj Patel 的实体级时间事实。最终公式改为前缀判断，Judge 改为逐实体核验，并分别增加针对性负例。
- 上游内容是 AutomationBench 明示的合成 fixture，因此 `synthetic=true`；任务边界和交付形式经人工改写，因此 `human_rewritten=true`。
