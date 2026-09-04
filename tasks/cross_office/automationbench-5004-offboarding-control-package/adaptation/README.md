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

## 评分口径 v2（2026-09-03）

- 原严格 `rubric.json`/`tests/judge.py` 保存在 `rubric.v1.json`/`tests/judge_v1.py`；当前文件为 v2，`tests/test.sh` 显式选择 v2。
- v1 将任一员工、政策、公式、通知或跨文件局部错误设为 hurdle，并叠加最高 30 分的重复 penalty。v2 只在 XLSX、DOCX、MBOX、Slack JSON 任一核心交付缺失、为空或损坏时归零；其余按 4 名员工、4 条政策、62 个公式、28 个邮件字段、16 个跨文件事实、结构和渲染组件给连续部分分。
- v2 对 v1 聚合 detail 做本题专属原子映射：例如缺少一封邮件按该邮件的字段组扣分，工作簿不可用才扣完整公式组；公开 Diana、提前通知 Raj、重复处理 Greg Forster、两名 Greg 混淆、遣散费越权和虚假执行再按实际违规类别处罚，避免同一错误重复扣满。
- 实测 Golden 与 fresh replay 均为 `1.000000`；缺失、损坏、仅原样输入三类不可用交付均为 `0.000000`；30 个有效变体均为部分分，范围 `0.672043–0.994167`。代表性中等样例：近似姓名混淆 `0.672043`；删除一封 MBOX 草稿 `0.961696`；向 Payroll 草稿加入遣散金额 `0.899643`；公开 Diana 或虚假发布约 `0.920000–0.940000`。
- 当前仓库内只有 oracle 的 Cross-Office 回放，没有可定位的 Claude Code、OpenCode、OpenHands 候选 artifacts，因此没有伪造三套 Harness 的重算差异。
- 剩余风险：openpyxl 保存 Excel 变体会清除公式缓存，导致部分 Excel 负例呈现真实级联错误；报告按实际受影响单位计分，并明确不把它们称作单字段独立扰动。
