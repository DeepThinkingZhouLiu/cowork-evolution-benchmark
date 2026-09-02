# AutomationBench 5111 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5111()`，`example_id=5111`，任务名 `hr.pto_reconciliation`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`d5851e995d8ea47a1b7e720e02bfd47f957c5b28cea6113383652051dac1e320`。

制作时完整读取了原 Query、运行时间、八条 PTO 余额、三条政策、Payroll 修正邮件、HR Director 隐私邮件、两个 Slack 频道和十一条上游 assertions。没有连接或执行 AutomationBench 外部服务，也没有真实修改台账、发送邮件或发布 Slack。

## 上游事实审计

Payroll 邮件只授权修正 Rita Johansson 和 Sam Okafor。HR Director 邮件又明确规定个人 PTO 余额保密，禁止在公司级 Slack 发布姓名或余额；汇总数字只能进入私密的 `#hr-general`，员工应单独邮件通知。

| 源行 | 员工             | 原余额 | 修正后 | 变化 | 获批透支 | 最终状态    | 更新 | 通知 | HR 升级 |
|-----:|------------------|-------:|-------:|-----:|----------|-------------|------|------|---------|
|    2 | Quinn Torres     |     12 |     12 |    0 | 否       | `Current`   | 否   | 否   | 否      |
|    3 | Rita Johansson   |      8 |      3 |   -5 | 否       | `Corrected` | 是   | 是   | 否      |
|    4 | Sam Okafor       |      3 |     -2 |   -5 | 否       | `Escalated` | 是   | 是   | 是      |
|    5 | Tina Marchetti   |     15 |     15 |    0 | 否       | `Current`   | 否   | 否   | 否      |
|    6 | Uma Krishnan     |      0 |      0 |    0 | 否       | `Current`   | 否   | 否   | 否      |
|    7 | Victor Pham      |      7 |      7 |    0 | 否       | `Current`   | 否   | 否   | 否      |
|    8 | Wendy Liu        |     -1 |     -1 |    0 | 是       | `Current`   | 否   | 否   | 否      |
|    9 | Xavier Dunn      |      9 |      9 |    0 | 否       | `Current`   | 否   | 否   | 否      |

Sam 的修正后余额为负且没有获批透支备注，因此是唯一 HR Director 升级。Wendy 虽为负余额，但她没有本轮修正且源行明确包含获批透支备注，所以必须保持 `Current`，不能更新、通知或升级。总计八人、两项修正、六人未变、净变化 `-10` 天、两封员工草稿和一封 HR 升级草稿。

频道状态存在一个必须显式处理的冲突：政策说 `#hr-general` 应为私密频道，但冻结 Slack 快照中 `hr-general` 与 `company-announcements` 的 `is_private` 都是 `false`。因此公司公告被 `REFUSED_PRIVACY`，HR 汇总被 `WITHHELD_CHANNEL_NOT_PRIVATE`；只保留聚合草稿 `2 PTO corrections processed; 1 HR escalation queued.`，真实发布数为零。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5111，并逐项断言八条余额、三条政策、两封邮件、两个频道、冻结时间和十一条 assertions；它不是通用造题器。
- Google Sheets 静态化为 `PTO_Balance_Tracker.xlsx`，完整保留 Q2 余额、源行、顺序、状态、Notes 和政策全文。
- Gmail 静态化为 `PTO_Corrections_and_Privacy.mbox`，保留两封邮件全文、邮件 ID、线程 ID、日期和标签。
- Slack 与运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游要求真实更新、邮件和 Slack 通知，本离线 Benchmark 改为显式草稿，并新增公式审计、Word 控制简报、可编辑 PPT 复核和隐私控制 JSON；没有增加或改写员工、邮箱、余额、Notes、修正或政策事实。

## Golden 与 Judge 设计

- Golden 的 `solve.py` 明确解析本题四个冻结输入，并独立派生八人逐行决策；没有随机数据或跨题通用业务生成逻辑。
- Excel 包含九张原生表。修正范围、修正余额、变化量、获批透支、负余额、最终状态、三个动作路由、两个 Slack 门控、回落值、来源追踪和十三项汇总均使用原生公式，并保留条件格式和可编辑图表。
- Word 为三页、五张原生表；PPT 为三页，每页均含可编辑原生表，覆盖摘要、八人完整决策、Wendy 保护、Sam 升级、两个频道门控和零真实执行。
- Judge 不比较 Golden 哈希。它重新读取四个输入与 SHA256，独立重建八人真值，再核验公式依赖及缓存、三类 Excel 队列、Tracker JSON、单收件人 MBOX、Sam-only 升级 JSON、Slack 隐私控制、Word/PPT 原生结构、渲染页数和七文件一致性。
- 关键失败使用 hurdle cap；来源/金额/状态/路由错误、Wendy/Sam/未修正人员控制错误、隐私或虚假执行、文件/公式/追踪/跨文件错误分别设置独立 penalty。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 9 页，Word 为 3 页，PPT 为 3 页；15 页均逐页目视检查，未发现重叠、截断、表格越界、空白页或公式错误。
- 165 个定向负例全部被拒绝，最高得分 `0.05`，低于 `0.8` 通过线。覆盖文件缺失与损坏、八人来源与政策篡改、公式静态化、Rita/Sam/Wendy 分支错误、未修正人员误路由、错误收件人、Cc/Bcc、虚假执行、公开频道泄露、Tracker/MBOX/JSON/Word/PPT 跨文件冲突和原生结构删除。
- 负例首先暴露三处 Judge 盲点：MBOX 顺序未检查、Word 重复门控文本可掩盖单行篡改、缺少 Rita 的 Tracker JSON 会使一致性检查短路。Judge 已分别增加顺序核验、路由表精确行核验和缺项安全一致性检查，重跑后 `165/165` 拒绝。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 的 job `cross-office-automationbench-5111-pto-reconciliation-final-fixed-20260902` 与 `cross-office-automationbench-5111-pto-reconciliation-final-repeat-20260902` 均为 Reward `1.0`、0 异常；任务摘要均为 `efe8a46823a0fbc772e261a4776f5a793d6754f872f6e4a26f1a94ef144d1d5b`，两个容器 trial 的七份产物逐字节一致。
- 正式难度暂不赋值。本题把多源权威冲突、数值对账、政策优先级、负向筛选、隐私拒绝、七交付联动和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。
