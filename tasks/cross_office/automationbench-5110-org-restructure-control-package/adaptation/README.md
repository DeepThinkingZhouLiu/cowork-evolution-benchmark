# AutomationBench 5110 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5110()`，`example_id=5110`，任务名 `hr.org_restructure_update`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`bfeccf5c4c7d44c0d5f083091dcea02b46be4299b10438c39bc1a9da74211848`。

本题制作时完整读取了原 Query、`initial_state.meta`、Google Sheets 八条员工目录、VP People 批准邮件全文、Slack 频道、九条上游 assertions 和仓库 LICENSE。没有执行 AutomationBench 的外部服务，也没有真实更新员工目录、发送邮件或发布 Slack。

## 上游事实审计

VP People 批准邮件是唯一改线依据，生效日为 `2026-04-14`。逐人真值如下：

| 源行 | 员工           | 部门        | 原经理               | 更新后经理            | 处理结果                         |
|-----:|----------------|-------------|----------------------|-----------------------|----------------------------------|
|    2 | Aisha Jordan   | Engineering | Lisa Wang            | Marcus Chen           | `APPROVED_CHANGE`                |
|    3 | Bo Lindqvist   | Engineering | Marcus Chen          | Derek Huang           | `APPROVED_CHANGE`                |
|    4 | Chiara Russo   | Engineering | Derek Huang          | Lisa Wang             | `APPROVED_CHANGE`                |
|    5 | Dante Obi      | Product     | Lisa Wang            | Marcus Chen-Williams  | `APPROVED_CHANGE`                |
|    6 | Elena Vasquez  | Engineering | Derek Huang          | Derek Huang           | `UNCHANGED_NOT_IN_APPROVED_LIST` |
|    7 | Felix Braun    | Engineering | Marcus Chen          | Marcus Chen           | `UNCHANGED_NOT_IN_APPROVED_LIST` |
|    8 | Grace Kim      | Product     | Marcus Chen-Williams | Marcus Chen-Williams  | `UNCHANGED_NOT_IN_APPROVED_LIST` |
|    9 | Hugo Fernandez | Engineering | Lisa Wang            | Lisa Wang             | `UNCHANGED_NOT_IN_APPROVED_LIST` |

经理邮箱只能从目录中已有记录解析。`Marcus Chen / marcus.chen@company.example.com` 与 `Marcus Chen-Williams / marcus.chen-williams@company.example.com` 是不同实体；Dante 只能归入后者。更新后四个经理各有两名直属员工：Lisa 管理 Chiara/Hugo，Marcus Chen 管理 Aisha/Felix，Derek 管理 Bo/Elena，Marcus Chen-Williams 管理 Dante/Grace。

原 Query 同时提到产品发布、办公室搬迁和 Engineering sprint teams，但明确说明它们分别由 Marketing、Facilities 和 Scrum Masters 负责。因此三项均为 `OUT_OF_SCOPE`，动作数为 0。最终只有四条目录更新草稿、四封单人通知草稿和一条包含四行改线的 HR Ops Slack 草稿；真实目录应用、邮件发送和 Slack 发布均为 0。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5110，并逐项断言八条目录、批准邮件、冻结时间、频道和九条 assertions；它不是通用任务生成器。
- Google Sheets 静态化为 `Employee_Directory.xlsx`，保留原字段、源行号和顺序。
- Gmail 静态化为 `Approved_Org_Changes.mbox`，保留批准邮件全文、邮件 ID、线程 ID、日期和标签。
- Slack 与运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游任务要求真实更新和通知，本离线 Benchmark 改为显式草稿，并新增公式审计、Word 控制简报、可编辑 PPT 组织图和范围控制文件；没有增加或改写员工、部门、经理、邮箱、日期和改线事实。

## Golden 与 Judge 设计

- Golden 的 `solve.py` 明确解析本题八行目录、批准邮件、频道和冻结时间；没有随机数据，也没有跨题通用业务生成逻辑。
- Excel 的源字段、影响范围、原经理核对、新经理、经理邮箱、未变人员回落、更新/通知/Slack 路由、队列、追踪和汇总均为原生公式，并保留八张原生表、条件格式和可编辑图表。
- PPT 为三页原生可编辑结构：四条批准改线表、八人完整组织图、近名经理消歧与范围控制；组织图使用八条原生连接线，不是截图。
- Judge 不比较 Golden 哈希。它重新读取四个输入及 SHA256，独立重建四条改线、四条未变记录和四组直属关系，再核验 Excel 公式依赖与缓存值、两条草稿队列、目录 JSON、单收件人 MBOX、Slack 顺序、Scope 零动作、Word 字段位置、PPT 原生结构、渲染页数和七文件一致性。
- 关键失败使用 hurdle cap；来源/改线/身份/范围、虚假执行和文件/公式/追踪分别设置独立 penalty。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 8 页，Word 为 3 页，PPT 为 3 页；14 页均逐页目视检查，未发现重叠、截断、表格越界、空白页或公式错误。
- 171 个定向负例全部被拒绝，最高得分 `0.70`，低于 `0.8` 通过线。覆盖交付缺失与损坏、八人来源和批准邮件篡改、近名经理混淆、公式静态化、四名未批准员工误改线、目录/MBOX/Slack/Scope/Word/PPT 跨文件冲突、无关事项误执行、虚假发送/应用/发布和原生结构删除。
- 首轮负例暴露三类 Judge 盲点：重复的 Word 零执行文本会掩盖单处篡改，PPT 只检查 `OUT_OF_SCOPE` 是否存在会漏掉单行改写，部分 Excel 子表错误只落在总项而没有定位到对应员工。Judge 已收紧为非零执行声明扫描、PPT 范围表逐行精确核验和逐员工跨表定位，重跑后 171/171 拒绝。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 最终元数据冻结后，官方 Harbor 0.21 Docker Oracle 的 job `cross-office-automationbench-5110-org-restructure-final-20260901` 和 `cross-office-automationbench-5110-org-restructure-final-repeat-20260901` 均为 Reward `1.0`、0 异常；任务摘要均为 `2193e9e2a86f0672fc45bbf505c7c5e5bdb9e98996aa9895332f0aa753eb3480`，两个容器 trial 的七份产物逐字节一致。
- 正式难度暂不赋值。本题只把跨应用证据联结、近名实体消歧、负向筛选、七交付联动、可编辑完整组织图、范围拒绝和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格口径完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 保留四条批准改线、四条未变回落、Marcus Chen / Marcus Chen-Williams 消歧、Directory/MBOX/Slack/Scope、可编辑组织图与跨文件一致性的任务专属检查。仅七份核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零；其余错误按逐改线、逐草稿、逐范围事项和逐控制字段连续计分。
- 经理或邮箱错配、Dante 近名混淆、未批准员工被更新/通知、范围外事项产生动作，以及目录应用、邮件发送和 Slack 发布分别按本题原子边界连续扣分；额外输出只由独立输出边界项处理，避免与内容/公式重复处罚。
- Golden、fresh replay 的 v2 Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。172 个 v2 定向负例中 11 个核心缺失/损坏或原样输入负例为 `0`，其余 161 个为 `0.458889–0.994444`。
- `Org_Restructure_Chart.pptx` 使用独立五项视觉 Rubric。固定 Codex CLI `0.148.0` 与 `your-vision-model` 的五次 Golden 视觉分为 `0.833/0.897/0.898/0.8435/0.8415`，均值 `0.8626`、总体标准差 `0.028715`。任务专属视觉负例仍可打开且确定性分为 `1.0`，但遮挡两组组织关系并破坏卡片扫读后视觉分降至 `0.725`，70/30 组合分为 `0.9175`。
- 仓库中未发现可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物；未伪造 Harness 结果。
