# AutomationBench 5106 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5106()`，`example_id=5106`，任务名 `hr.compliance_training_enforcement`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`cf8ba527624a40066b3e8472581f7b4ddaf92cb040b90023b74a32cf7737e33e`。

本题制作时完整读取了：原 Query、`initial_state.meta`、完整合规手册邮件、Google Sheets 中八名员工、Slack 频道、十三条上游 assertions、仓库 LICENSE。没有执行 AutomationBench 的外部服务或写操作。

## 上游事实审计

手册 Section 6 的优先级是：来源状态 `Complete` -> 有效豁免 -> 逾期分段。委员会只接收超过 30 天的案例；1–30 天案例不得抄送或通知委员会。逐案真值如下：

| 源行 | 员工             | 截止日     | 冻结日逾期 | 覆盖条件             | 最终动作                 | 路由                                       |
|-----:|------------------|------------|-----------:|----------------------|--------------------------|--------------------------------------------|
|    2 | Patricia Chen    | 2026-04-01 |         14 | 无                   | `WARNING`                | 本人，CC Vikram Shah                       |
|    3 | Patricia Chang   | 2026-03-10 |         36 | 无                   | `COMMITTEE_ESCALATION`   | 合规委员会                                 |
|    4 | Roberto Alvarez  | 2026-03-01 |         45 | 豁免有效至 4 月 30 日 | `EXEMPT`                 | 不通知                                     |
|    5 | Samantha Green   | 2026-03-25 |         21 | 无                   | `SUSPENSION_REQUEST`     | IT Security 请求 + Diana Torres 告知       |
|    6 | Theodore Kim     | 2026-02-10 |         64 | 无                   | `COMMITTEE_ESCALATION`   | 合规委员会，不直接警告或请求 IT            |
|    7 | Ursula Martin    | 2026-04-10 |          5 | 无                   | `WARNING`                | 本人，CC Diana Torres                       |
|    8 | Victor Pham      | 2025-12-15 |        121 | 无                   | `COMMITTEE_ESCALATION`   | 合规委员会，不错误套用 15–30 天规则        |
|    9 | Wendy Torres     | 2026-04-14 |          1 | 来源状态为 Complete   | `NO_ACTION`              | 不通知                                     |

最终分布为 2 个警告、1 个暂停请求、3 个委员会升级、1 个豁免和 1 个无动作。五封邮件均为离线草稿；访问暂停只形成一项请求草稿，实际执行记录为 0。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5106，并逐项断言八名员工、完整手册、冻结时间、频道和十三条 assertions；它不是通用任务生成器。
- Google Sheets 状态静态化为 `Training_Compliance_Tracker.xlsx`，保留全部原字段、原字符串、源行号和顺序。
- Gmail 状态静态化为 `Compliance_Handbook.mbox`，保留手册全文、邮件 ID、线程 ID、日期和标签。
- Slack 与运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 没有新增员工、日期、培训、状态、豁免、收件人或政策。新增的公式审计簿、Word 简报和草稿/控制文件只用于把已审计真值变成离线、可复算、可确定检查的 Office 交付包。

## Golden 与 Judge 设计

- Golden 的 `solve.py` 明确解析本题八行状态、冻结日期、完整手册和豁免截止日；没有随机数据或跨题通用业务生成逻辑。
- Excel 的来源、逾期天数、完成覆盖、有效豁免、动作、路由、通知队列和汇总均保留原生公式；LibreOffice 重算后再冻结包结构。
- Judge 不比较 Golden 哈希。它重新读取四个输入和来源 SHA256，独立构造八案真值，再检查公式、缓存值、原生表格、条件格式、图表、MBOX 收件人/抄送/正文、Slack 聚合隐私、访问控制、Word 同行事实和跨文件一致性。
- 关键失败使用 hurdle cap；来源/政策/日期/豁免/动作、路由/隐私、虚假发送/执行和文件/公式/追踪分别设置独立 penalty。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 6 页，Word 在宿主机和 Harbor Debian 容器中均渲染为 3 页；9 页均逐页目视检查，未发现文字重叠、截断、表格越界或公式错误。
- 初版 Word 在宿主机为 3 页、容器为 4 页；原因是容器字体度量使案例表最后一行溢出。案例表改用短角色标签，完整邮箱仍保留在路由表，修复后两边均稳定为 3 页。
- 74 个定向负例全部被拒绝，最高得分 `0.70`。首轮负例暴露 Word 中正确数字和政策词可被其他位置掩盖，Judge 随后收紧为“员工、截止日、天数、动作必须在同一案例行”以及四条政策逐行精确核验。
- 两次宿主机空目录 fresh replay 均为 `1.0`，五份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 的 job `cross-office-automationbench-5106-compliance-enforcement-fixed-20260901` 和 `cross-office-automationbench-5106-compliance-enforcement-repeat-20260901` 均为 Reward `1.0`、0 异常；容器内五份产物逐字节一致。
- 正式难度暂不赋值。该题只把长政策抽取、日期分段、状态/豁免优先级、相似姓名隔离、五路草稿、聚合隐私和访问执行边界作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。

## v2 连续评分改造（2026-09-04）

- 原严格口径完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 逐项保留本题八名员工、Section 6 四条政策、完成/豁免优先级、五封通知草稿、Slack 聚合隐私和单一暂停请求的任务专属检查。仅五份核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零；其余事实、公式、结构和跨文件错误保留连续部分分。
- 虚假邮件发送、Slack 发布和访问暂停按本题九个执行原子边界连续扣分；Slack 逐案泄露、禁止直接收件人、委员会越界及多余暂停对象按十四个隐私/路由原子边界连续扣分，不使用旧版统一封顶。
- Golden、fresh replay 的 v2 Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。75 个 v2 定向负例中 3 个致命负例为 `0`，其余 72 个为 `0.355000–0.995000`；虚假访问执行、虚假发布和禁止路由显著低于单字段、追踪或公式局部错误。
- 本题只交付 XLSX、DOCX、MBOX 和 JSON，不含 PPT/PDF，因此不调用 VLM。仓库中未发现可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物。
