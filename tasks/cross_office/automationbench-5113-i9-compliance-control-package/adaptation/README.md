# AutomationBench 5113 逐题改造记录

## 上游定位

- 数据集：Zapier AutomationBench。
- 精确版本：`4a8e1061254004d9dac807054eed33fad7d1ff14`。
- 原任务：`automationbench/domains/hr/tasks.py` 中 `get_task_5113()`，`example_id=5113`，任务名 `hr.i9_verification_tracking`。
- 许可证：MIT；本目录保留精确上游 `LICENSE` 副本。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`。
- 冻结单题快照 SHA256：`0475b34d45ff90266f5bf4bea0e5cd9935277e3c5ec68408f30081b31e51dd5`。

制作时完整读取了原 Query、冻结时间、八条 recent-hire 记录、Legal 就业验证政策、Slack `#hr-general` 频道快照和十条上游 assertions。没有连接或执行 AutomationBench 外部服务，也没有真实发送邮件、修改 I-9 台账或发布 Slack。

## 上游事实审计

Legal 政策要求：雇主应在员工开始日起 3 个工作日内完成 Section 2；窗口内未完成者向员工发送主题为 `I-9 Reminder` 的提醒并 Cc 经理，建议状态为 `Reminder Sent`；超过窗口者通知 `legal@company.example.com` 并建议状态为 `Violation - Legal Notified`。本题没有节假日输入，因此确定性地只按周一至周五计算，经过日数从开始日下一日起算至冻结日 `2026-04-15`，不超过 3 日仍在窗口内。

| 源行 | 员工            | 开始日期   | 源状态     | 经过工作日 | 决策                                  |
|-----:|-----------------|------------|------------|-----------:|---------------------------------------|
|    2 | Amara Obi       | 2026-04-14 | Incomplete |          1 | 提醒员工并 Cc 经理                    |
|    3 | Boris Petrov    | 2026-04-07 | Incomplete |          6 | 法务违规通知                          |
|    4 | Cassie Lin      | 2026-04-10 | Complete   |          3 | 无动作                                |
|    5 | Dmitri Volkov   | 2026-04-01 | Incomplete |         10 | 法务违规通知                          |
|    6 | Esme Nakamura   | 2026-04-11 | Incomplete |          3 | 提醒员工并 Cc 经理                    |
|    7 | Frankie Reeves  | 2026-03-25 | Incomplete |         15 | 法务违规通知                          |
|    8 | Gina Torres     | 2026-04-15 | Incomplete |          0 | 提醒员工并 Cc 经理                    |
|    9 | Hank Morrison   | 2026-04-08 | Incomplete |          5 | 法务违规通知                          |

上游 assertions 没有点名 Hank Morrison，但完整台账与政策明确表明他已超过窗口，所以必须进入法务违规路由。Cassie Lin 的源状态为 `Complete`，日期计算不能覆盖该权威状态，她不得进入邮件或台账更新草稿。最终是八人审查、三人提醒、四人违规、一人无动作，七封邮件草稿、七条未应用更新、零 Slack 内容和零真实副作用。

## 静态化与改造边界

- `materialize_inputs.py` 只接受 5113，并逐项断言八条台账、政策全文、冻结时间、Slack 频道和上游 assertions；它不是跨题通用造题器。
- Google Sheets 静态化为 `I9_Verification_Tracker.xlsx`，保留八行源顺序、员工/经理邮箱、开始日期、I-9 状态和 Notes。
- Gmail 静态化为 `Employment_Verification_Policy.mbox`，保留 Legal 政策全文与元数据。
- Slack 和运行时间分别静态化为 JSON；`data/input_files` 与 `environment/input_files` 逐字节一致。
- 上游任务原本要求实际发送邮件和更新台账，本离线 Benchmark 只生成明确未发送的 MBOX 草稿、未应用的 JSON 更新、公式审计簿、Word/PPT 复核和执行控制记录；没有新增员工、邮箱、经理、日期、状态或政策事实。

## Golden 与 Judge 设计

- Golden `solve.py` 明确解析本题四个冻结输入，独立计算八人的截止日、经过工作日和三类路由，并在题内断言完整真值；没有随机数据或跨题通用业务生成逻辑。
- Excel 包含九张原生表。开始日、第 3 个工作日截止日、经过工作日、完成/窗口判断、提醒/法务/无动作路由、建议状态、队列和十三项汇总都由原生公式派生，并保留三类条件格式和可编辑图表。
- Word 为三页、五张原生表；PPT 为三页，每页一张可编辑原生表，覆盖八人决策、三人提醒、四人违规、Hank 补全、Cassie 保护、工作日口径和零真实执行。
- Judge 不比较 Golden 哈希。它重新读取四个输入与 SHA256，独立重建八人真值，再核验公式依赖与缓存、三类 Excel 队列、七封 MBOX、七条 Tracker 更新、Slack/Execution JSON、Word/PPT 原生结构、渲染页数和七文件一致性。
- 关键失败使用 hurdle cap；来源/日期/状态/工作日错误，Hank/Cassie/邮件路由错误，Slack 或虚假执行，以及文件/公式/原生结构/追踪/跨文件错误分别设置独立 penalty。

## 视觉、负例与可回放性

- Golden Excel 经 LibreOffice 渲染为 9 页，Word 为 3 页，PPT 为 3 页；15 页均逐页目视检查，未发现重叠、截断、表格越界、空白页或公式错误。
- 254 个定向负例全部被拒绝，最高得分 `0.05`，低于 `0.80` 通过线。覆盖文件缺失/损坏、八人逐字段事实、政策/时间/频道证据、工作日公式静态化、三人提醒、四人法务路由、Hank 漏项、Cassie 误处理、错误 To/Cc/Bcc/顺序/正文、台账应用、Slack 泄露、虚假执行、跨文件冲突和原生结构删除。
- 负例暴露并修复了三类 Judge 盲点：MBOX 反转顺序只触发个人事实但未触发队列标准；Word 某处声称 `real side effects=1` 可能被另一处正确摘要掩盖；PPT 某名员工的草稿路由只记在表现标准而未回写其个人标准。最终全量重跑为 `254/254` 拒绝。
- Golden 正例为 `1.0`，空交付为 `0.0`；两次宿主机空目录 fresh replay 均为 `1.0`，七份产物分别与 Golden 逐字节一致。
- 官方 Harbor 0.21 Docker Oracle 的 job `cross-office-automationbench-5113-i9-compliance-final-20260902` 与 `cross-office-automationbench-5113-i9-compliance-repeat-20260902` 均为 Reward `1.0`、0 异常；两次 trial 的任务摘要均为 `0989be8b9647246c7a7fba9fc080c5eeb235fb8c22268d6556382074733c08b4`，七份容器产物逐字节一致。该摘要对应回放时的完整任务快照；本行回放记录与 `metadata/validation.json` 是 trial 结束后写入的作者侧证据，不反向声称自身已包含在该摘要中。
- 正式难度暂不赋值。本题把多源合规推导、工作日边界、上游漏项补全、完成状态保护、七交付联动和零副作用作为 L4 / Expert 的结构设计目标，最终标签等待冻结多 Harness Rollout。
