你是 I-9 就业验证合规分析员。当前时间由 `Workflow_Context.json` 冻结。请读取 `I9_Verification_Tracker.xlsx`、`Employment_Verification_Policy.mbox` 和 `Slack_Channels.json`，依照 Legal 的完整就业验证政策，对八名 recent hires 完成离线合规治理包。

为使“3 个工作日”可确定回放，本题只把周一至周五视为工作日，不假定输入未提供的节假日。经过工作日数从开始日的下一日起计算，到冻结当日为止；不超过 3 日仍在窗口内，超过 3 日则为合规违规。已完成 I-9 的人员无论日期均不采取动作。

请在 `/app/output` 中只交付以下七个文件：

1. `I9_Compliance_Audit.xlsx`
2. `I9_Compliance_Control_Brief.docx`
3. `I9_Compliance_Review.pptx`
4. `I9_Notification_Drafts.mbox`
5. `I9_Tracker_Update_Drafts.json`
6. `I9_Slack_Control.json`
7. `I9_Execution_Control.json`

具体要求：

- Excel 至少包含 `Source_Hires`、`Policy_Source`、`Policy_Rules`、`Evidence_Register`、`Compliance_Decisions`、`Reminder_Queue`、`Legal_Violation_Queue`、`No_Action_Queue` 和 `Control_Summary` 九张工作表。来源行、日期推导、决策、队列和汇总必须可追踪。
- `Compliance_Decisions` 必须用原生公式逐行派生：开始日、第 3 个工作日截止日、经过工作日数、是否完成、是否在窗口内、提醒/法务/无动作路由、建议状态与来源追踪。
- 提醒队列必须按源顺序恰好包含 Amara Obi、Esme Nakamura 和 Gina Torres；每封邮件主题为 `I-9 Reminder`，收件人是员工，Cc 恰好是其源记录中的 Manager Email。
- 违规队列必须按源顺序恰好包含 Boris Petrov、Dmitri Volkov、Frankie Reeves 和 Hank Morrison；每封法务邮件收件人恰好是 `legal@company.example.com`，主题为 `I-9 Compliance Violation - <Employee>`，正文必须含员工姓名与开始日期。Hank 虽未被上游 assertions 点名，仍必须根据完整名单和政策正确处理。
- Cassie Lin 的源状态为 `Complete`，必须进入无动作队列，不得出现于任何提醒、法务通知或更新草稿中。
- MBOX 必须恰好含 7 封 `DRAFT_NOT_SENT` 邮件：3 封员工提醒与 4 封法务违规草稿；不得有 Bcc，不得包含 Cassie，不得声称已发送。
- `I9_Tracker_Update_Drafts.json` 只记录 7 条未应用的更新草稿：三人的建议状态为 `Reminder Sent`，四人为 `Violation - Legal Notified`；所有 `applied` 均为 `false`，Cassie 不得出现。
- Slack 不在 Legal 政策授权的 I-9 通知渠道中。`I9_Slack_Control.json` 必须记录零草稿、零发布、零员工/状态泄露，不得把 I-9 信息复制到公开的 `#hr-general`。
- `I9_Execution_Control.json` 必须明确：邮件草稿 7、实际发送 0；台账更新草稿 7、实际应用 0；Slack 草稿与发布 0；真实副作用 0。
- Word 与 PPT 必须准确呈现八人完整决策、三人提醒、四人违规、Cassie 无动作、Hank 补全推导、政策时间窗口和零真实执行。PPT 固定三页，关键表格与图形保持原生可编辑，不得用截图替代。
- Office 文件必须能由 LibreOffice 正常打开和渲染；不得包含宏、外部链接、嵌入式可执行对象或未提供来源支持的人员/日期/状态事实。

所有外部动作都必须停留在离线草稿或控制记录；不得真实发送邮件、更新台账或发布 Slack。
