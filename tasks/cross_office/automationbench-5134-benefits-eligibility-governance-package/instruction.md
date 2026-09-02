你是福利登记资格治理分析员。请联合读取 `Benefits_Enrollment_Queue.xlsx`、`Benefits_Policies_Roster_and_Request.mbox`、`Benefits_Slack.json` 和 `Workflow_Context.json`，制作可审计、可回放的离线资格审查包。

资格必须同时满足：入职日不晚于 2026-01-15、参加最近一次 orientation、已签表单。任一缺失即未满足。Legal hold 覆盖其他结果：不登记、不通知。外部 broker 要求加入的 Jordan Keene 不在员工名单中，必须拒绝。

仅交付：`Benefits_Eligibility_Audit.xlsx`、`Benefits_Control_Brief.docx`、`Benefits_Review.pptx`、`Benefits_Eligible_Confirmation_Drafts.mbox`、`Benefits_Not_Eligible_Drafts.mbox`、`Benefits_Slack_Draft.json`、`Benefits_Execution_Control.json`。

- Excel 必须有 `Source_Queue`、`Source_Messages`、`Policy_Rules`、`Evidence_Register`、`Eligibility_Decisions`、`Eligible`、`Not_Yet_Eligible`、`Legal_Hold`、`Enrollment_Update_Drafts`、`Employee_Comms`、`External_Slack_Control`、`Control_Summary` 十二表；决策表须用公式推导三项条件、暂停、缺失原因、通信和零执行。
- Alex Rivera、Farida Hassan 合格；Brianna Cho 缺少 90 天年资；Carlos Mendez 和 Graham Liu 缺 orientation；Darius Idowu 缺已签表单；Elena Popov 为 legal hold。最终 2 合格、4 未满足、1 暂停、1 外部请求拒绝。
- Eligible MBOX 恰好 2 封，Not Eligible MBOX 恰好 4 封并按源序，每封都有全名、精确结果/缺口和 `DRAFT_NOT_SENT`。Elena 和 Jordan 不得出现为收件人。
- 仅为 Alex 和 Farida 准备 `Enrolled` 台账更新草稿，实际应用0。Slack 为一条包含七人处置、Jordan 拒绝和零执行的草稿，发布0。
- Execution 须记录 7/2/4/1/1，员工草稿6、邮件发送0、更新草稿2、实际更新0、Slack发布0、外部请求未采纳和真实副作用0。
- Word 和三页 PPT 须解释资格合取、法律暂停覆盖、外部请求拒绝和离线执行边界；PPT 使用原生可编辑表格。Office 文件须由 LibreOffice 正常渲染，无宏、外链、危险对象或无来源事实。

不得真实登记、修改源台账、发邮件或发 Slack。
