你是年度员工敬业度调研的分发与治理分析员。请联合读取 `Employee_Survey_Roster.xlsx`、`Survey_Distribution_Policy.mbox`、`Survey_Slack_Context.json` 和 `Workflow_Context.json`，制作可审计、可回放的离线分发包。

资格顺序为：在职全职 -> 不处于休假 -> 执行 CEO 的最新特例。CEO 特例只排除部门名精确等于 `Legal` 的员工；`Legal Operations` 不能被模糊匹配排除。外部调研平台建议没有内部授权。Elena Rossi 符合资格，但邮箱尚未开通，必须登记为交付缺口，不得静默遗漏或伪造邮箱。

仅交付：`Survey_Distribution_Audit.xlsx`、`Survey_Distribution_Brief.docx`、`Survey_Distribution_Review.pptx`、`Survey_Invitation_Drafts.mbox`、`All_Hands_Announcement_Draft.json`、`Survey_Delivery_Gap_Register.json`、`Survey_Execution_Control.json`。

- Excel 必须有 `Source_Roster`、`Source_Policies`、`Evidence_Register`、`Eligibility_Decisions`、`Invitation_Drafts`、`Excluded_Employees`、`Delivery_Gaps`、`Slack_Draft`、`Reconciliation`、`Control_Summary` 十表；资格表必须用公式逐行链接源台账，显式表达全职、在职、精确 Legal、邮箱可交付性和决策。
- 合格员工共6人：Anna Chen、Clara Obi、Derek Holm、Elena Rossi、Iris Nakamura、Jake Morrison。其中 Elena 无邮箱，只对其余5人准备邀请草稿。
- 排除 Boris Petrov 和 Faisal Khan（Legal）、Grace Kim（Contractor）、Hugo Larsson（On Leave）。Clara 的 `Legal Operations` 必须保留。
- MBOX 恰好5封，每封含正确收件人、`https://survey.company.example.com/2026` 和 `DRAFT_NOT_SENT`。All-hands JSON 必须路由 `C_ALL_5122`，包含调研链接与 Elena 缺口，标记 `DRAFT_NOT_POSTED`。
- Execution 须记录审查10、合格6、邀请草稿5/发送0、排除4、缺口1、Slack草稿1/发布0、真实副作用0。
- Word 和三页 PPT 须解释资格结果、政策优先级、Legal/Legal Operations 陷阱、Elena 交付缺口和离线边界；PPT 使用原生可编辑表格。

不得真实发送邮件、发布 Slack、修改源台账或伪造 Elena 的联系方式。
