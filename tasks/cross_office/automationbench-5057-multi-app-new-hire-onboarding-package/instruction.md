你是新员工跨应用入职编排与对账分析员。请联合读取 `Full_Onboarding_Checklist.xlsx`、`Initial_Application_State.json` 和 `Workflow_Context.json`，为本周两位新员工分别准备 BambooHR 档案、Employee Documents 子文件夹、Jira 入职工单和 Gmail 欢迎邮件，并制作可审计、可回放的离线执行包。

仅交付：`New_Hire_Onboarding_Audit.xlsx`、`New_Hire_Onboarding_Brief.docx`、`New_Hire_Onboarding_Review.pptx`、`BambooHR_Profile_Create_Drafts.json`、`Drive_Folder_Create_Drafts.json`、`Jira_Onboarding_Ticket_Drafts.json`、`New_Hire_Welcome_Drafts.mbox`、`Onboarding_Execution_Control.json`。

- Excel 必须有 `Source_Checklist`、`Source_System_State`、`Evidence_Register`、`Hire_Decisions`、`BambooHR_Drafts`、`Drive_Drafts`、`Jira_Drafts`、`Welcome_Drafts`、`Manager_Exclusion`、`Cross_System_Reconciliation`、`Readiness_Matrix`、`Control_Summary` 十二表；决策表须用公式链接源清单，验证四系统一一匹配和经理排除。
- Alicia Fernandez：Engineering / Frontend Developer / David Okonkwo / 2026-03-24 / Austin / `alicia.fernandez@company.example.com`。Tyrone Washington：Sales / Account Executive / Lisa Wang / 2026-03-25 / Remote / `tyrone.washington@company.example.com`。所有系统字段须逐字保真。
- BambooHR、Drive、Jira 各恰好2条草稿；Drive 父目录必须是 `folder_employees` / `Employee Documents`；Jira Summary 必须包含对应新人姓名，Description 包含完整入职字段。
- 欢迎 MBOX 恰好2封，只发给两位新人，包含姓名、部门、角色、日期、地点、经理和 `DRAFT_NOT_SENT`。David Okonkwo 和 Lisa Wang 只是经理字段，不得获得新人档案或欢迎邮件。
- Execution 须记录新人2，四类草稿各2，Bamboo/Drive/Jira 创建均0，邮件发送0，经理动作0，外部动作0，真实副作用0。
- Word 和三页 PPT 须解释两人四系统覆盖、身份一一对账、经理边界和离线执行；PPT 使用原生可编辑表格。

不得真实创建 BambooHR 档案、Drive 文件夹、Jira 工单或发送邮件。
