你是高校奖学金合规协调员。请联合阅读 `/app/input_files/school_policy.jpg` 和 `/app/input_files/fellowship_application.pdf`，逐条审计申请是否满足政策，并完成决策、转发草稿和复核会议产物。

在 `/app/output` 中交付且只交付以下四个文件：

1. `Fellowship_Compliance_Audit.xlsx`，依次包含四个工作表：
   - `Policy_Criteria`：逐字转录政策图中六条规则，并记录 Criterion_ID、Policy_Text、Application_Evidence、Status、Determinative 和 Notes。`Status` 只能为 `PASS`、`FAIL` 或 `NOT_VERIFIABLE`。
   - `Application_Evidence`：记录申请人姓名、邮箱、邮箱格式有效性、Statement of Purpose 词数、参考人数与逐条参考人、政策截止日期、提交时间和证据来源。英文词数按连续英文字母组计算，句读不单独计词。
   - `Decision_Summary`：用原生公式统计 PASS、FAIL 和 NOT_VERIFIABLE 条数，并用原生公式给出 Overall_Decision。任一可判定政策为 FAIL 时结论为 `NO`；无 FAIL 但存在不可核验的必要条件时为 `NOT_VERIFIABLE`；其余为 `YES`。
   - `Delivery_Control`：列出四份交付物的文件名、用途、与决策的一致性要求以及交付状态。
   
   工作簿必须使用原生表格、筛选、冻结窗格、条件格式和一张可编辑的政策状态图；不得把公式或图表拍平。
2. `Fellowship_Compliance_Decision.docx`，标题为 `Fellowship Application Compliance Decision`，至少包含 `Decision`、`Applicant Evidence`、`Policy-by-Policy Assessment`、`Evidence Gaps`、`Required Remediation` 和 `Delivery Controls` 六节。报告必须用原生表格展示申请证据与六条政策判定，并和 Excel 结论完全一致。
3. `Fellowship_Application_Forward.eml`，作为给 Fellowship Officer 的 MIME 邮件草稿，主题为 `Fellowship application compliance review - Alice Johnson`，附加原始 `fellowship_application.pdf`，附件必须与输入逐字节一致。输入没有提供经办人邮箱，因此 `To` 必须使用明确的不可投递占位地址 `Fellowship Officer <address-unavailable@example.invalid>`，并写入 `X-EvoCowork-Delivery-Status: BLOCKED_MISSING_OFFICER_ADDRESS`。邮件正文必须给出同一合规结论、关键原因和证据缺口，不得把占位地址写成真实经办人联系方式。
4. `Fellowship_Compliance_Followup.ics`，包含且仅包含一个 `VEVENT`，开始时间为 `2024-05-31 13:00:00 UTC`，主题为 `Fellowship compliance follow-up with officer`，状态为 `TENTATIVE`。上游没有提供会议时长、地点或经办人邮箱，因此不得添加 `DTEND`、`DURATION`、`LOCATION` 或 `ATTENDEE`；在 Description 中写明合规结论、需复核的缺失材料与截止日期证据缺口。

政策图未提供具体截止日期，申请 PDF 也未提供可验证的提交时间。这一项只能标为 `NOT_VERIFIABLE`，不得推断“已按时提交”或“已过期”。不得编造第二位参考人、截止日期、提交时间、经办人真实邮箱、会议时长、地点或最终获奖资格。
