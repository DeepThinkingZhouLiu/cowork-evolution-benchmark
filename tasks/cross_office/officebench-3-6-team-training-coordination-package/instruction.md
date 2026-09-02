你是团队培训协调员。请以 `team_training_schedule.xlsx` 为唯一人员与分组依据，把 OfficeBench 任务中“2020-05-01 下一周”的培训安排整理成可审计、可继续编辑的跨 Office 交付包。

只在 `/app/output` 中交付下列四个文件：

- `Team_Training_Coordination_Audit.xlsx`
- `Team_Training_Coordination_Brief.docx`
- `Team_Training_Master.ics`
- `Team_Training_Reminders.mbox`

必须满足以下要求：

- 完整保留来源表的 12 名成员、所属 Section 和原始时间文本，不得合并同名、换组、改时间或漏人。
- 培训日期为 2020-05-08。Section A 为 09:00-11:00，Section B 为 11:00-13:00，Section C 为 14:00-16:00。这些是本题上游 Query 和评测中的已知真值。
- Excel 必须是可复算的审计工作簿：包含来源人员、会话安排、逐人提醒控制和交付摘要；成员数、时长、间隔、消息数与阻塞数等关键统计必须使用原生公式，不得把计算结果静态化。
- Master ICS 必须恰好包含 3 个可编辑 VEVENT，与三个 Section 一一对应，时间使用未绑定时区的 floating local time。不得伪造时区、地点、组织者邮箱或参会人邮箱。
- MBOX 必须包含恰好 12 封独立的 multipart MIME 提醒草稿，每人一封。每封邮件的姓名、Section、日期和时间必须与 Excel/ICS 一致，并附加该 Section 的 `text/calendar` 邀请。
- 来源表没有任何真实邮箱、地点或时区。因此邮件必须使用 `.invalid` 保留地址，并明确标记为“缺少真实邮箱，不可发送”。不得声称邮件已发送。
- Word 需要清楚总结 12 人、3 个 Section、各组时间、Section B 与 C 之间 60 分钟间隔、全部 12 封草稿的阻塞状态，以及未提供的邮箱/地点/时区边界。
- 四个文件中的人数、分组、时间、日期、UID、收件人和交付状态必须一致。不得增加调试文件、危险外链或其他交付物。

所有 Office 文件都必须能由 LibreOffice 无界面模式打开和渲染，不得用截图替代可编辑表格或图表。
