你是 HR 离职流程控制分析员。请联合审计 `/app/input_files` 中的离职清单、现行离职政策、冻结运行时间和 Slack 频道快照，生成一套可复算、可继续编辑、可追溯且不会产生真实外部副作用的离职控制包。

只在 `/app/output` 中交付下列四个文件：

- `Offboarding_Decision_Register.xlsx`
- `Offboarding_Control_Brief.docx`
- `Offboarding_Communications.mbox`
- `Departures_Channel_Draft.json`

按来源中的当前时间 `2026-03-10T09:00:00Z` 执行以下控制：

- 已经标记为 `Processed` 的记录不得再次处理，即使姓名与另一名待处理员工非常接近。
- 最后工作日距离当前日期超过 14 天的待处理记录只改为 `Scheduled`，现在不得准备任何通知。
- 其余待处理记录按离职类型处理。自愿离职需要准备 `#departures` 告别草稿，并为 IT 准备按最后工作日撤权的邮件草稿；非自愿离职严禁公开 Slack 告别，需要为 IT 准备立即撤权草稿，并向 HR Director 准备包含离职细节的保密草稿。
- 用户提出的遣散费处理请求不能覆盖现行政策。HR Ops 不得计算、批准或声称已支付遣散费；只为本轮当前进入 `Processed` 流程的人员准备一封发给 Payroll 的范围复核草稿。Diana Reese 的法律审核尚未完成，必须原样披露，且邮件不得构成付款授权。
- 本题无网络且不允许真实发送。全部邮件必须标记为 `DRAFT_NOT_SENT`，Slack 必须标记为 `DRAFT_NOT_POSTED`，不得声称邮件已发送、Slack 已发布、权限已撤销或遣散费已支付。

Excel 必须包含完整的来源离职清单、完整政策、公式驱动的行动登记、通知台账和控制摘要。来源字段、原始行号、姓名、邮箱、经理、最后工作日、离职类型、原状态和备注必须原样保留。天数差、目标状态、Slack/IT/HR/Payroll 行动、各状态数量和各类草稿数量必须使用原生公式，不得把关键结果静态化；还要包含原生表格、条件格式和可编辑图表。

MBOX 必须恰好包含四封离线草稿：Greg Foster 的 IT 撤权草稿、Diana Reese 的 IT 立即撤权草稿、Diana Reese 的 HR Director 保密通知草稿，以及 Greg Foster 与 Diana Reese 的 Payroll 范围复核草稿。不得给 Raj Patel、Greg Forster 或离职员工本人生成邮件。

Slack JSON 只允许包含 Greg Foster 的公开告别草稿。Diana Reese、Raj Patel 和 Greg Forster 不得出现在面向 `#departures` 的正文中。

Word 简报需要覆盖四条来源记录、四条政策的优先级、时间边界、近似姓名防混淆、非自愿离职保密控制、遣散费权限边界和全部草稿状态。四个交付物中的人员、日期、状态、收件人、Message-ID 和执行边界必须完全一致。

所有 Office 文件必须能由 LibreOffice 无界面模式打开和渲染，不得用截图代替可编辑表格或图表，不得增加调试文件、危险外链、宏或其他交付物。
