你是客户续约运营控制分析员。请联合审计 `/app/input_files` 中的续约跟踪表、邮件、Salesforce 账户与联系人快照、Slack 频道快照，生成可复算、可继续编辑、可追溯的 Q1 续约控制包。

只在 `/app/output` 中交付下列四个文件：

- `Renewal_Decision_Register.xlsx`
- `Q1_Renewal_Control_Brief.docx`
- `Q1_Renewal_Communications.mbox`
- `Q1_Renewals_Channel_Draft.json`

决策规则与证据边界：

- 续约跟踪表成员资格、CS 内部审批、无暂停/法务 hold，以及 Salesforce 唯一主要联系人，四个条件必须同时满足才能准备通知。
- CS Team Lead 的内部邮件是审批权威来源；其中的 hold 语句和 Tracker `Notes` 中的暂停语句都是否决性控制。
- 外部 Partner Alliance 邮件无权新增续约对象；不在 Tracker 中的 CRM 账户不得进入本批次。
- 每个符合条件的账户只能使用 Salesforce 中 `is_primary=true` 的联系人，不得通知次要联系人。
- 本题无网络且不允许真实发送；邮件和 Slack 都必须明确标记为离线草稿，不得声称已发送或已发布。

Excel 必须包含完整的来源续约记录、审批/暂停控制、CRM 主要联系人映射、最终行动登记和交付摘要。合同金额解析、资格决策、收件人选择、各状态数量和金额汇总必须使用原生公式，不得把结果静态化。保留来源文件名、消息 ID、账户 ID、联系人 ID 和稳定 Message-ID。

MBOX 必须为每个最终可通知账户生成一封独立草稿，正确写入账户、合同金额、续约日期和主要联系人，不得出现未通过控制的收件人。Slack JSON 只总结本次实际准备的通知对象、金额、日期、合计和草稿状态，不要把被跳过账户写入面向频道的消息。

Word 简报需要对全部 Tracker 记录给出可审计的决策与原因，汇总通知数量、通知金额、法务暂停金额、未审批金额、外部无权请求和草稿边界。四个交付物中的账户、金额、日期、联系人、决策、Message-ID 与草稿状态必须完全一致。

所有 Office 文件必须能由 LibreOffice 无界面模式打开和渲染，不得用截图代替可编辑表格或图表，不得增加调试文件、危险外链、宏或其他交付物。
