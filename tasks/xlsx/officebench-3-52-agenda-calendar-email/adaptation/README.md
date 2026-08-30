# C022 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=advanced` 是 L3 作者假设，不是最终难度或发布决定；`calibration_status=uncalibrated`。

## 来源与输入

- 来源：OfficeBench 固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，任务 `tasks/3-52/subtasks/0.json`；原始 Query 和上游断言保留在 `upstream-subtask-0.json`。
- 输入：`tasks/3-52/testbed/data/meeting_agenda_2.xlsx`，本地命名为 `meeting_agenda.xlsx`；SHA256 `ffd3161364fde910fbd5491cc1dd19ba346134231e60f217ad668a874af80b0e`，11,199 bytes。
- 许可：OfficeBench 仓库 Apache-2.0，LICENSE 保留在 `adaptation/LICENSE-OfficeBench`；该工作簿为公开 synthetic benchmark fixture。

输入 `Meeting agenda` 表含三条议程：Introductions/Linda 13:00–13:30、Agenda/Davide 13:30–14:10、Strategy/Charlie 14:10–14:30。上游任务快照日期为 2020-05-01，Query 要求将议程加入 Bob 日历并分别给参与者发送提醒。

## 改造边界与独立真值

原任务的日历/邮件服务副作用改为 `/app/output/Bob.ics`、`Linda.eml`、`Davide.eml`、`Charlie.eml` 四个本地交付物。Bob 事件日期使用快照日期，起止时间由输入第一项 Start 和最后一项 End 计算；事件描述逐项保留输入 Item/Contact。每封邮件只包含对应联系人和对应议题，From 固定为离线草稿发送者 Bob，不新增收件人、邮箱事实或外部服务状态。

Judge 从输入工作簿独立重算议程顺序、时间和参与者，并解析 MIME/VEVENT，而非比较 Golden 字节或上游关键词。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 三条议程及联系人 | 输入工作簿行 3–5 | R003/R004 |
| 事件日期和起止时间 | 上游任务快照日期 + 输入首尾 Start/End | R002 |
| Bob 日历与三封独立提醒 | 原始 Query/evaluator | R001/R004/R005 |
| 原生交付格式 | iCalendar VEVENT、RFC MIME message | R001/R005 |
| 文件集合和输入追溯 | 上游输出意图 + SHA256 | R006 |

## 验证记录

已执行 `validate-task`、`validate-task-assets`；Golden/fresh replay 均由 `tests/test.sh` 判为 `1.000000`。`solution/Golden/agenda_source_evidence.xlsx` 仅作为通用公开资产校验所需的隐藏输入证据副本。负例覆盖空/损坏 ICS 或 EML、真正缺少 `Charlie.eml`、额外参与者、错误 DTSTART/DTEND、遗漏议题和错误收件人；缺失文件负例的 Judge 结果为 `0.500000`（R001/R004/R006 失败，hurdle 未通过），其余核心缺陷均低于通过阈值。输入工作簿可用 LibreOffice 打开，ICS/EML 由独立解析器检查。

## 难度、泄漏与待审

- 作者假设：L3/advanced。任务需要从单表抽取多行议程，构造一个跨时间范围的日历事件，并把行级信息一致地分发到三封邮件；预计熟练办公人员约 35 分钟完成。
- leakage group：`officebench-agenda-calendar-email`，与同一议程工作簿的其它变体隔离。
- 待人工决定：上游仅给出事件日期元数据而未给定时区，当前采用无时区本地时间；OwnCloud/邮件副作用的本地投影是否符合发布标准；最终难度及是否纳入正式题库。
