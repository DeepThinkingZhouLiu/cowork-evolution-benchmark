你是团队日历冲突审计员。请联合读取 `/app/input_files/Bob.ics` 和 `/app/input_files/Tom.ics`，以 UTC 时间完成事件对账、冲突计算、合并日历与管理报告。输入的 `DTSTART`/`DTEND` 都以 `Z` 结尾，不得自行换算本地时区。

在 `/app/output` 中交付且只交付以下三个文件：

1. `Calendar_Conflict_Audit.xlsx`，依次包含五个工作表：
   - `Source_Events`：按开始时间、Owner、Summary 排序，列出两份日历的全部事件，包含 Event_ID、Owner、Summary、DTSTART_UTC、DTEND_UTC、Duration_Minutes、UID、Source_File、Internal_Conflict_Count 和 Cross_Owner_Conflict_Count。Duration 与两个冲突计数必须使用可复制的原生公式。
   - `Pairwise_Conflicts`：列出所有 Bob 事件与 Tom 事件的严格正时长交叉，记录双方 Event_ID、Summary、交叉起止时间、Overlap_Minutes 和 Same_Title。后两项必须使用原生公式。
   - `Conflict_Windows`：把互相重叠或首尾相接的跨 Owner 冲突区间合并成最大连续窗口，列出起止时间、原生公式时长、活跃事件和包含的 pairwise 冲突数。
   - `Joint_Free_Windows`：在 `2024-05-01 08:00:00` 至 `20:00:00` UTC 之间，列出 Bob 和 Tom 同时无事件的全部正时长窗口，用原生公式计算时长、标记最长窗口并给出可追溯的建议。
   - `Summary`：全部指标使用原生公式汇总，至少包含事件数、各 Owner 事件数、内部冲突数、pairwise 冲突数、合并冲突窗口数与总分钟数、联合空闲窗口数与总分钟数、最长空闲窗口分钟数、同标题同时间事件数以及合并日历事件数。提供一个可编辑图表，对比冲突时长和联合空闲时长。
2. `Merged_Schedule.ics`：包含两份输入中的全部 10 个 `VEVENT`，保留每个事件的 UID、DTSTART、DTEND 和 SUMMARY，不新增、删除或改期；按开始时间、Owner、Summary 排序，并为每个事件增加 `X-EVOCOWORK-OWNER` 与 `X-EVOCOWORK-SOURCE`。
3. `Calendar_Conflict_Report.docx`，标题为 `Calendar Conflict Audit Report`，至少包含 `Scope and Time Semantics`、`Source Event Reconciliation`、`Conflict Findings`、`Joint Availability`、`Merged Calendar Control` 和 `Audit Conclusion` 六节。报告必须用原生表格列出全部事件、pairwise 冲突、合并冲突窗口和联合空闲窗口，并明确区分“单人日历内部冲突”和“两人之间同时被占用”。

冲突按半开区间 `[DTSTART, DTEND)` 计算：只有交叉时长严格大于 0 才算冲突，一个事件结束时另一个事件恰好开始不算冲突。Excel、ICS 和 Word 中的事件、时间、分钟数、冲突与结论必须完全一致；不得编造时区、参会人、地点、优先级或新会议。
