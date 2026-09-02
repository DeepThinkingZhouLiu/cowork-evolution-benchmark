# OfficeBench 3-91 日历冲突审计任务逐题适配记录

## 为什么选择这道题

上游 subtask 0 要求判断 Bob 和 Tom 的日历是否冲突，并在 `schedule.docx` 中记录冲突或合并日程。这是一条天然的“两份 ICS -> 时间区间审计 -> 可复算 Excel -> 合并 ICS -> Word 决策报告”跨 Office 链路。

结构设计目标为 L3 / Hard：Agent 需要保留 10 个来源事件，区分单人内部冲突、跨人事件对冲突和合并后的冲突窗口，再在指定时间范围内计算共同空闲，并保证 XLSX、ICS 和 DOCX 三份产物一致。按项目规则，正式难度仍保持 `unassigned`，等待冻结多 Harness Rollout 后统一校准。

## 实际读入上下文的上游材料

- `tasks/3-91/subtasks/0.json`：本题选中的完整 Query 和原 Evaluator。
- `tasks/3-91/testbed/calendar/Bob.ics`：逐事件读取的 5 个 Bob UTC 日程；本题 Agent 输入为其逐字节副本。
- `tasks/3-91/testbed/calendar/Tom.ics`：逐事件读取的 5 个 Tom UTC 日程；本题 Agent 输入为其逐字节副本。
- `utils/evaluate.py`：上游通用评测入口。
- OfficeBench 根目录 `README.md` 和 `LICENSE`，固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`。

为保留审计链，本目录保存了选中 subtask、原 Evaluator 和 Apache-2.0 许可证的精确快照。

## 逐事件人工真值

| 审计对象             | 人工核对结果                                                                                              |
|----------------------|-----------------------------------------------------------------------------------------------------------|
| 来源事件           | Bob 5 个，Tom 5 个；全部为 2024-05-01 UTC；UID、标题、开始和结束时间均保留                                 |
| 单人内部冲突       | Bob 0 对，Tom 0 对                                                                                       |
| 跨人事件对冲突     | 5 对：08:00-11:00、12:30-13:00、13:00-13:30、16:00-18:00、18:30-19:30 UTC                           |
| 合并冲突窗口       | 4 个：08:00-11:00、12:30-13:30、16:00-18:00、18:30-19:30 UTC；并集总时长 420 分钟                      |
| 共同空闲窗口       | 在 08:00-20:00 UTC 内共 4 个：11:50-12:00、14:30-16:00、18:00-18:30、19:30-20:00；总时长 160 分钟 |
| 最长共同空闲       | 14:30-16:00 UTC，90 分钟                                                                               |
| 同名同时间事件对   | 2 对：class 与 dinner                                                                                   |
| 时间语义             | 严格使用 UTC 和半开区间 `[start, end)`；首尾相接不算冲突，不擅自转换时区                            |

## 跨文件因果链

```text
Bob.ics + Tom.ics
-> 10 个冻结 UTC VEVENT
-> Excel 原生公式计算时长、冲突对、合并窗口和共同空闲
-> 保留来源 UID/时间/标题并补充 Owner 追溯字段的合并 ICS
-> 与 Excel 和 ICS 完全一致的 Word 审计报告
```

Judge 不读取 Golden 当真值，而是重新解析两份冻结 ICS，核对 SHA256，独立重算冲突对、冲突并集和共同空闲，再检查 Excel 公式与重算缓存、合并 ICS 事件、Word 事实、跨文件一致性和无来源断言。

## 对上游的必要改造

- 保留上游两份 ICS 和“检查冲突/合并日程”的业务意图，不新增会议、地点、与会人、优先级或时区。
- 上游 Evaluator 只检查 `schedule.docx` 存在且包含拼错的 `confliction`，无法发现事件改期、漏冲突、重复计时、合并 ICS 丢事件或跨文件矛盾。
- 本题人工重写 18 条 Rubric 和任务专属 Judge，并加入 19 类定向负例。
- 首次 Harbor 回放暴露出 Judge 仅从宿主机任务目录取输入的兼容问题；修正为同时支持 `/app/input_files` 后，两次官方 Docker Oracle 均通过。

## 许可、隐私与泄漏

来源为 Apache-2.0 的 OfficeBench 公开任务。按上游 benchmark 属性保守标记 `synthetic=true`。任何复用这两份日历或相同事件真值的变体都必须归入 `officebench-3-91-bob-tom-calendars` leakage group，不得跨 split。

## 验证结果

- 单题协议、输入清单与 SHA256 已通过严格校验。
- Golden 与两次空目录 fresh replay 的 Judge Reward 都是 `1.0`，两次本地三份产物与 Golden 逐字节一致。
- 19 个定向负例全部低于 `0.8` 通过线，最高错误得分为 `0.67`。
- XLSX 经 LibreOffice 渲染为 5 页，DOCX 渲染为 3 页，共 8 页已逐页目视检查；时间列无 `###` 或截断。
- 官方 Harbor 0.21.0 Docker Oracle 独立回放两次，均为 Reward `1.0`、0 异常；两次容器产物逐字节一致。
