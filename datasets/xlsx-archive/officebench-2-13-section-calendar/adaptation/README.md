# C013 逐题改造记录

状态：`pending_human_review`。这是 pilot 的 L3 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench，commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，任务 `tasks/2-13/subtasks/1.json`；同任务的全体成员版本保留为 `upstream-subtask-0.json`，Section A 版本保留为 `upstream-subtask-1.json`。
- 原始 Query：`Schedule a team training session for members in section A from excel file, create calendar events file for each member in their name`。
- 原 evaluator 要求 `calendar/Alice.ics`、`David.ics`、`Grace.ics`、`Jack.ics` 存在，并要求 `calendar/Bob.ics` 不存在。
- 输入：`team_training_schedule.xlsx`，实际固定来源文件名；SHA256 `a5628f45d9f48eccd8424634fda4d3c9048200571cc03db1e6d34cf9447498f4`，5,164 bytes。
- 许可：上游仓库 Apache-2.0；来源 LICENSE 和任务记录由作者侧保留。

## 改造边界

输入工作簿的 `Training Schedule` 表含 `Name`、`Section`、`Time` 三列。独立读取结果显示 Section A 成员按行顺序为 Alice、David、Grace、Jack，四行时间均为 `9:00 AM - 11:00 AM`。上游子任务提供日期 `2020-05-01`，因此每个事件使用该日期和时间范围。

原任务的 `calendar/` 外部工作区路径和日历 MCP 副作用不适合固定离线 Harbor 包；这里投影为输出根目录中的四个 `.ics` 文件。文件名保留成员姓名，VEVENT 的 SUMMARY/DESCRIPTION 明确写入团队培训、成员和 Section A。没有引入邮箱、地点或其他输入未提供的事实，也没有复制上游 Golden 作为唯一真值。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 仅 Section A 四位成员各一个事件 | 输入表筛选、上游 evaluator | R002 |
| 日期和时间 | 上游子任务 `date`/`time` 与输入行 | R003 |
| 成员可识别且为团队培训 | 原始 Query 与输入 Section 字段 | R004 |
| 原生可导入日历交付物 | iCalendar RFC 结构要求 | R001/R005 |
| 输入来源可追溯 | 固定文件 SHA256/provenance | R006 |

## 验证记录

已执行：`validate-task`、`validate-task-assets` 通过；Golden 与 fresh replay 的 Judge 均为 `1.000000`。空输出/缺少成员均为 `0.333333`，额外 Bob 和错误时间均为 `0.833333` 但核心门槛失败；四个 `.ics` 的 VCALENDAR/VEVENT 结构独立解析通过。`solution/Golden/calendar_source_evidence.xlsx` 用 LibreOffice 7.3.7.2 打开并导出 PDF，文本包含 Training Schedule、Alice 和 Section A。当前 Judge 从输入工作簿重算 Section A，不读取 Golden。

## 风险与人工决策

- 上游 evaluator 只检查文件存在/不存在，不检查事件字段；日期来自上游子任务元数据而非工作簿，因此需人工确认该离线投影是否符合预期。
- 元数据使用 `language=bilingual`：任务说明为中文，iCalendar 事件内容保留原始英文语境。
- 同源培训排程题归入候选清单指定的 leakage group `officebench-class-calendar`。
- Codex/OpenCode 难度 rollout 尚未执行；`difficulty_tier=advanced` 为未校准的 L3 结构假设。
