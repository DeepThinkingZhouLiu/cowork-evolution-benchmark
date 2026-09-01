# C014 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=advanced` 是 L3 作者假设，不是最终难度或发布决定；`calibration_status=uncalibrated`。

## 来源、输入和许可

- 来源：OfficeBench，固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，任务 `tasks/2-14/subtasks/0.json`；原始 Query、日期/时间快照和文件存在 evaluator 均保留在 `upstream-subtask-0.json`。
- 输入：上游 `tasks/2-14/testbed/data/score.xlsx`，本地重命名为 `score_input.xlsx` 以区分输入和输出目录；SHA256 `a5c8026dc09f38411c6b8783dad74f7f66f08387eeb21b9d8eb7f28a90dd5b6e`，6,450 bytes。
- 许可：OfficeBench 仓库 Apache-2.0；仓库 LICENSE 保留在 `LICENSE-OfficeBench`。输入为公开 benchmark synthetic fixture，无个人敏感数据。

## 独立真值和改造边界

输入 `Sheet1` 的表头为 `Name / midterm1 / midterm2`，20 条记录中按 `midterm1` 降序、同分按原行稳定排序的前三名为 Tina=100、Rachel=100、Sam=99。上游 Query 提供日期 `2020-05-01`、星期 Friday 和 `10:00 AM`。

上游要求为前三名学生分别创建奖励会议事件。原环境会把事件写入每个用户的日历；本地任务将副作用投影为 `/app/output/Tina.ics`、`Rachel.ics`、`Sam.ics` 三个 UTF-8 iCalendar 文件。由于上游没有提供会议结束时间、地点、邮箱或时区，适配只要求 DTSTART、奖励会议标题、学生姓名和前三名语义，不加入这些未有证据的字段。每个文件使用一个原生 VEVENT，保留可导入结构。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 取前三名 | 输入成绩独立排序，稳定处理同分 | R002/R006 |
| 日期和时间 | 上游子任务 `date`/`time` | R003 |
| 奖励会议与学生对应 | 原始 Query、输出文件名和输入姓名 | R003/R004 |
| 原生事件交付 | 上游 calendar_create_event 的 VEVENT 语义和 iCalendar 结构 | R001/R005 |
| 仅目标文件 | 上游 evaluator 的三个文件存在断言 | R002/R005 |
| 输入可追溯 | 固定 commit 和 SHA256 | R006 |

Judge 不比较上游 Golden 或文件字节；它从输入工作簿独立重算排名，解析每个 VEVENT 的结构、DTSTART、标题/描述与文件集合。

## 参考解法和验证

参考解法 `solution/solve.py` 校验输入哈希和表头，读取全部记录后按 `(-midterm1, original_row)` 排序，生成三个 CRLF 规范的 iCalendar 文件。已运行：

```bash
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/xlsx/officebench-2-14-score-calendar
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/xlsx/officebench-2-14-score-calendar
```

Golden 与 fresh replay 均由 `tests/test.sh` 判为 `1.000000`；Judge 连续重复运行结果一致。`solution/Golden/score_source_evidence.xlsx` 仅是通用资产校验所需的隐藏输入证据副本，不是 solver 输出。负例覆盖空输出/缺文件/损坏 ICS、错误前三名、错误日期、缺标题或缺描述、额外学生文件；缺目标文件、错误排名或错误时间会触发核心门槛失败，其余语义缺陷保留部分分数。LibreOffice 可打开输入工作簿；ICS 用独立解析器检查，无需外部服务。

## 难度、泄漏和待审事项

- 作者假设：L3/advanced。任务包含一个成绩工作簿的稳定排序、跨应用日历交付和三个一致的事件文件；预计熟练办公人员约 25 分钟完成。文件数量本身不是难度依据，难点在排名到多交付的一致传递。
- leakage group：`officebench-score-calendar`。它与 OfficeBench `score.xlsx` 的其它任务共享输入 fixture，后续切分时应与同一工作簿变体隔离。
- 待人工决定：无结束时间的离线事件是否满足项目的日历交付标准、是否接受副作用投影、最终难度和是否纳入正式题库。Codex 不标记接受、校准或发布。
