# OfficeBench 3-6 团队培训协调包逐题适配记录

## 为什么选择这道题

上游任务要求从 Excel 人员表安排下周培训，生成日历并给所有成员分别发送提醒。这是一条天然的“冻结名单 -> 会话审计 -> 日历 -> 逐人消息 -> 交付控制”跨 Office 链路，而不是把几个无关文件放在一起。

结构设计目标为 L3 / Hard：Agent 需要保留 12 个源行、建立 3 个会话、生成 12 封独立 multipart MIME 草稿和 12 个与 Master ICS 同源的日历附件，同时不得伪造邮箱、时区和地点。正式难度保持 `unassigned`，等待冻结多 Harness Rollout 后统一校准。

## 实际读入上下文的上游材料

- `tasks/3-6/subtasks/0.json`：完整 Query、任务日期和原 Evaluator。
- `tasks/3-6/testbed/data/team_training_schedule.xlsx`：逐行读取的 12 名成员、Section 和时间；Agent 输入为其逐字节副本。
- `utils/evaluate.py`：上游通用评测入口。
- OfficeBench 根目录 `README.md` 和 `LICENSE`，固定 commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`。

本题保存了上游 subtask、原 Evaluator 和 Apache-2.0 许可证的精确快照。

## 逐行人工真值

| 审计对象       | 人工核对结果                                                                 |
|----------------|------------------------------------------------------------------------------|
| 源行           | 表头 1 行 + 12 个人员行；源行号为 2-13                                    |
| Section A      | Alice、David、Grace、Jack；09:00-11:00；4 人                         |
| Section B      | Bob、Eva、Hank、Kelly；11:00-13:00；4 人                            |
| Section C      | Carol、Frank、Ivy、Leo；14:00-16:00；4 人                           |
| 培训日期       | 2020-05-08；由任务日期 2020-05-01 的“next week”与上游 Evaluator 共同固定 |
| 时间关系       | A/B 在 11:00 首尾相接，不重叠；B/C 之间 60 分钟；每组 2 小时                |
| 已知交付数     | 3 个 Master VEVENT；12 封独立 MIME 草稿；每封恰好 1 个 Section 日历附件    |
| 证据缺口       | 真实邮箱、时区、地点、组织者/与会人邮箱都未提供                         |
| 安全状态       | 所有邮件使用 `.invalid` 占位地址并标记 `BLOCKED_MISSING_REAL_EMAIL`；不宣称已发送 |

## 跨文件因果链

```text
team_training_schedule.xlsx
-> 12 个不可丢失的源行
-> Excel 原生公式计算人数、时长、间隔和阻塞数
-> 3 个稳定 UID 的 Master VEVENT
-> 12 封逐人 MIME 草稿 + 与 Master ICS 同真值的日历附件
-> 与 Excel/ICS/MBOX 一致的 Word 协调简报
```

Judge 不读取 Golden 当真值：它会校验源表 SHA256，重新读取 12 个源行，独立重算三个会话和间隔，再检查 Excel 公式缓存与原生图表、Master ICS、MBOX 的 12 个 MIME 邮件及其日历附件、Word 事实、无来源断言和四产物一致性。

## 对上游的必要改造

- 上游 Evaluator 仅用关键词检查一个 ICS 和各人邮件，无法发现漏人、换组、改时间、重复收件人、日历附件错配或伪造邮箱/时区/地点。
- 本题人工重写 18 条 Rubric 和任务专属 Judge，并加入 22 类定向负例。
- 首轮负例回放真实发现 Judge 对 `TZID` 参数和 Word “已发送”变体措辞的两个漏检；修复后两类错误分别降为 0 和 0.15。

## 许可、隐私与泄漏

来源为 Apache-2.0 的 OfficeBench 公开任务。按上游 benchmark fixture 属性保守标记 `synthetic=true`。任何复用该 12 人培训表或相同分组真值的变体都必须归入 `officebench-3-6-team-training-schedule` leakage group，不得跨 split。

## 验证结果

- Golden 与两次空目录 fresh replay 的 Judge Reward 均为 `1.0`；四份产物均与 Golden 逐字节一致。
- 空交付 Reward 为 `0.0`；22 个定向负例全部低于 `0.8` 通过线，最高错误得分为 `0.20`。
- XLSX 经 LibreOffice 渲染为 4 页，DOCX 渲染为 2 页，共 6 页已逐页目视检查；已修正表头断行、列边界、邮箱断行和第二页页首遮挡。
- 官方 Harbor 0.21.0 Docker Oracle 独立回放两次，均为 Reward `1.0`、0 异常；两次容器四份产物逐字节一致。
