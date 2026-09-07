# officeval-003-coordination-minutes-audit v0.3 适配记录

状态：`local_validated_harbor_pending`；difficulty 仍为 `unassigned`，未运行付费 Harness 难度校准。

## 替换关系与模块边界

- 迁移映射：`cowork-evo/officebench-1-16-test-sentence` -> `cowork-evo/officeval-003-coordination-minutes-audit`。
- 旧题只有一句 `This is a test file for task 1-16-1.`；可见 TXT 与请求无关，三个 Sonnet 5 Harness 各用一次简单 `python-docx` 写入即满分。这个 Seed 无法自然升级，故整体 supersede，而不是在测试句后堆要求。
- 新任务采用公开 Apache-2.0 的 OmegaUse-OfficeVal `officeval_003`，保留其“会议内容逐行整理为细黑横线表”的 Word 核心。
- 新增复杂度只来自同一源记录：修复顶层表套内嵌表、删除空行、校对 5 条偏差、抽取 8 项决议、保留表决顺序、标记未明确字段、制作修订/批注/书签/内链及横向审计附录。
- 最终只交付一个可编辑 DOCX；PNG 是样式参考，不要求 PDF/PPTX/XLSX 或跨应用状态，仍属于 DOCX 模块。图片不会进入 Judge reward。

## 上游、许可与来源贡献

- Dataset revision：`cd6ba6d8fb83b3fb551e24eebc20e1fb0bd154a5`，由 Hugging Face 本地 metadata 逐文件复核。
- 代码仓 commit：`ffbeecb8752447c8e40b594a0eeb1db7236ecb36`。
- 许可：Apache-2.0；完整 `LICENSE`、`NOTICE`、Dataset README、双语 Query/Rubric 和 1,108 行上游 evaluator 均冻结在 `adaptation/`。
- 源贡献：353,660 字节 DOCX 提供全部会议事实与原始 Word 结构；451,593 字节 PNG 只提供横线表版式示意。
- 适配贡献：任务说明、源事实重算规则、输出结构、题目专属解题器、确定性 Judge 和负例；没有添加任务自建业务素材。

## 逐文件主体已读证据

| 文件 / 文件组 | SHA256 / 结构 | 实际读取 | 关键发现 |
|---|---|---|---|
| 旧 OfficeBench `subtasks/1.json` / 冻结 legacy source | `2651e80f...90d38` | 逐字段读取 Query、User、Time Limit、file-exist/contain evaluator | 只要求建 `test.docx` 并写一句固定测试句；输入 `random_paragraph.txt` 与请求无关 |
| 旧题 instruction/task/rubric/judge/solution/adaptation/metadata | 逐文件全文 | 读取全部文本、原 Solver 分支和输入；解析 Golden 全部 OOXML | 旧 Golden 为 1 个 Normal 段、0 表、1 节、17 个 OOXML 成员，SHA256 `8c89c00d...a0894` |
| 三个旧 Harness artifact、日志与收紧后 Judge | `4eb7c4cf...8cefb` / `70026ccd...7bd05` / `b3c1f464...92108` | 读取 Claude Code 2.1.251、OpenCode 1.4.14、OpenHands SDK 1.44.1 产物和结果 | 三者均 1 段/0 表，收紧后 Reward 仍为 1.0，结构完全等价 |
| OfficeVal 双语 task / rubric | `ead07295...4a86f` / `f6c03276...fd07` / `61fab037...f124` / `56d1f63d...e4d8d` | 全文逐项读取任务、劳动时间、来源文件和全部评分点 | 原任务由真实行政事务需求适配，人工时间 80 分钟；核心是细黑横线表、无空行、保护首页字段和可编辑性 |
| OfficeVal 上游 evaluator | `b870ac1d...2eb19`，1,108 行 | 全文读取递归表检查、边框、行高、可编辑性、首页字段、缩进、入口和结果结构 | 上游会递归处理嵌套表，但完成度仍主要集中在横线表版式，未核对 5 条偏差和 8 项决议 |
| OfficeVal Dataset README / LICENSE / NOTICE | `d6499f6a...97e3` / `8b5a791d...f1375` / `6d5a662f...18c9` | README 343 行、代码仓 README 227 行、Apache-2.0 201 行全文读取 | 上游说明任务经隐私保护重构、人工复核和确定性 evaluator 验收；许可允许修改和再分发并要求保留声明 |
| 新源 DOCX | `3c0d0060...97344`，12 段/1 顶层表/1 内嵌表/1 节/30 ZIP 项 | 逐段、逐 run、逐行、逐表、样式、边框、页眉页脚、分节和全部 ZIP 成员读取 | 顶层表 108 行，内嵌表 56 行，共 181 个 XML 段落、122 个非空段；59 个空段/行和多处硬断行；事实含 5 条偏差、8 项决议 |
| 新样式参考 PNG | `9da14688...4dc7` | 原始分辨率查看 | 明确要求一行一条、下方全宽细黑实线；没有业务事实 |
| 新源 DOCX LibreOffice 渲染 | 8 页 A4 | LibreOffice 7.3.7.2 导出 PDF，8 页全部原图查看 | 源文可读但表中表、长行、空白横线和跨页布局混杂；第 8 页只余散会句并有大量空白 |

## 真值与因果边界

- 会议日期 `2026年3月18日`；主持人许承柏、记录人林若棠、地点综合楼二层会议室，应/实出席数均 5，列席人员陈予宁、侯景澜、魏预、周言朔。
- “二月份作业偏差共发生若干件”后实际连续列出 5 条，故只允许以未接受修订把“若干”改为“5”，并用批注说明计数来源。
- 5 条偏差的日期、时段、检查场景与发现、责任人和“一般偏差1件”定性逐项见 `evidence-ledger.md`；只有 2 月 12 日一条没有明确时段，写“未记录”。
- 8 项后续事项均有 4 条正式“同意”记录，但表决人员顺序并不完全相同，必须逐议题抽取；所有责任人和完成时限在源记录中均未明确。
- Judge 从冻结源 DOCX 递归读取顶层表和内嵌表，合并排版断行、拆分拥挤在单元格内的偏差条目，再独立重算表格真值；不导入 Solver、不读取 Golden 哈希、不接受候选自报成功。

## Golden、Judge 与负例验收

- `solution/solve.py` 只处理本题两份输入：递归读取源 DOCX、生成 161 行单列表、1 个修订对、1 个批注、13 个书签/内链、5×2 摘要、6×6 偏差表和 9×8 决议表，没有跨题分支。
- Judge 逐字段检查源文保护、全文顺序、修订/批注、一列表格每行宽度/边框/行高/不拆行、两张审计表、13 个回链、两分节、页眉页脚字段和安全部件。
- 本题专属负例覆盖：缺失/损坏/额外交付物、首页事实破坏、正文漏行、接受修订、计数错误、批注删除、嵌套/空行/错误边框、偏差责任人错配、表决顺序错配、虚构责任人/时限、书签断链、图片压平、破坏横向分节，以及仅丢失导航的中等样例。

## 当前状态

- `validate-task`、`validate-task-assets`、compileall、JSON 解析、两份输入的 data/environment 逐字节比对和 `git diff --check` 均通过。
- Golden 与 fresh replay 各5次都为 1.0；两份 DOCX 的27/27个解压业务成员逐字节一致，仅 ZIP 时间戳不同。
- 缺失、损坏、额外交付物和 13 类任务定向负例均按因果边失败；只移除13个书签/内链的中等样例为0.93。
- LibreOffice 7.3.7.2 对 Golden/fresh 打开另存后均为1.0，两份另存件的33/33个解压业务成员一致。Golden、fresh、两份另存件各 8 页均已逐页实看，无裁切、重叠、乱码、空白页或表格拆行。
- 负例中红色粗线、图片压平、破坏横向分节和重新引入嵌套表的视觉因果已实看；竖向附录负例的右侧列明显被裁切。
- 不增加 VLM 或 `visual_rubric`；上述视觉检查只是作者侧可用性验收，不进入 Reward。下一步是对冻结快照做 Harbor Docker Oracle A/B 独立双回放。
