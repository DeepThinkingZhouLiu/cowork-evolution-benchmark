# officebench-1-16-random-paragraph v0.3 适配记录

状态：`local_validated`；difficulty 仍为 `unassigned`，未运行付费 Harness 难度校准。

## 适配契约与替换关系

- 角色与用途：文案编辑把一段咖啡馆简介整理成可追踪、可复核、可导航的 Word 审校工作稿。
- 权威事实源：OfficeBench 冻结的 `random_paragraph.txt`；最终文档不得引入人物、地址、价格、菜品、营业时间、评分或可量化承诺。
- 主模板：`cafe_profile_editorial_template.docx`，含受控原文、五个待修订句子、横竖分节、页眉页脚、台账和汇总表。
- 编辑票据：`cafe_profile_approved_copy_ticket.docx`，只增加五个经批准的事实保守改写、编辑类型、三个批注锚点与边界说明；不是新的业务事实源。
- 主交付物：`cafe_profile_editorial_review.docx`，不要求 PDF、PPTX、XLSX 或外部应用状态，仍属 DOCX 模块。
- 替换关系：v0.2 只把一个 TXT 复制到单个 `Normal` 段落，三个 Harness 均可满分。v0.3 保留“把冻结原文变成可编辑 Word 内容”的核心，但以真实文案审校流程整体 supersede 旧题，不是在“复制一段”后无关堆要求。

## 上游、许可与来源贡献

- Dataset：OfficeBench task `1-16/0`，仓库 `https://github.com/zlwang-cs/OfficeBench`。
- 冻结 commit：`b978b808667c32b52ce19a67ce1def1de9ae02b7`，Reference Code 工作树 HEAD 已复核为该值。
- 许可：Apache-2.0，完整副本保留在 `adaptation/LICENSE-OfficeBench`。
- 上游贡献：原 Query、evaluator 和 579 字节原文；原文是所有叙述事实的唯一权威。
- 适配自建贡献：两份 DOCX 输入的编辑流程、模板结构、批准改写和评注边界；`adaptation/build_inputs.py` 是本题专用的可复现构造器，不是跨题通用造题器。

## 逐文件主体已读证据

| 文件 / 文件组 | SHA256 / 结构 | 实际读取 | 关键发现 |
|---|---|---|---|
| 上游 `subtasks/0.json` / 冻结 `source-task.json` | `e2f86f4a...03b20` | 逐字段读取 Query、用户、时间、两个 evaluator | 只要求创建文件并命中首句短语，不验完整性或 Word 结构 |
| 上游 `random_paragraph.txt` / 两份打包副本 | `4439b9fb...85f55`，579 字节 | 全文逐句读取，核对 data/environment 字节一致 | 5 句、103 词；提供城市/咖啡馆、室内/受众、香气、顾客活动、内外节奏五组可追溯事实 |
| v0.2 `instruction.md` / `task.toml` / `rubric.json` / `judge.py` | 全文读取 | 核对输出名、7 项 rubric、Judge 字符串等值逻辑 | 核心只是一段原文与 SHA，没有分析、修订、批注、样式或保护内容 |
| v0.2 `solve.py` / `solve.sh` / Golden 两副本 | Golden `68d97565...21331`，1 段/0 表/1 节 | 读取解题分支和全部 OOXML 成员；LibreOffice 渲染 1 页并查看原图 | Golden 只是 Letter 纸上的一个默认段落，大量空白 |
| v0.2 `adaptation/*` / `metadata/*` / `tests/*` | 逐文件全文读取 | 核对许可、provenance、manifest、expectations 和 test 入口 | 来源冻结正确，但仍是 pending review 的低区分度快照 |
| 三个 v0.2 Harness 候选与收紧后 Judge 结果 | 各3个一段/0表，Reward 均1.0 | 读取 Claude Code、OpenCode、OpenHands 的 artifact 文本、criterion 和 reward | 三套 Harness 完成路径等价，证明旧题无区分度 |
| 上游 `README.md` / `LICENSE` | `04104d99...1d5ed2` / `c71d239d...d0ab4` | 全文读取 benchmark 目标、运行/评估方式与 Apache-2.0 条款 | 允许修改与再分发，需保留许可和修改说明 |
| 新 `cafe_profile_editorial_template.docx` | `f53dfaf5...a916b`，17 段/4 表/2 节 | 逐段、逐表、样式、页眉页脚和 OOXML 读取；渲染 2 页并逐页看图 | 第 1 页是受控源与 5 句工作文案，第 2 页是横向台账/汇总；无裁切或空白页 |
| 新 `cafe_profile_approved_copy_ticket.docx` | `7f728aac...066ed`，8 段/6×5 表 | 逐行读取批准文案、类型、锚点和评注；渲染 1 页并查看原图 | 五行全部在一页可读，三行有精确批注要求，两行明确不加批注 |

## 从输入独立重算的真值台账

| ID | 源句线索 | 源词数 | 批准词数 | 差值 | 事实保留边界 |
|---|---|---:|---:|---:|---|
| S1 | `In the heart of the bustling…` | 18 | 14 | -4 | 城市中心、繁忙环境、小咖啡馆与宁静休歇 |
| S2 | `Its cozy interior adorned with vintage…` | 21 | 15 | -6 | 舒适室内、复古家具、柔和灯光、当地人与游客 |
| S3 | `The aroma of freshly brewed coffee…` | 21 | 13 | -8 | 新鲜咖啡、烘焙点心两类香气，不增加菜品 |
| S4 | `Patrons sat at mismatched tables engrossed…` | 18 | 10 | -8 | 不匹配的桌子、交谈或沉思、饮用饮品 |
| S5 | `Outside the city buzzed with activity…` | 25 | 16 | -9 | 城市外部活动与室内时间放慢/短暂休息的对比 |

汇总真值是源文 103 词、批准文案 68 词、净变化 -35；这些值由 Judge 从 TXT 和票据 DOCX 独立重算，不读 Golden。

## Golden、Judge 与负例边界

- `solution/solve.py` 只解析本题的两份 DOCX 和一份 TXT；构造 5 个 `w:del`/`w:ins` 修订对、3 个原生 comments、5 个书签/内链、台账、汇总和 108 词证据说明，没有跨题分支或通用生成器。
- Judge 从冻结输入独立读取 5 个源句、5 个批准句、编辑类型、批注锚点与正文，逐修订对、逐批注、逐表格字段、逐书签和结构打分。
- 核心修订、受控源、批注、台账和跨区一致性均有业务合理 cap；次要导航可作为中等质量部分分。
- 本题专属负例需覆盖：缺失/损坏、受控源改字、接受修订或少一组修订、错批准句、错词数/差值、删批注、无根据价格扩写、图片压平、破坏横向分节，以及仅丢失书签内链的中等样例。

## 当前验收状态

- 旧题、所有上游/许可、全部旧输入/Golden/题面/Rubric/Judge/Solution/Adaptation/Metadata 和三个 Harness 候选已读入并记录。
- 旧 Golden 1 页、新模板 2 页、新票据 1 页均已使用 LibreOffice 7.3.7.2 渲染并逐页原图检查；输入结构清晰、无裁切、无空白页。
- 2026-09-07 使用项目专属 Python 3.11.16 环境执行 `validate-task`、`validate-task-assets`、JSON 解析、四个单题脚本 compileall、三份 data/environment 逐字节比较和 `git diff --check`，全部通过。
- Golden 与从空目录生成的 fresh replay 各重复评分 5 次，10 次 Reward 均为 `1.000000`；同一候选的 5 份 Judge JSON 逐字节一致。Golden SHA256 为 `cb98438fcb2373a7cd98ed2c3cc0b22be8e6130950b8cf2e69204d49fd643117`，fresh 为 `dc481c782b24b0d281cec28570a22bbbd361baf00503446f098edcd1ed70ad5b`；两个 ZIP 的 22/22 个解压成员逐字节一致，差异仅来自 ZIP 容器元数据。
- 缺失输出、损坏 DOCX 与额外交付物均为 `0.0`；定向负例为：受控源改字 `0.35`、接受所有修订 `0.58`、缺 S5 修订 `0.58`、S3 错批准文案 `0.58`、错词数/差值 `0.70`、删批注 `0.72`、无据价格与营业时间 `0.76`、图片压平 `0.1125`、破坏横向分节 `0.984444`；中等样例“内容完整但丢失导航”为 `0.93`。所有核心负例未达 0.8 通过线，结构破坏样例被明确扣分，中等样例严格介于 0 和 Golden 之间。
- LibreOffice 7.3.7.2 对 Golden 和 fresh 打开/另存后再评分，两者均为 `1.0`，另存后 20/20 个解压成员一致。Judge 只容忍 LibreOffice 为锚定批注把单个批准插入拆成多个同作者 `w:ins`、以及重排批注 ID；完整删除/插入文字、作者、锚点、批注正文和书签目标仍全部精确校验。
- Golden/fresh 各2页、LibreOffice 另存后各2页，以及 9 类相关负例共21页已逐页查看原图。第1页受控源和五组修订全部在界内，第2页台账、汇总、证据说明一页完整；无空白页、表格跨页、重叠、裁切、字体缺失或不可读内容。横向分节被破坏的负例右侧台账列可见被裁，截图作弊样例只剩单页黑色图块，可见失败与 Judge 因果一致。
- Harbor Docker Oracle A/B 待当前题目冻结后执行；证据只写入题目外 repo-level audit，避免回改被测 task digest。
- 作者侧页面检查不进 Reward；本题不增加 VLM 或 `visual_rubric`。

## 已知边界

- Word 和 LibreOffice 对修订显示颜色、批注气泡和字段缓存的视图可不同；Judge 只认 OOXML 修订/批注/字段和实际内容，不依赖界面颜色。
- 原始文段是 OfficeBench benchmark fixture，provenance 保持 `synthetic=true`，不冒充真实客户文案。
- 正式难度尚未校准；Oracle 成功只能证明任务可解，不代表 Agent 难度已验证。
