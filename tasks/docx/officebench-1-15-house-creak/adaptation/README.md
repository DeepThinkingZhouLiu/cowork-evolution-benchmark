# officebench-1-15-house-creak v0.3 适配记录

状态：`local_validated`；difficulty 依然为 `unassigned`，未执行付费 Harness 难度校准。

## 适配契约

- 角色与受众：短篇选集文字编辑，为主编与校对员制作可导航、可批注、可追溯的 Word 工作稿。
- Agent 可见输入：`House_creak.docx`，同时承担需保护的故事原稿和附录中概要、时序、角色分析、批注的唯一事实来源。
- 主交付物：`house_creak_editorial_review.docx`，只有 DOCX，不要求 PDF/XLSX/PPTX 或外部应用状态，因此仍属 DOCX 模块。
- 保留核心：OfficeBench 原题的“编辑既有 Word 文档并保护非目标内容”能力仍然保留，但 v0.3 不再把追加一句话当作完整任务。
- 新的真实工作流：保护原稿 -> 从原稿进行内容综合 -> 用 Word 原生分节、标题、TOC、表格、书签、内链和批注形成编辑工作稿。
- 不可妥协约束：47 个原文现有段落（含标题与空段）的文字/顺序/基本格式和第一节页面几何保留；分析事实不得超出原文；三个批注必须锚定到精确原文；六个证据链接必须返回原稿。
- 设计目标：复杂 Word 编辑、证据综合与内容保护候选；正式难度只能由后续冻结 Panel 确定。

## 上游与许可

- Dataset：OfficeBench task `1-15/0`。
- 来源仓库：`https://github.com/zlwang-cs/OfficeBench`。
- 冻结 commit：`b978b808667c32b52ce19a67ce1def1de9ae02b7`，已在 `003-Reference/002-Reference-Code/sources.lock.json` 复核。
- 许可：Apache-2.0，完整副本保留为 `adaptation/LICENSE-OfficeBench`。
- 上游原 Query 与唯一 evaluator 已冻结为 `adaptation/source-task.json`。原 evaluator 只检查末尾句子关键词，不能支撑内容保护、故事分析、Word 对象和证据链评分，所以只作迁移审计证据，不作为 v0.3 Judge 的真值容器。
- 上游标注为 benchmark fixture；当前 provenance 保持 `synthetic=true`、`human_rewritten=true`、`publishability=public`，不冒充真实客户文档。

## 逐文件已读证据

| 文件 | 字节 | SHA256 | 实际检查 | 关键发现 |
|---|---:|---|---|---|
| 上游 `tasks/1-15/subtasks/0.json` / 冻结 `adaptation/source-task.json` | 454 | `54eec8e5bfd089322a7f26470efa557ef5d7761e67ea85c0d60865f76473fcb7` | 逐字段读取 Query、用户、时间和 1 条 evaluator | 原题只追加固定句，evaluator 只做 `evaluate_contain` |
| 上游/打包 `House_creak.docx` | 23,193 | `e549d600d201cefcf67cc4e42cf04ff1b2828d567689cb36a29af8e6e7655fd7` | `python-docx` 逐段逐 run、ZIP/OOXML/关系展开；LibreOffice 7.3.7.2 导出 4 页 PDF；4 页 120 DPI PNG 全部原尺寸查看 | 47 段、0 表、0 图片、1 节；Century Schoolbook 正文，带页码；无宏/OLE/外链；故事完整提供风险线索和人物行动事实 |
| `environment/input_files/House_creak.docx` | 23,193 | `e549d600d201cefcf67cc4e42cf04ff1b2828d567689cb36a29af8e6e7655fd7` | 与公开输入逐字节比较 | 与 `data/input_files` 完全一致，Docker 复制到 `/app/input_files` |
| v0.2 Golden `solution/Golden/house_creak.docx` | 20,540 | 历史文件，已由 v0.3 Golden 取代 | 逐段、OOXML 与 4 页渲染检查 | 只在第 48 段追加测试句，证实旧题缺乏区分度 |
| `adaptation/LICENSE-OfficeBench` | 11,357 | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4` | 完整读取许可条款 | Apache-2.0 允许修改和再分发，需保留许可和修改说明 |
| 上游 `README.md` | 6,969 | `04104d99107c2f109ea57d06f75541051efd4e7b5a8b2e41083315b3d21d5ed2` | 读取 benchmark 目标、任务格式和 evaluator 说明 | OfficeBench 定位为办公自动化 benchmark；当前适配保留来源标识但替换过弱交互 |

## 从原文重算的真值台账

| fact_id | 原文位置（0 起始段序） | 可核对原文 | 对附录的贡献 |
|---|---:|---|---|
| `F-MOVE` | P10 | `The house ... was Victorian Gothic` | 时序第 1 行；搬入偏远房子的概要事实 |
| `F-CREAK` | P19 | `The sound was worst in the master bedroom and children's rooms next to it` | 时序第 2 行；结构异响证据 |
| `F-POWER` | P19 | `power point ... over 1000 watts ...` | 时序第 3 行；电气风险批注锚点 |
| `F-BED` | P31 | `a double bed crash through the dining room ceiling` | 时序第 4 行；结构失效批注锚点 |
| `F-GAS` | P36 | `cake baking in the gas oven` + `gas main ... fractured` + `huge 'boom!'` | 时序第 5 行；爆炸因果链 |
| `F-CONDEMNED` | P42 | `That house was condemned years ago!` | 时序第 6 行；最终风险确认批注锚点 |
| `F-RELIEF` | P43 | 邻居联系紧急服务；Women's Institute 提供衣物、食物和狗粮，并安排到外婆家 | Synopsis 结局和“Neighbour and Women's Institute”角色行 |
| `F-MOTHER` | P16, P29, P44 | Mother 尝试查房屋来历、命令大家外逃、后来拒绝父亲继续单方决策 | Mother 角色行的证据、影响与编辑解读 |
| `F-ROS` | P19, P25–P29, P32 | Ros 倾听异响、急促吠叫并奔向大门 | Ros 角色行与概要中的预警功能 |
| `F-CHILDREN` | P10–P12, P23–P35, P45–P46 | 孩子被迫隔离、参与临时课堂、逃生后住外婆家并返校 | Peter and Amanda 角色行 |
| `F-FATHER` | P4–P6, P10–P13, P17, P44 | 父亲为声望单方决定搬家，忽视家人现实处境，房屋未投保 | Father 角色行的决策与后果 |

所有 Golden 分析句都由上表证据综合，没有新增房价、维修、保险、人员伤亡、警方处置等原文不存在的故事事实。

## Golden 与可编辑性

- `solution/solve.py` 只实现本题的原文锚点、六行时序、五行角色台账和编辑表达，不是跨题生成器。
- Golden 在 47 个原文段落后添加下一页分节，附录含 1 个 TOC 字段、4 个二级标题、2 张原生表格、6 个书签、6 个内部链接和 3 个精确锚定批注。
- 两张表的表头带 `w:tblHeader`，数据行带 `w:cantSplit`，文本、链接、批注、书签和表格都为原生 Word 对象。

## Judge 和负例设计

- Judge 从 `/app/input_files/House_creak.docx` 独立读取 47 个原段落、run 格式、第一节页面几何和页码，不读 Golden 哈希或 solve 标记。
- 时序表按 6 行的原文证据、风险含义与后果逐行评分；角色表按 5 个实体逐行评分。
- 批注通过 `comments.xml` 与 `commentRangeStart/End` 联合检查锚定文本；导航通过书签与 `w:hyperlink/@w:anchor` 联合检查。
- 本题负例覆盖：缺文件、损坏 DOCX、原文改字、错误“房屋通过检验”核心时序、缺角色行、新增无来源事件、截图式作弊、删批注，以及作为中等样例的“主要内容完成但内链丢失”。

## 当前验收状态

- 完整上下文审计与 v0.3 造题：已完成。
- 2026-09-07 使用项目专属 Python 3.11.16 环境执行 `validate-task` 与 `validate-task-assets`，两者均通过；`compileall` 和 `git diff --check` 通过。
- Golden 与从空输出目录独立生成的 fresh replay 各重复 5 次，10 次 Reward 均为 `1.000000`，且两份 DOCX 逐字节一致（SHA256 `700dd565a63f576ec5b6155f366f08dcbc4efadd1b731203e6fde5c974f20121`）。
- 空交付与损坏 DOCX 均为 `0.0`；定向负例 Reward 为：原文改字 `0.35`、错误核心时序 `0.72`、缺角色行 `0.72`、无来源额外事件 `0.72`、截图式作弊 `0.0`、删批注 `0.72`；中等样例“主要内容完成但丢失六个内链”为 `0.94`。所有定向负例都未达 `0.8` 通过线，中等样例严格介于 0 与 Golden 之间。
- LibreOffice 7.3.7.2 打开/另存 Golden 和 fresh 后再次 Judge，两者仍为 `1.0`。Golden 和 fresh 各渲染 6 页，12/12 页逐页原图查看：第 1–4 页故事无改字、裁切或空白页，页码正常；第 5 页概要与 6 行时序表完整；第 6 页 5 行角色表和编辑注完整，表格未跨页断裂，无重叠、越界、字体缺失或不可读内容。
- 相关负例也用同一 LibreOffice 渲染并逐页原图查看：中等样例 6/6 页、错核心时序 6/6 页、原文改字 6/6 页、截图式作弊 1/1 页。前三份保持可打开且错误内容可见；截图作弊只有一页标题与黑色图块，与 `0.0` 判定一致。
- Harbor Docker Oracle A/B：待当前题目冻结后执行；结果放题目外 repo-level audit，避免使已测 task digest 失效。

## 已知边界

- 作者侧视觉检查使用固定 LibreOffice/PDF/PNG 渲染，不进入 Reward，本题不引入 VLM。
- 目录为可更新字段；不同 Word 引擎是否立即刷新显示文字存在差异，Judge 检查字段指令与标题层级，不依赖某一引擎的缓存页码。
- 正式难度尚未校准；Oracle 成功只证明题目可解，不能当作 Agent 难度证据。
