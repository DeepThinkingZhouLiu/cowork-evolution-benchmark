# officebench-1-15-syllabus-integrity v0.3 适配记录

状态：`local_validated`；difficulty 依然为 `unassigned`，未执行付费 Harness 难度校准。

## 适配契约

- 角色与受众：课程协调员为学院政策复核会议制作一份可追踪、可导航、可审计的课程大纲政策整合工作稿。
- Agent 可见输入：`sample_syllabus.docx` 是必须保护的主文档；`academic_integrity.docx` 是必须原样纳入的研究诚信政策来源。两份文件共同构成唯一事实基础。
- 主交付物：`sample_syllabus_policy_review.docx`，只有 DOCX，不要求 PDF、PPTX、XLSX 或外部应用状态，因此仍属 DOCX 模块。
- 保留核心：OfficeBench 原题的“把研究诚信来源纳入现有课程大纲”能力完整保留；v0.3 不再接受无定位、无修订痕迹、无冲突解释的文末粘贴。
- 新的真实工作流：保护既有模板 -> 识别 Turnitin 与 Special Needs 的语义边界 -> 以修订插入完整政策 -> 连接一般诚信、研究设计和校外实践三类范围 -> 用批注、书签、内链与控制矩阵形成复核工作稿。
- 不可妥协约束：187 个现有段落、34×4 课程表、标识图像、外链和首节几何保持；7 个非空政策段落按源顺序且只插入一次；所有新增事实必须可回溯到两份输入。
- 设计目标：复杂 Word 政策整合、冲突消歧、修订/批注和模板保护候选；正式难度只能由后续冻结 Panel 确定。

## 上游与许可

- Dataset：OfficeBench task `1-15/1`。
- 来源仓库：`https://github.com/zlwang-cs/OfficeBench`。
- 冻结 commit：`b978b808667c32b52ce19a67ce1def1de9ae02b7`，已在 `003-Reference/002-Reference-Code/sources.lock.json` 复核。
- 许可：Apache-2.0，完整副本保留为 `adaptation/LICENSE-OfficeBench`。
- 上游原 Query 与 evaluator 已冻结为 `adaptation/source-task.json`。原 evaluator 只检查 `ACADEMIC INTEGRITY FOR RESEARCH`，不能验证放置位置、完整性、模板保护、修订痕迹或范围一致性，所以只作迁移审计证据。
- 上游标注为 benchmark fixture；当前 provenance 保持 `synthetic=true`、`human_rewritten=true`、`publishability=public`，不冒充真实学生或真实客户数据。

## 逐文件已读证据

| 文件 | 字节 | SHA256 | 实际检查 | 关键发现 |
|---|---:|---|---|---|
| 上游 `tasks/1-15/subtasks/1.json` / 冻结 `adaptation/source-task.json` | 542 | `9e5504efbf2f9e3e245f753fd92481cb94f23d9e97aeff6c8988fcdd413c5ae7` | 逐字段读取 Query、用户、时间和唯一 evaluator | 原题只要求文末追加；evaluator 只做关键词包含检查 |
| 上游/打包 `sample_syllabus.docx` | 188,874 | `9efefdcac29942e59edf9537aa5cfade576ea55b408dde34205792fd773c8505` | `python-docx` 逐段逐 run、34×4 表格、图片/关系/样式和 OOXML 展开；LibreOffice 7.3.7.2 导出 8 页并逐页原尺寸查看 | 187 段、1 表、1 标识图；既有一般诚信、Turnitin、Special Needs 和校外实践条款已构成需要消歧的完整语境 |
| 上游/打包 `academic_integrity.docx` | 16,079 | `260bd988b35c3a3577238d9f15f6a15010a39e5de3a167cc73b65bbaad3291cc` | 逐段逐 run、列表属性与 OOXML 展开；LibreOffice 导出 1 页并逐页查看 | 9 段、7 个非空段；定义研究数据范围，并给出目的、选择、风险收益、匿名四项原则 |
| `environment/input_files/*.docx` | 同上 | 同上 | 分别与 `data/input_files` 逐字节比较 | Docker 固定复制到 `/app/input_files`，无宿主路径或运行时联网依赖 |
| v0.2 Golden `solution/Golden/sample_syllabus_edited.docx` | 160,301 | `1562b68f8f5c7df2525d7397443c0a408b4367f6efc403fb51b2807134a0f462` | 逐段、表格、OOXML 与 8 页渲染检查 | 只把来源文本无修订地追加至最后一页，证实旧题缺乏区分度 |
| `adaptation/LICENSE-OfficeBench` | 11,357 | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4` | 完整读取许可条款 | Apache-2.0 允许修改和再分发，需保留许可和修改说明 |
| 上游 `README.md` | 6,969 | `04104d99107c2f109ea57d06f75541051efd4e7b5a8b2e41083315b3d21d5ed2` | 读取 benchmark 目标、任务格式和 evaluator 说明 | OfficeBench 定位为办公自动化 benchmark；本适配保留来源标识但替换过弱交互 |

## 从输入重算的真值台账

| fact_id | 输入位置（0 起始段序） | 可核对短语 / 事实 | 对工作稿的贡献 |
|---|---:|---|---|
| `F-GENERAL` | syllabus P92–P101，重点 P96 | `fabrication, falsification or sabotage of research data` | “一般研究数据不当行为”控制行；第一条批注锚点 |
| `F-TURNITIN` | syllabus P101 | `verified using Turnitin technology` | 新政策的前置定位锚点；第二条批注解释文本原创性与研究行为规范的互补关系 |
| `F-SPECIAL` | syllabus P104 | `Special Needs` | 新政策必须位于此前，防止破坏后续章节边界 |
| `F-FIELD` | syllabus P184–P185 | `apply to your interactions with the staff ... community organizations and its clients` | “校外/社区适用范围”控制行；第三条批注锚点 |
| `F-DATA-SCOPE` | integrity P1 | 学生、教师、同事；口头、书面、观察、记录、作品均是研究数据 | 研究数据范围控制行和 Scope Decision |
| `F-PURPOSE` | integrity P4 | 明确研究目的或问题；收集的数据必须与之相关 | 研究目的控制行 |
| `F-SELECTION` | integrity P5 | 无偏见选择；除非与研究问题有关，不纳入受保护特征 | 无偏见选择控制行 |
| `F-RISK` | integrity P6 | 识别并告知风险、不适、危害以及潜在收益 | 风险收益控制行 |
| `F-ANONYMITY` | integrity P7 | 所有参与者在任何时候保持匿名 | 匿名控制行 |

`Scope Decision` 和七行矩阵只综合上表内容；不增加 IRB 审批、书面同意、数据保留年限、加密方案、具体处分或其他输入未规定的程序。

## Golden 与可编辑性

- `solution/solve.py` 只实现本题的两份输入、精确段落锚点、三条批注、七项控制和政策整合表达，不是跨题生成器。
- 7 个非空政策段落被插入在 Turnitin 段之后、Special Needs 标题之前，使用 `w:ins` 原生修订并启用 `w:trackRevisions`；四项原则是原生编号列表。
- 附录位于下一页分节，含页眉、155 词 Scope Decision、8×4 原生矩阵、7 个书签和 7 个内部链接；矩阵表头带 `w:tblHeader`，数据行带 `w:cantSplit`。
- 评论通过 Word `comments.xml` 和正文范围标记精确锚定，而不是把批注文本伪装成正文。

## Judge 和负例设计

- Judge 从 `/app/input_files` 的两份冻结输入独立读取段落、run 格式、表格、图片、关系和页面几何，不读 Golden 哈希、解题脚本标记或候选自报结果。
- 评分分别覆盖模板保护、7 段政策完整/顺序/修订/锚点、Scope Decision、七行控制矩阵、三条精确批注、书签内链、原生节/页眉/表格和安全可打开性。
- 核心模板破坏上限 `0.35`；政策错误上限 `0.62`；Scope Decision、矩阵、批注缺失分别触发业务合理 cap；内链作为次要结构允许中等质量部分分。
- 本题负例覆盖：缺失/损坏、模板重写、课程表损坏、标识替换、未追踪插入、插错章节、缺匿名原则、无依据 IRB 规定、匿名矩阵事实错误、删批注、截图式作弊，以及内容完整但丢失内链的中等样例。

## 当前验收状态

- 完整上下文审计、v0.3 题面/Golden/Judge/负例和来源台账：已完成。
- 2026-09-07 使用项目专属 Python 3.11.16 环境执行 `validate-task`、`validate-task-assets`、`compileall` 和 `git diff --check`，全部通过。
- Golden 与从空目录独立生成的 fresh replay 各重复评分 5 次，10 次 Reward 均为 `1.000000`；同一候选的 5 份 Judge JSON 分别逐字节一致。Golden SHA256 为 `59683a988a05522604eeee090b6a450d8189e09a662a2bf17312f6b10e0a0985`，fresh 为 `efe4be8c201b74f46ad157a9963f37f06b11d549701c8ec8749a63efbd652334`；二者仅 ZIP 容器时间戳不同，25/25 个解压成员逐字节一致。
- 空交付与损坏 DOCX 均为 `0.0`；定向负例 Reward 为：模板改字 `0.35`、课程表损坏 `0.35`、标识替换 `0.35`、未追踪插入 `0.35`、插错章节 `0.62`、缺匿名原则 `0.62`、无依据 IRB 规则 `0.78`、匿名矩阵反向事实 `0.72`、删批注 `0.78`、截图式作弊 `0.106`；中等样例“内容完整但丢失七个内链”为 `0.94`。所有定向负例都未达 `0.8` 通过线，中等样例严格介于 0 与 Golden 之间。
- LibreOffice 7.3.7.2 打开/另存 Golden 和 fresh 后再次 Judge，两者均为 `1.0`。Judge 只对 LibreOffice 的等价序列化做窄范围兼容：`Hyperlink`/`InternetLink` 字符样式别名、末尾空段与分节属性合并、无内容页眉页脚距离归零；正文、样式、run 外观、表格、媒体、外链和核心页面几何仍全部逐项保护。
- 最终 Golden 与 fresh 各渲染 9 页，18/18 页逐页原图查看：第 1–4 页原大纲无裁切；第 5 页可见跟踪插入和四项编号；第 6–7 页课程表仅在行间分页；第 8 页校外条款完整；第 9 页页眉、Scope Decision、8×4 矩阵和状态说明均在一页内，无空白页、重叠、越界、字体缺失或不可读内容。
- 相关负例也用同一 LibreOffice 渲染并逐页原图查看：中等样例 9/9 页、匿名矩阵错误 9/9 页、未追踪插入 9/9 页、模板改字 9/9 页、截图式作弊 1/1 页。错误分别在语义结构、矩阵第六行、修订显示、首页使命文本和单页黑色图块中可见，且均无渲染器异常。
- Harbor Docker Oracle A/B：待当前题目冻结后执行；结果放题目外 repo-level audit，避免使已测 task digest 失效。

## 已知边界

- 作者侧视觉检查使用固定 LibreOffice/PDF/PNG 渲染，不进入 Reward，本题不引入 VLM。
- Word/LibreOffice 对修订显示颜色和目录缓存的呈现可能不同；Judge 检查真实 OOXML 修订、段落层级和锚点，不依赖界面颜色。
- 正式难度尚未校准；Oracle 成功只能证明题目可解，不能当作 Agent 难度证据。
