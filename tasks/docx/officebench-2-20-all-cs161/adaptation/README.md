# officebench-2-20-all-cs161 v0.3 适配记录

状态：`local_validated`；difficulty 保持 `unassigned`，未执行付费 Harness 难度校准。

## 适配契约

- 角色与受众：系办课程数据管理员，为教务组制作可追溯的名册控制手册。
- Agent 可见输入：`CS161_Class_Roster.xlsx`，是学生身份、年级、专业与 5 项评分的唯一事实来源。
- 主交付物：`cs161_roster_control_book.docx`，只有 DOCX；不要求 PDF/XLSX/PPTX 或外部应用状态，所以仍属 DOCX 模块。
- 保留的上游意图：从 CS161 名册中将所有学生身份信息落到 Word，保持源顺序。
- v0.3 扩展：旧 `students.docx` 只有 9 个普通段落；新交付要求名册全字段控制表、专业/年级分布、5 项评分对账、4 项数据质量控制、按姓名索引、书签/内链和横竖页节。
- 设计区分：本题是“全名册源顺序与控制总数”手册；同源的 `cs-major-list`、`final-perfect` 和 `3-60-cs161-folder` 必须分别转向其他业务目标，不得复用本题结构。
- 迁移映射：任务 ID 保留；`0.2.0 / students.docx` -> `0.3.0 / cs161_roster_control_book.docx`。

## 上游与许可

- Dataset：OfficeBench task `2-20/0`。
- 来源仓库：`https://github.com/zlwang-cs/OfficeBench`。
- 冻结 commit：`b978b808667c32b52ce19a67ce1def1de9ae02b7`，已与 `003-Reference/002-Reference-Code/sources.lock.json` 核对。
- 许可：Apache-2.0；完整副本为 `adaptation/LICENSE-OfficeBench`，SHA256 为 `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`。
- 上游 Query 与 evaluator 参数保留为 `adaptation/source-task.json`，与上游字节相同，SHA256 为 `402315cb9ee0f0907e978c97ecd432d6b0c5bce45f23fb10d85f6b47d5ee9e62`。
- `adaptation/upstream-evaluator.py` 冻结了本题使用的关键词包含判定核心；完整上游 `utils/evaluate.py` 已逐段读取，SHA256 为 `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90`。它只检查 `Vufuh / Pom / S1006` 是否存在，无法验证完整名册、计算、顺序和 Word 结构，因此只作迁移审计证据。
- 上游资料是 benchmark fixture；provenance 继续标记 `synthetic=true`、`human_rewritten=true`、`publishability=public`，不冒充真实学校生产数据。

## 逐文件已读证据

| 文件 | 字节 / SHA256 | 实际检查 | 关键发现 |
|---|---|---|---|
| 上游 `tasks/2-20/subtasks/0.json` / 冻结 `source-task.json` | 454 / `402315cb…e62` | 逐字段读取用户、时间、Query 和唯一 evaluator | 原题要求全班 `Student ID, name`；evaluator 只检查 1 名学生的 3 个词 |
| 上游/打包 `CS161_Class_Roster.xlsx` | 10,383 / `c3289e2f…ed7` | `openpyxl` 读取 1 个工作表、10×10 单元格的全部值/类型/样式；LibreOffice 7.3.7.2 导出 2 页 PDF 并逐页原图查看 | 9 名学生，10 列；5 个评分列共45 个整数；原表默认打印时横向裁成 2 页，Word 名册需显式用横向页节 |
| `environment/input_files/CS161_Class_Roster.xlsx` | 10,383 / `c3289e2f…ed7` | 与公开输入和上游原件逐字节比较 | 三份完全一致，容器离线复现 |
| v0.2 Golden `students.docx` | 36,704 / `859cbd11…b70` | 逐段、ZIP/OOXML 和 LibreOffice 1 页渲染检查 | 只有 9 个 Normal 段落、0 表格、1 节；无源行、全字段、统计或导航 |
| 上游 `README.md` | 6,969 / `04104d99…d5e` | 完整读取 benchmark 目标、任务组织和 evaluator 运行方式 | OfficeBench 定位为多应用办公自动化 benchmark；本适配保留来源身份 |
| 上游 `utils/evaluate.py` | — / `f4a675a5…c90` | 读取 `_evaluate_contain_text` 和 `evaluate_contain` 的完整判定路径 | 忽略大小写的子串包含，对数字先去逗号；不验证排他性、顺序和结构 |
| 上游/打包 `LICENSE` | 11,357 / `c71d239d…ab4` | 完整读取 Apache-2.0 条款 | 允许修改和再分发，需保留许可和修改说明 |

## 从输入重算的真值台账

- 控制总数：9 个非空数据行，源行是 `Sheet1!2`至 `Sheet1!10`；Student ID 为 `S1000`至 `S1008`，全部唯一且符合 `S\d{4}`。
- 专业分布：CS 3（33.3%）、ECE 2（22.2%）、Math 3（33.3%）、Stats 1（11.1%）。
- 年级分布：Freshman 4（44.4%）、Junior 2（22.2%）、Senior 2（22.2%）、Sophomore 1（11.1%）。
- 评分对账：HW1 `N=9 / avg=30.44 / min=6 / max=58`；HW2 `9 / 58.11 / 3 / 94`；Midterm `9 / 53.44 / 6 / 88`；HW3 `9 / 50.44 / 0 / 96`；Final `9 / 53.89 / 1 / 100`。
- 数据质量：9/9 行十列全部非空；9/9 ID 唯一；9/9 ID 格式合法；45/45 评分是 0–100 的整数。
- 索引顺序：Cani/Jehecat -> Cixarim/Sipaq -> Dipuj/Fufowot -> Pegoqeju/Gafo -> Pivozu/Zoye -> Sude/Mec -> Vufuh/Pom -> Xihozun/Meki -> Yehamucu/Tukemoy。

## Golden、Judge 与负例设计

- `solution/solve.py` 只实现本题的 9 行名册、8 行分布、5 行评分对账、4 项质量控制、9 行索引和 Word 原生导航，不是跨题生成器。
- Golden SHA256：`325bb9b8a1dfcb5f3763d0e4f9583a8a8671fe4d84202a3d5a68460274dd912c`；程序独立生成到 `solution/`、`solution/Golden/` 和 fresh 目录后逐字节一致。
- Judge 从冻结 XLSX 独立重算所有人、所有字段、分布、评分统计和质量控制；检查表格、标题、字段、页节、书签、内链、可编辑性和 LibreOffice 渲染，不读 Golden 哈希或 solve 秘密标记。
- 本题专属负例：缺文件、损坏 DOCX、修改 S1006 Final、交换源行顺序、错专业分布、缺 Final 对账行、额外复制学生、九个名册内链失效、整页图片式作弊；中等样例只缺最后一项数据质量控制。

## 当前验收状态

- 完整上下文与来源审计：已完成。
- v0.3 题包、Golden、Rubric、Judge 和专属负例：已制作。
- 项目专用 Python 3.11.16 环境中 `validate-task`、`validate-task-assets`、`compileall` 和 `git diff --check` 均通过。
- Golden 与从空目录独立生成的 fresh 各连续评分 5 次，10 次均为 `1.000000`；三份独立生成的 DOCX 逐字节一致，SHA256 均为 `325bb9b8a1dfcb5f3763d0e4f9583a8a8671fe4d84202a3d5a68460274dd912c`。
- 空输出、损坏 DOCX 与整页图片式作弊均为 `0.0`；定向负例：错学生评分 `0.35`、错源顺序 `0.35`、错专业分布 `0.68`、缺 Final 对账 `0.68`、额外学生 `0.35`、内链失效 `0.72`，全部低于 `0.8` 通过线；中等样例缺 1 项质量控制，得分 `0.970833`。
- LibreOffice 7.3.7.2 打开/另存 Golden 和 fresh 后再次 Judge，均为 `1.0`。Golden/fresh 各 5 页、错学生评分 5 页、内链失效 5 页、中等样例 5 页和图片作弊 1 页，共 26 页逐页原图查看。Golden/fresh 的页面方向为纵向、纵向、横向、纵向、纵向；无空白页、裁切、重叠、越界、表行断裂或字体缺失。
- Harbor Docker Oracle A/B：待本地门槛全部通过后执行。

## 已知边界

- Excel 输入只含评分，不含课程政策或学分规则；本题不伪造通过/不通过、补考或学业处理结论。
- 目录是可更新字段；不同 Word 引擎可能在打开时才刷新页码。Judge 检查字段指令与标题层级，不依赖缓存页码。
- Oracle 满分只证明任务可解；正式难度等待冻结 Panel Rollout。
