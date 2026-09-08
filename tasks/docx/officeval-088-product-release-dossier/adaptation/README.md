# officeval-088-product-release-dossier v0.3 适配记录

状态：`local_validated_harbor_pending`；difficulty 仍为 `unassigned`，未运行付费 Harness 难度校准。

## 替换关系与 Word 模块边界

- 迁移映射：`cowork-evo/officebench-1-3-rename-docx@v0.2` -> `cowork-evo/officeval-088-product-release-dossier@v0.3`。
- 旧题的唯一输入是 13 KB、只含“PDF test file 2”的 DOCX；上游 evaluator 只检查 `test_file1.docx` 是否存在。三个 Sonnet 5 Harness 都只执行一次复制，输出 SHA256 与输入完全相同，收紧后仍为 1.0。该 Seed 没有可供业务整合的信息，因此整体 supersede，而不是在一句话后堆要求。
- 新题采用 Apache-2.0 的 OmegaUse-OfficeVal `officeval_088`，保留其“将 Word 模板按产品分析命名并发布”核心，但把“改名”放在真实的发布控制链末端：五源取证 -> 价格排名 -> 前三模板填充 -> 图片裁剪 -> 证据矩阵和批注 -> 以唯一合规文件名发布。
- 原 OfficeVal 要求同时交付 XLSX 和 DOCX。本适配明确是 Word-only derivative：XLSX 只用作字段语义输入，最终仅交付一份可编辑 DOCX，不要求交付 Excel，因此仍属 DOCX 模块。
- 本题没有 `visual_rubric`、VLM 或图像模型评分。图片根据输入媒体、裁剪比例、感知哈希、OOXML 定位和替代文本确定性验证；LibreOffice 全页查看只是作者侧验收。

## 上游、许可与来源贡献

- Dataset revision：`8e19c575b5be64c8a2964a4e28aafb696490c2ae`；代码/evaluator commit：`ffbeecb8752447c8e40b594a0eeb1db7236ecb36`。
- 许可：Apache-2.0。完整 `LICENSE`、`NOTICE`、代码仓 README、双语 Query/Rubric 和 1,536 行上游 evaluator 均冻结在 `adaptation/`。
- 五份产品 DOCX 贡献全部产品名称、型号、指导价格、核心参数、适用场景和产品原图；Word 模板贡献“产品摘要”和 3×2 产品表骨架；XLSX 模板只贡献五列字段语义。
- 适配层只贡献 Word-only 范围、排名/复算/证据边界、文档发布结构、题目专属 Solver/Judge/负例；没有新建产品业务事实。

## 逐文件主体已读证据

| 文件 / 文件组 | SHA256 / 结构 | 实际读取 | 关键发现 |
|---|---|---|---|
| 旧 OfficeBench Query / evaluator | `ab1c917a...5cb0` | 逐字段读取 `task` 和唯一 `evaluate_file_exist` | 只要求文件存在，不检查内容 |
| 旧题 instruction/task/rubric/judge/solution/adaptation/metadata | 全文 | 逐文件读取并解析 Solver 全部分支 | 旧 Judge 的核心是与输入字节完全一致，Golden/solution 也是字节复制 |
| 旧 input / Golden / solution | `764139ec...a0880`，1 段/0 表/1 节/11 个 ZIP 部件 | 读取全部段落、核心属性和 OOXML 成员 | 唯一文本为 `PDF test file 2`，三份字节相同 |
| 三个旧 Harness 产物、trajectory 与收紧后 Judge | 产物均 `764139ec...a0880` | 读取 Claude Code 2.1.251、OpenCode 1.4.14、OpenHands SDK 1.44.1 的实际复制命令、文件、result 和 7 项 criterion | 三者都没有读内容，一次 `cp` 后 1.0 |
| OfficeVal 双语 task / rubric | `d048e43c...45df` / `8381bc25...d23` / `35f1b693...5f19` / `5f310a4e...3b39` | 逐条读取 78 分钟原任务、双交付物范围、5 产品真值、图片、页眉页脚和命名规则 | 原任务同时要 XLSX/DOCX，其中 Word 部分已具有明确的前三选择和发布命名语义 |
| OfficeVal 上游 evaluator | `c3b92b1f...907d`，1,536 行 | 全文分块读取 OOXML、XLSX、图片、页眉页脚、定位、计分与入口 | 上游已做产品块绑定和图片比例校验，但需同时交付 Excel；本题不直接复用该计分器 |
| OfficeVal README / LICENSE / NOTICE | `58331fd5...83ddc` / `8b5a791d...f1375` / `6d5a662f...18c9` | README 227 行、Apache-2.0 201 行、NOTICE 6 行全文读取 | 框架强调去识别真实长链 Office 任务与确定性评估；许可允许修改和再分发并要求保留声明 |
| Word 模板 | `f3229e51...f782`，4 段/1 张 3×2 表/1 节/16 成员 | 逐段、逐单元格、样式和 ZIP 成员读取 | 只含产品摘要和三个产品/图片占位块；样式集没有标准 Title/Heading 名称 |
| XLSX 表头模板 | `ce9b5866...724f`，1×5/0 公式/0 图 | 逐单元格、列宽、行高、合并、冻结窗格、公式和绘图读取 | 只定义产品名称/核心参数/场景/指导价格/展示图五列，无业务记录 |
| 五份产品 DOCX | 详见 input manifest | 共 191 段；逐段读取名称、型号、价格、性能、场景、核心参数、注意事项，并查看 5 张内嵌原图 | 指导价为 4500/3999/3039/999/846；前三为云溪/统帅/小天鹅；图片都含营销文字，必须裁剪而非整图复制 |

## 真值、适配和 Judge 边界

- Judge 在每次评分时用 `python-docx` 从 5 份冻结说明独立抽取产品名称、型号、指导价格、全部核心参数和场景，根据价格数值自行排序；不导入 Solver，不读 Golden 哈希，不接受候选文档自报成功。
- 图片 Judge 逐产品块解析 drawing relationship，验证来源图裁剪、比例、显示高度、感知哈希和替代文本。容差允许合理的重新缩放，但不允许整张营销海报、张冠李戴或空白占位。
- 核心产品块是 hurdle：缺主要字段、顺序错或张冠李戴会触发分数上限。其他证据行、图片、批注、导航和页面对象均可按实体部分得分。

## 本地验收

- `validate-task`、`validate-task-assets`、compileall、JSON、7 份 data/environment 输入逐字节比对和 `git diff --check` 均通过。
- Golden/fresh 各重复 5 次均为 `1.0`，17/17 个 DOCX 解压成员逐字节一致；缺失、损坏、错文件名、模板破坏、排序错、事实错、证据行缺失、错误入选、错差额、整幅海报、图片错配、删批注、删页码域、破坏横向分节、虚构承诺和截图压平等定向负例均低于通过线，中等样例 `0.96`。
- LibreOffice 7.3.7.2 另存 Golden/fresh 后均为 `1.0`，两份另存件 18/18 个解压成员一致。7 份源输入 13 页、Golden/fresh/两份另存件 8 页和 4 个关键负例 7 页，共 28/28 页逐页查看；正向成品稳定为 2 页，无空白页、裁切、重叠、表行拆分、缺字或不可读内容。

## 当前状态

- 本地验收已闭环；尚待同一冻结快照的 Harbor Oracle A/B 双回放。
