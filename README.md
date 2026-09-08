# CoworkEvoBench：面向办公任务的 Harness 自进化基准

CoworkEvoBench 是一个面向论文研究的 Benchmark。它研究的问题是：在基础模型、代码智能体、办公工具、运行预算、初始 Harness 和评测器全部固定的前提下，一种 Evolution 算法能否利用训练题上的执行轨迹与评分反馈，持续改进做题 Harness，并在从未见过的办公任务上取得更高分数。

**本项目的核心资产是“题目 + 评测”。** Evolution 算法、Agent Harness 和模型都是被测对象或受控变量，不能反过来修改题目、评分规则和最终测试集。

## 研究对象

一次完整实验固定以下内容：

- 固定基础模型、代码智能体、办公工具和单题预算。
- 固定初始 Harness `H0`，包括提示词、`AGENTS.md`、技能、记忆、工具说明、脚本和工作流。
- 固定题目版本、数据划分、Rubric、Judge 和汇总规则。
- Evolution 算法只能读取允许公开的训练反馈，并在路径白名单内生成 Candidate Harness。
- 最终 Harness 锁定后，才允许在密封测试集上运行一次。

实验流程如下：

```text
固定 H0
-> 在反馈集运行并产生轨迹、产物、分数与失败信息
-> Evolution 算法生成 Candidate Harness
-> 检查修改范围、预算和安全边界
-> 在选择集做聚合评测并决定保留或回退
-> 重复若干 Epoch
-> 冻结最终 Harness
-> 在密封测试集运行 H0 与最终 Harness
-> 计算最终成功率和 Evolution Gain
```

## 设计原则

- **题目自包含**：每题都包含完整用户指令、输入文件、环境、Golden、Rubric 和 Judge，不依赖临时外链。
- **评分可复现**：确定性检查优先，语义与视觉 Judge 只处理程序难以判断的部分。
- **评测不可变**：Evolution 算法不能修改题目、Rubric、Judge、预算计量和晋升规则。
- **训练测试隔离**：按任务家族、输入文件哈希、来源和语义模板分组切分，禁止近重复样本跨集合出现。
- **运行条件配对**：H0 与 Candidate 使用相同题目、模型、预算、随机种子和 Judge。
- **大文件不进 Git**：仓库只保存协议、清单、哈希、公开样例和代码；私有文件由外部数据根目录管理。
- **模块单一职责**：题目生产、单题评分、结果汇总、Harness 运行和 Evolution 更新通过明确协议解耦。

PPTX/PDF 的主观视觉质量由公共 Codex-VLM 评分层处理：保留每题现有代码 Judge，
视觉题再增加任务特定 Rubric，分别保存确定性分和视觉分，并按 `70% / 30%` 聚合。
配置、安全边界、调用方法与失败语义见
[`docs/codex-vlm-judging.md`](docs/codex-vlm-judging.md)。

## 为什么不能只分 PPT、Excel 和 All

文件类型只说明“最后交什么”，不能说明“智能体到底做了什么”。例如，同样交付 PPTX，“从数据做决策汇报”“在模板中补内容”“只修版式”和“审计整份演示文稿”需要完全不同的能力。为了能分析 Evolution 到底让 Harness 学会了什么，本项目使用多轴分类：

```text
主领域
-> 任务家族
-> 操作类型
-> 输入模态与依赖范围
-> 通用能力
-> 约束画像
-> 难度层级
```

当前分类法包含 4 个主领域、42 个任务家族：

| 主领域       | 家族数 | 任务家族                                                                                                                                                                                                                                    |
|--------------|-------:|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| 演示文稿     |     12 | 材料成稿、大纲成稿、数据成稿、模板生成、内容编辑、视觉编辑、全稿润色、演示合并、受众适配、本地化、演示交付包、演示审计                                                                                                                      |
| 文字文档     |      8 | 材料成报告、提纲成文、模板填充、编辑改写、红线审阅、格式保真、版本合并比较、文档抽取                                                                                                                                                        |
| 电子表格     |     14 | 文档转表、数据清洗、表格变换、公式编写、模型构建、情景分析、透视分析、图表制作、管理看板、模板补全、表格修复、表格审计、多工作簿整合、分析成报告                                                                                            |
| 跨办公应用   |      8 | 抽取分析汇报、研究成果包、同源多交付物、更新传播、跨文件一致性、模板驱动交付包、多文件评审修订、工作区编排                                                                                                                                |

每题还会标注 `create / edit / transform / analyze / audit / repair / package / orchestrate` 等操作，以及单文件、多文件、跨应用或工作区依赖。论文中的主能力标签使用 `K1-K13`，覆盖内容抽取、证据落地、数据转换、计算建模、分析推理、结构化写作、编辑保持、视觉表达、模板遵循、跨文件整合、一致性传播、审计质检和工作区编排。甲方需求中的 `C1-C13` 原样保存在单独字段，只用于需求对齐，避免两套编号混在一起。

难度采用 `L1 / Easy`、`L2 / Medium`、`L3 / Hard`、`L4 / Expert` 的一一映射。四个主领域共享冻结 Agent Panel、最终 Reward 和重复运行协议，但使用各自的领域结构标准，并在领域内按 Panel 平均分排序。PPT 模块预先固定 `12 / 18 / 18 / 12` 配比；2026-09-04 已用 Claude Code、OpenCode 和 OpenHands 的 Sonnet 5 单轮答案完成 60 题试校准，逐题排名见 [`benchmarks/ppt-difficulty-pilot-20260904.csv`](benchmarks/ppt-difficulty-pilot-20260904.csv)。由于尚未完成多模型、多次重复，当前 240 道候选题的 `difficulty_level` 与 `difficulty_tier` 仍为 `unassigned`，不能把试校准结果伪装成最终冻结标签。

完整定义见 [`docs/task-taxonomy.md`](docs/task-taxonomy.md) 和 [`configs/taxonomy/cowork-office-v0.2.json`](configs/taxonomy/cowork-office-v0.2.json)。

## 当前逐题改造状态

**本开发分支新增：七道 Excel 满分题已逐题换成 Workspace 来源的复杂任务。**不是只改 Judge：输入、任务目标、Golden、Rubric、独立 Judge 和公式变更测试均已重做；7/7 官方 Harbor Oracle 通过，Excel 总数仍60，全仓仍240。详细映射和证据见 [七题升级记录](docs/xlsx-hardening-7-progress.md)。输入是可追溯的上游模拟办公素材，不是代理编造的私人业务数据；不新增VLM。旧七题分数不再代表新题，尚未重跑真实Agent。

下面的v0.3是主线历史发布记录；**它的内容锁不能加载本分支替换后的七题**。七题已提交到 `lz/data-xlsx-hardening-7`，新快照为 [v0.4](benchmarks/cowork-evo-240-v0.4/README.md)，保持90/30/120与其他233题不变，0个跨分区来源组件。尚未合入main；Claude重测入口为 `tools/run_xlsx_hardening_claude.py`，仅新测7题、核验哈希后复用53题，难度结果是单Harness试分档。

仓库不再保留先前批量生成的 42 道合成题。当前题目必须从许可和来源可追溯的公开任务或真实候选池中逐题选择，并经过“查看原始输入 -> 阅读原始 Query、Rubric 和 Verifier -> 决定改造边界 -> 制作 Golden 与 Judge -> replay 和视觉检查”的人工流程。

目前已完成 **240 道真实上游来源的端到端任务**，覆盖全部 42 个正式任务家族。它们由首轮 57 道候选题、后续逐题扩展的 41 道 PPT 题、50 道 Cross-Office 深耦合题，以及 CXY 分支带来的 46 道 DOCX 和 46 道 XLSX 题组成；四个模块现在均为 60 道。新增题不会另起一套家族编号，而是映射回既有 42 家族；更细的差异通过操作类型、输入模态、约束画像和能力标签表达。17 道旧 Word 满分题已逐题重构并合入 `main`，当前快照为 [`cowork-evo-240-v0.3`](benchmarks/cowork-evo-240-v0.3/README.md)；正式难度标签仍需后续完成。

完整代码与题目的统一入口是本仓库根目录；`tasks/docx/` 包含全部 60 道 Word 题，不必分别进入 A/B 工作树取文件。17 道升级题均已完成专属输入、Golden、Rubric、代码 Judge、负例及两轮 Harbor Oracle 回放，具体证据见 [A 组审计](docs/docx-hardening-a-9-20260907.md) 与 [B 组审计](docs/docx-hardening-b-8-audit.md)。Oracle 满分只说明参考解和评测链可运行，不代表真实 Agent 得分；新版三套 Harness 评测计划见 [整合与重测说明](docs/docx-hardening-integration-20260908.md)。

| 主领域       | 家族数 | 候选题数 | 主要来源                                                                                      | 当前验证状态                                      |
|--------------|-------:|---------:|-----------------------------------------------------------------------------------------------|---------------------------------------------------|
| 演示文稿     |     12 |       60 | OmegaUse-OfficeVal、Cowork-Bench、PPTArena、AutomationBench、OfficeBench、TheAgentCompany      | 60/60 Harbor Oracle 成功；3 Harness 单轮试校准完成，待重复运行冻结 |
| 文字文档     |      8 |       60 | OmegaUse-OfficeVal、OfficeBench、LLM-DOCX-Editing、FORTE、Cowork-Bench                       | 60/60 Harbor Oracle 双回放；3 Harness 单轮 Panel 完成，Judge 已收紧 |
| 电子表格     |     14 |       60 | OfficeBench、SpreadsheetBench-2、Spreadsheet-RL、AutomationBench、FORTE、Workspace-Bench 等   | 60/60 Harbor Oracle 双回放；真实 Agent Rollout 待做 |
| 跨办公应用   |      8 |       60 | Cowork-Bench、OmegaUse-OfficeVal、OfficeBench、AutomationBench                                 | 60/60 均完成两次受测快照同 checksum Harbor Oracle 回放 |

这 240 道题都保留了许可证、原 Query/任务记录、Agent 可见输入 SHA256、改造说明和版本标识。输入来源分布如下：

| 上游公开数据集            | 题数 | 当前输入的来源与处理方式                                                                                       |
|---------------------------|-----:|----------------------------------------------------------------------------------------------------------------|
| OmegaUse-OfficeVal        |   32 | 上游 PPTX、DOCX、XLSX、PDF 和图片；保留固定 revision 与逐题改造记录                                           |
| PPTArena                  |   31 | 上游原始 PPTX、Query 和 Ground Truth；逐题改造为可回放的 PPT 任务                                            |
| Cowork-Bench              |   20 | 上游任务包、数据库状态和工作区素材；在线服务状态静态化为可离线审计输入                                       |
| AutomationBench           |   50 | 上游业务状态、触发数据与最终状态断言；改造成可确定检查的演示或跨办公任务                                   |
| OfficeBench               |   61 | 上游文档、表格、PDF、图片和日历素材；适配为离线、可复算的办公任务包                                         |
| SpreadsheetBench-2        |   19 | 固定 Hugging Face revision 中的 Excel 模型与公式修复样例                                                     |
| Spreadsheet-RL            |   12 | 固定 revision 中的结构化表格、公式和数据处理任务                                                           |
| Workspace-Bench           |    4 | 固定工作区 fixture 的任务元数据与输入快照                                                                  |
| TheAgentCompany           |    6 | 固定 GitLab、OwnCloud、RocketChat 等状态快照；不执行在线副作用                                               |
| LLM-DOCX-Editing          |    3 | 上游 DOCX 与 Query；保留可验证的编辑结果                                                                    |
| FORTE                     |    2 | 上游公开的法律、财务 Office 文件、任务要求和 Rubric                                                         |
| **合计**                  | **240** | **每个 Agent 可见输入均记录 SHA256；具体来源和许可见各题 `metadata/provenance.json`**                         |

对 Cowork-Bench 的动态服务题，本项目只把上游固定数据库状态和服务结果静态导出，不增加新的订单、金额、评论或业务事实；对 OfficeVal 和 LLM-DOCX 的文件编辑题，保留上游办公文件和要求；对 FORTE 与 SpreadsheetBench-2 样例，保留上游标准答案、Rubric 或修复真值。新增的处理器、审计工作簿和评测器只为离线回放和更严格验收服务，不把自编数据冒充上游题。

这里的“真实上游来源”是指题目确实来自已发布的公开 Benchmark，并且输入、Query 和真值能够追溯到对应来源；它不等于“未经处理的企业生产数据”。当前 240 道题均保留来源类型、revision、许可和输入哈希；不同数据集对 synthetic、隐私改写和答案容器的定义不同，不能把这些字段混成一个结论。OfficeVal 等数据集的指令可能经过隐私改写、输入重建和人工脱敏，但这不表示输入或 Query 是本项目临时编造的。

输入来源和 Golden 来源是两件事。上游存在 Golden 时，本项目保留或按其真值重建；上游只提供输入、任务要求或 Rubric 时，Golden 由制题者逐题阅读原始文件后，依据上游证据人工制作和检查，不能用统一脚本随机生成业务事实。Judge 必须从输入、Rubric 或冻结真值独立重算关键事实，避免只比较文件哈希或只接受参考解法的固定版式。

来源追踪仍有一个发布前风险：部分历史 OfficeVal 题的 `source_uri` 仍指向 Hugging Face 的 `main` 或 `resolve/main` 路径。当前本地输入已经由 SHA256 固定，`provenance_id` 也记录了上游 commit；正式 Release 前仍应把所有可变链接改成不可变 revision，保证第三方可以从同一个公开版本复核原文件。

任务包的目录结构、输入、Golden 和哈希静态校验已通过 `240/240`；四个模块历史版本完成了双轮官方 Harbor 0.21 Docker Oracle 回放，本版变动的 17 道 Word 题也各自重新完成双轮回放，Reward 均为 `1.0` 且无异常。每题的来源文件、修改边界、真值和作者侧验证记录位于自己的 `adaptation/` 与 `metadata/`。三段式任务归属、任务内容哈希和泄漏组件已在 `cowork-evo-240-v0.3` 冻结；论文正式结果仍需区分历史评分和新版重测，并完成全模块重复运行与最终难度校准。

本轮 14 道题不是用统一脚本批量造输入或随机生成 Golden，而是逐题读取上游 Query、输入、Rubric 和 Verifier 后分别构造：

| 领域       | 题号 | 上游输入与逐题改造                                                                 | Golden 和 Judge 的关键真值                                                                 |
|------------|-----:|------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------|
| 文字文档   |  003 | 读取真实会议记录 DOCX，清理表格空白并改为会议内容横线式表格                       | 核对 95 行结构、13 个待清理空白、内容保持和 Word 可编辑性                                  |
| 文字文档   |  010 | 读取压平流程图 DOCX 和参考图，人工重建可编辑技术路线图                           | 检查 30 个图形、39 条连接线、目标行删除、无图片替代和页面布局                              |
| 文字文档   |  029 | 读取表格图片，重建为可编辑 Word 培训记录表                                       | 按图片真值检查表头、字段、合并关系、可编辑表格和版式；上游 Verifier 矛盾单独留档           |
| 文字文档   |  033 | 读取 1,435 个真实零售价，逐项乘以 1.1 并保留目录格式                             | 使用十进制定点舍入独立重算全部价格，同时检查非目标文字与样式未被改写                       |
| 文字文档   |  040 | 读取真实简介素材和格式要求，完成样式、段落、页面与水印设置                       | 检查要求中的字号、段落、页面、图片和水印；上游一项标题误报已记录                           |
| 演示文稿   |  043 | 从中英名片图片重建两页可编辑 PPT                                                 | 两页分别含 52/55 个原生对象；仅二维码保留为图片，并检查双语文字与布局                      |
| 演示文稿   |  046 | 从两张研究框架图分别重建两页可编辑图示                                           | 两页分别含 132/108 个原生对象，拒绝整页截图替代，并检查节点、连线和视觉层级                |
| 演示文稿   |  051 | 从异常检测模型图重建可编辑流程                                                   | 检查 166 个原生对象、关键标签、关系和布局，要求零嵌入参考图                                |
| 演示文稿   |  054 | 从工业水回用流程图重建可编辑流程                                                 | 检查 114 个原生对象和流程顺序；依据真实输入纠正上游第 8 步事实错误                         |
| 演示文稿   |  056 | 从客户投诉闭环图重建可编辑流程                                                   | 检查 107 个原生对象、字体、节点、连接关系和零截图替代                                      |
| 演示文稿   |  060 | 读取五页演示文稿，只修复指定单页的图片布局                                       | 限定只改第 4 页 XML；其他四页逐像素一致，源图片媒体不变                                    |
| 演示文稿   |  064 | 读取模板和 12 张真实照片，按顺序完成两页四种画框版式                            | 检查 1→12 顺序、居中裁切、无拉伸、模板 46 个原对象保持和双线标题框可编辑                   |
| 跨办公应用 |  059 | 读取四页 Word 申报书和 18 页 PPT，在指定单元格加入真实相对链接                   | PPT 必须逐字节不变；Word 链接目标、显示文字、页面和未要求内容均检查；上游实体名错误已留档  |
| 跨办公应用 |  078 | 读取演示文稿和三表工作簿，在第 8 页加入可编辑 Excel 链接按钮                    | 工作簿必须逐字节不变；仅第 8 页及其关系文件可变，并检查按钮样式与相对链接                  |

参考仓库审计见 [`datasets/reference-source-audit.md`](datasets/reference-source-audit.md)。首轮 57 道候选题见 [`datasets/task-candidate-map.md`](datasets/task-candidate-map.md)，PPT 扩展 41 道及 60 道总账见 [`docs/ppt-task-intake-v0.1.md`](docs/ppt-task-intake-v0.1.md)，Cross-Office 逐题扩展总账见 [`docs/cross-office-task-intake-v0.1.md`](docs/cross-office-task-intake-v0.1.md)。

## Harbor 兼容性现状

“能够被 Harbor 识别”和“能够通过官方 `harbor run` 完整回放”是两个不同层次。当前 240 道候选题不仅通过目录、输入、Golden 和哈希静态检查，也全部完成官方 Harbor 0.21 Oracle Docker 双回放；每轮都从独立空输出目录开始，并保留成功 Trial、Reward、任务 checksum 和实际产物。

| 检查项                                  | 当前结果 | 结论                                                                                         |
|-----------------------------------------|---------:|----------------------------------------------------------------------------------------------|
| 官方 Harbor 任务目录与配置静态解析      |  240/240 | 通过目录级资产校验与逐题容器回放                                                               |
| 项目级结构、输入、Golden 和哈希         |  240/240 | 通过                                                                                           |
| 独立空输出目录 Oracle fresh replay      |  240/240 | 四个模块均完成两次官方 Harbor Docker 回放                                                        |
| 空交付拒绝                              |   待统一 | 各题 Judge 已包含相应检查，尚未形成 240 题统一汇总                                               |
| Harbor 容器内输入打包                   |  240/240 | 四个模块均有 `environment/input_files/`                                                        |
| Harbor 标准 Reward 文件                 |  240/240 | 四个模块的两轮回放均成功生成标准 Reward                                                         |
| 无网络运行与单一产物采集配置            |  240/240 | 全部任务显式使用 `network_mode = "no-network"`                                                 |
| 完整官方 Harbor Oracle 端到端双回放     |  240/240 | PPT、DOCX、XLSX 与 Cross-Office 各 60 道均完成                                                   |

四个模块均已补齐 `environment/input_files/`、标准 Reward、`network_mode = "no-network"` 和单一 `/app/output` 采集路径。DOCX/XLSX 最终整批回放作业为 `docx-60-final-a-20260906`、`docx-60-final-b-20260906`、`xlsx-60-final-a-20260906` 和 `xlsx-60-final-b-20260906`：四轮均为 60/60、均值 1.000、0 异常，每题两轮 checksum 一致，OOXML 正文、结构与媒体负载无差异。逐题机器可读证据见 [`docs/docx-xlsx-final-replay-20260906.json`](docs/docx-xlsx-final-replay-20260906.json)。

`metadata/validation.json` 是回放后写入的证据记录；Harbor 0.21 的全目录 checksum 会把该记录本身算入，因此“当前目录 checksum”与已记录的“受测快照 checksum”不应被混为同一个值。精确的受测快照由 `jobs/<job>/<trial>/lock.json` 和 `result.json` 冻结。

Cross-Office 最终整批回放作业为 `cross-office-60-final-a-20260902` 和 `cross-office-60-final-b-20260902`：两轮都是 60/60、均值 1.000、0 异常，每题两轮 checksum 相同且与回放后未再改动的当前任务目录 checksum 一致。每轮产出 343 份文件；52 道的两轮产物逐字节一致，另外 8 道的 OOXML 正文、结构和媒体负载一致，仅 ZIP 时间戳或 `docProps/core.xml` 的 modified 时间不同。逐题机器可读证据见 [`docs/cross-office-final-replay-20260902.json`](docs/cross-office-final-replay-20260902.json)。

## 目录结构与模块职责

```text
.
├── docs/                       # 研究协议、架构、数据治理和题目制作规范
├── schemas/                    # 跨模块交换对象的机器可校验契约
├── configs/                    # Benchmark、Harness、评测和 Evolution 的实验配置
├── datasets/                   # 数据来源清单、版本锁和整理记录，不保存私有大文件
├── tasks/                      # 以 Harbor 完整兼容为目标的候选题包、公开样例和制题模板
├── benchmarks/                 # 冻结后的 Benchmark Release 与三段式数据划分
├── harnesses/                  # 初始 Harness 定义，不保存运行期间的 Candidate
├── mutation-policies/          # 三条赛道的可写路径与禁止修改范围
├── evolvers/                   # Evolution 算法的标准接入位置与参考实现
├── src/cowork_evo_bench/       # 控制、校验、评测、适配和报告代码
├── tests/                      # 单元测试、集成测试和协议夹具
└── .coworkevo/                 # 本地运行、产物、Candidate 和报告，不提交 Git
```

各模块的责任边界如下：

| 模块                 | 负责内容                                      | 不负责内容                                  |
|----------------------|-----------------------------------------------|---------------------------------------------|
| `tasks/`             | 单题指令、输入、环境、Golden、Rubric、Judge   | 数据集切分、Epoch 调度、排行榜汇总          |
| `judging/`           | 解析办公产物并产生逐条 Rubric 评分            | 选择 Candidate、更新 Harness                |
| `evaluation/`        | 汇总逐题结果、配对 H0/Candidate、计算统计指标 | 启动 Agent、读取隐藏答案后指导 Evolution    |
| `benchmarks/`        | 冻结题目版本、划分和可见性                    | 保存运行产物、直接实现 Judge                |
| `datasets/`          | 数据来源、许可、隐私、去重和版本追踪          | 把原始候选数据直接认定为正式 Benchmark      |
| `harnesses/`         | 定义可复现的初始 H0                           | 保存 Judge、密封题目或运行时 Candidate      |
| `evolvers/`          | 读取反馈并输出 Harness 修改                   | 修改题目、评分、预算或 Controller            |
| `mutation-policies/` | 定义每条赛道允许修改的路径                    | 依靠提示词代替真正的文件权限与 Diff 检查    |
| `controller/`        | 编排 Epoch、隔离 Candidate、记录谱系与决策    | 替 Evolution 算法分析失败原因或决定具体修改 |
| `adapters/`          | 对接 Harbor、Agent Harness、Evolver 和 Judge  | 承担核心业务规则                            |
| `reporting/`         | 生成论文表格、逐题报告和可审计运行摘要        | 重新计算与评测模块不一致的分数              |

## 单题格式

公开任务以如下结构为统一目标。当前四个模块各 60 道都具备 `environment/input_files/`、Golden、Judge 和脚本；PPT 与 Cross-Office 已完成官方 Oracle 双回放，DOCX 与 XLSX 仍需完成同样的容器回放和统一配置审计。

```text
tasks/<领域代码>/<任务编号>/
├── instruction.md
├── task.toml
├── rubric.json
├── environment/
│   ├── Dockerfile
│   └── input_files/           # 打入 Harbor 环境、供 Agent 读取的输入副本
├── data/
│   └── input_files/           # 作者侧来源管理副本，必须与环境输入逐字节一致
├── solution/
│   ├── solve.sh
│   ├── solve.py
│   └── Golden/                # 隐藏参考交付物
├── tests/
│   ├── judge.py
│   ├── test.sh
│   ├── upstream_verifier.py   # 上游提供时保留；本地 Judge 不盲从其缺陷
│   └── expectations.json      # 可选，声明确定性检查目标
├── metadata/
│   ├── task.json              # 多轴分类与论文分析元数据
│   ├── provenance.json        # 来源、许可、隐私状态与输入哈希
│   └── input-manifest.json    # 逐文件大小、哈希和用途
└── adaptation/                # 原 Query、Rubric、Verifier、许可和逐题改造记录
```

目标运行协议中，Agent 只能看到 `instruction.md`、允许公开的 `task.toml` 字段，以及由 `environment/` 提供的输入文件。`solution/`、完整 `rubric.json`、`tests/`、Golden 和 Judge 凭据属于隐藏评测侧。仓库中的 `data/input_files/` 是作者侧来源管理位置，不能假设 Harbor 会自动把它挂载到容器中。

每题至少包含六条客观、原子、可区分的 Rubric，并覆盖内容、数据准确性、事实性、结构、呈现、文件完整性或专业规范等多个维度。分数统一归一化到 `[0, 1]`，第一版全局成功阈值固定为 `0.7`。

## Benchmark 数据划分

每个正式 Release 包含三个互斥分区：

- `feedback`：Evolution 算法可以看到详细题目、轨迹、逐项得分和失败证据。
- `selection`：只返回聚合指标，用于 Epoch 期间选择 Candidate，不能泄露逐题答案。
- `final`：完全密封，最终 Harness 锁定后才能运行，只用于最终论文报告。

当前冻结划分见 [`benchmarks/cowork-evo-240-v0.3/`](benchmarks/cowork-evo-240-v0.3/README.md)。新版新增真实素材导致旧分区出现共享输入，因此在不改变其他模块分区的前提下，最少移动 9 道 Word 题，保持下表数量并消除跨分区泄漏；完整变更见 [migration.json](benchmarks/cowork-evo-240-v0.3/migration.json)。历史 v0.1/v0.2 原样保留，不能用于加载当前任务内容。

| 模块         | Feedback | Selection | Final | Final 角色       |
|--------------|---------:|----------:|------:|------------------|
| PPT          |       30 |        10 |    20 | In-domain        |
| DOCX         |       30 |        10 |    20 | In-domain        |
| XLSX         |       30 |        10 |    20 | In-domain        |
| Cross-Office |        0 |         0 |    60 | Compositional OOD |
| 合计         |       90 |        30 |   120 | 60 ID + 60 OOD   |

切分不是逐行随机完成。相同输入文件、相同 provenance、同源模板、同一客户材料、近重复指令，以及同一任务家族的参数替换变体必须先进入同一泄漏组件，再以组件为单位分配分区。PPT 还按现有试校准档位等比例分层；DOCX/XLSX 在正式经验难度可用前，按任务家族、来源、依赖范围和预计耗时分层。

## 评测结果

Judge 首先输出逐条 Rubric 结果，Task Evaluator 再归一化为统一的单题结果。Benchmark Evaluator 只消费标准结果，不依赖某个具体 Agent 或 Harbor 内部实现。

论文主指标为：

- 最终成功率：最终 Harness 在密封测试集上的成功比例。
- Evolution Gain：最终 Harness 成功率减去 H0 成功率。

同时必须报告平均 Rubric 分、配对净提升、退化题数量、置信区间、Token、时间和成本，避免二值成功率掩盖部分提升或明显退化。

单题 Judge 按以下顺序工作：

```text
文件存在且可打开
-> Office 结构、数值、公式与保留区域检查
-> 内容与引用检查
-> 页面渲染和视觉检查
-> 必要时进行语义或专业质量检查
-> 逐条 Rubric 聚合为 0–1 分数
```

能够稳定程序化判断的内容必须使用确定性检查；视觉或语义 Judge 只处理程序难以准确判断的部分。正式题至少需要 Golden、中等缺陷样本、明显负样本和损坏文件样本，且同一 Judge 重复运行的波动要受控。

## Evolution 赛道

- `prompt-only`：只允许修改系统提示词、`AGENTS.md` 和少量示例。
- `skill-memory`：允许修改技能、经验库、检索索引和加载规则。
- `full-harness`：允许修改完整做题 Harness，但仍不能触碰可信控制面。

路径白名单必须同时通过运行时只读挂载和结束后的 Diff 复核落实，不能只在提示词中声明。

## 私有数据管理

真实用户文件、企业材料和密封测试题不提交到 Git。`datasets/` 仅保存来源、许可、下载日期、文件哈希、隐私状态、转换版本和发布权限。代码通过独立的数据根目录读取资产，并使用 `dataset.lock` 固定精确版本。

原始 273 条 Paper2Any 数据只是候选池。进入正式 Benchmark 前仍需完成隐私审查、任务有效性筛选、内容哈希去重、Golden 补全、Rubric/Judge 制作、Judge 校准和分组切分。

本项目同时维护一份只读的开源参考区，其中固定了 16 个 Office / Cowork 相关仓库的精确 commit，用于研究任务结构、原子 Rubric、Golden 配对、工作区恢复和评测方法。参考代码与正式题目严格分离；仓库可见不等于二进制素材可以直接再发布，所有派生数据仍需逐项检查许可、隐私和来源。

## 本地验证

项目使用独立 Conda 环境，默认位置为：

```text
REMOVED_LOCAL_PATH
```

常用检查如下：

```bash
环境路径=REMOVED_LOCAL_PATH

# 运行项目级公开题校验；它不等同于官方 harbor run
PYTHONPATH=src "$环境路径/bin/python" -m cowork_evo_bench.cli validate-public-tasks tasks

# 运行仓库单元测试
PYTHONPATH=src "$环境路径/bin/python" -m unittest discover -s tests/unit
```

## 当前阶段

当前仓库已经完成第一版模块骨架、协议 Schema、多轴分类法、Harbor 题目模板，以及覆盖 42 个任务家族的 240 道真实上游候选题。四个模块均达到 60 道目标并完成官方 Harbor Oracle 双回放。下一阶段从“把参考答案跑通”转为“用真实 Agent 检验难度和 Judge”：

- 对 PPT 60 道使用至少 3 个最新 Agent Harness 做 pass@1 / pass@8 Rollout，依据真实通过率、耗时、成本和失败类型校准 L1-L4，剔除过易、不可解或 Judge 不稳定的题。
- 给 PPT 60 道补充真实 Agent 产生的中等样本和视觉缺陷样本；当前 Oracle、空交付、损坏文件和定向语义负例只是 Judge 校准的第一层。
- 对 XLSX 60 道运行冻结的三 Agent Harness Panel；对已完成单轮 Panel 的 DOCX 60 道补重复运行，收集逐题 Reward、失败类型、耗时与成本。
- 仅对 Rollout 显示覆盖不足或方差过大的家族定向补题，不为凑数量批量制造输入、Query 或 Golden。
- 复核已经冻结的 `feedback / selection / final` 分区，在任务内容或 Judge 发生变化时发布新版本，禁止原地改写已有 Release。
- 跑通同模型、同预算、同随机种子的 H0 与 Candidate 配对实验，再开始 Evolution 算法横向比较。

最终需要跑通的最小论文闭环是：

```text
冻结题目与 H0
-> feedback 上执行并收集轨迹与逐项反馈
-> Evolution 算法修改允许范围内的 Harness
-> selection 上决定保留或回退
-> 重复固定 Epoch 并冻结最终 Harness
-> final 上分别运行 H0 与最终 Harness
-> 报告成功率、平均分、Evolution Gain、退化题、置信区间和成本
```
