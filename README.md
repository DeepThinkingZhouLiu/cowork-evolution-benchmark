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

难度采用 `L1 / Easy`、`L2 / Medium`、`L3 / Hard`、`L4 / Expert` 的一一映射。四个主领域共享冻结 Agent Panel、最终 Reward 和重复运行协议，但使用各自的领域结构标准，并在领域内按 Panel 平均分排序。PPT 模块预先固定 `12 / 18 / 18 / 12` 配比；2026-09-04 已用 Claude Code、OpenCode 和 OpenHands 的 Sonnet 5 单轮答案完成 60 题试校准，逐题排名见 [`benchmarks/ppt-difficulty-pilot-20260904.csv`](benchmarks/ppt-difficulty-pilot-20260904.csv)。由于尚未完成多模型、多次重复，当前 148 道候选题的 `difficulty_level` 与 `difficulty_tier` 仍为 `unassigned`，不能把试校准结果伪装成最终冻结标签。

完整定义见 [`docs/task-taxonomy.md`](docs/task-taxonomy.md) 和 [`configs/taxonomy/cowork-office-v0.2.json`](configs/taxonomy/cowork-office-v0.2.json)。

## 当前逐题改造状态

仓库不再保留先前批量生成的 42 道合成题。当前题目必须从许可和来源可追溯的公开任务或真实候选池中逐题选择，并经过“查看原始输入 -> 阅读原始 Query、Rubric 和 Verifier -> 决定改造边界 -> 制作 Golden 与 Judge -> replay 和视觉检查”的人工流程。

目前已完成 **148 道真实上游来源的端到端候选题**，覆盖全部 42 个正式任务家族。它们由首轮 57 道候选题、后续逐题扩展的 41 道 PPT 题和 50 道 Cross-Office 深耦合题组成；PPT 与 Cross-Office 已分别达到阶段目标 60 道，Word 与 Excel 各有 14 道。新增题不会另起一套家族编号，而是映射回既有 42 家族；更细的差异通过操作类型、输入模态、约束画像和能力标签表达。论文规划的最终规模仍是 240 道，即 PPT、Word、Excel 和跨办公应用各 60 道，当前 148 道不是最终 Release 数量。

| 主领域       | 家族数 | 候选题数 | 主要来源                                                                                      | 当前验证状态                                      |
|--------------|-------:|---------:|-----------------------------------------------------------------------------------------------|---------------------------------------------------|
| 演示文稿     |     12 |       60 | OmegaUse-OfficeVal、Cowork-Bench、PPTArena、AutomationBench、OfficeBench、TheAgentCompany      | 60/60 Harbor Oracle 成功；3 Harness 单轮试校准完成，待重复运行冻结 |
| 文字文档     |      8 |       14 | OmegaUse-OfficeVal、LLM-DOCX-Editing、FORTE、Cowork-Bench                                     | 8 个家族全覆盖；待统一 Harbor 迁移和真实 Agent Rollout     |
| 电子表格     |     14 |       14 | OmegaUse-OfficeVal、Cowork-Bench、FORTE、SpreadsheetBench-2                                   | 14 个家族全覆盖；待统一 Harbor 迁移和真实 Agent Rollout    |
| 跨办公应用   |      8 |       60 | Cowork-Bench、OmegaUse-OfficeVal、OfficeBench、AutomationBench                                 | 60/60 均完成两次受测快照同 checksum Harbor Oracle 回放 |

这 148 道题都保留了许可证、原 Query/任务记录、Agent 可见输入 SHA256、改造说明和版本标识。输入来源分布如下：

| 上游公开数据集            | 题数 | 当前输入的来源与处理方式                                                                                       |
|---------------------------|-----:|----------------------------------------------------------------------------------------------------------------|
| OmegaUse-OfficeVal        |   36 | 上游 PPTX、DOCX、XLSX、PDF 和图片；数据卡说明指令经过隐私改写，输入经过 LLM 辅助重建和人工修订                |
| PPTArena                  |   31 | 上游原始 PPTX、原 Query 和人工 Ground Truth；逐题修正上游答案或评测器中可复现的缺陷                           |
| Cowork-Bench              |   18 | 上游任务包、数据库状态、CSV/JSON、PDF、邮件和工作区素材；在线服务部分被静态化为可离线审计输入                |
| AutomationBench           |   46 | 上游公开的合成业务状态、触发数据与最终状态断言；改造成可确定检查的复合演示或跨办公任务                         |
| TheAgentCompany           |    4 | 上游 Query、evaluator 和 GitLab、OwnCloud、RocketChat 等固定状态快照；不执行在线副作用                        |
| LLM-DOCX-Editing          |    3 | 上游原始 DOCX 和 Query；上游存在通过样例时，保留可验证的编辑结果                                               |
| FORTE                     |    2 | 上游公开的法律、财务 Office 文件、任务要求和 Rubric                                                           |
| OfficeBench               |    7 | 上游 PDF、图片、表格、发票、收入表和日历等真实样式输入与跨应用要求；适配为离线、可复算的多产物任务             |
| SpreadsheetBench-2       |    1 | 固定 Hugging Face revision 中的 LBO Excel 公式修复样例                                                        |
| **合计**                  | **148** | **全部登记为 `public_dataset`，每个 Agent 可见输入均记录 SHA256**                                            |

对 Cowork-Bench 的动态服务题，本项目只把上游固定数据库状态和服务结果静态导出，不增加新的订单、金额、评论或业务事实；对 OfficeVal 和 LLM-DOCX 的文件编辑题，保留上游办公文件和要求；对 FORTE 与 SpreadsheetBench-2 样例，保留上游标准答案、Rubric 或修复真值。新增的处理器、审计工作簿和评测器只为离线回放和更严格验收服务，不把自编数据冒充上游题。

这里的“真实上游来源”是指题目确实来自已发布的公开 Benchmark，并且输入、Query 和真值能够追溯到对应来源；它不等于“未经处理的企业生产数据”。当前 148 道题的 `source_type` 都是 `public_dataset`，其中 120 道按来源说明标记为 `synthetic=true`，28 道为 `synthetic=false`。`synthetic` 只描述上游数据的构造性质：例如 OfficeVal 明确说明其任务来自真实办公请求，但指令经过隐私改写、输入经过重建和人工脱敏；它不表示输入或 Query 是本项目临时编造的。148 道题都通过当前隐私检查并标记为可公开，其中 143 道记录为经过人工改写，5 道保留上游原始要求。

输入来源和 Golden 来源是两件事。上游存在 Golden 时，本项目保留或按其真值重建；上游只提供输入、任务要求或 Rubric 时，Golden 由制题者逐题阅读原始文件后，依据上游证据人工制作和检查，不能用统一脚本随机生成业务事实。Judge 必须从输入、Rubric 或冻结真值独立重算关键事实，避免只比较文件哈希或只接受参考解法的固定版式。

来源追踪仍有一个发布前风险：21 道 OfficeVal 题的 `source_uri` 仍指向 Hugging Face 的 `main` 路径。本轮新增 14 道已经全部固定到数据 revision；其余题的本地输入已经由 SHA256 固定，`provenance_id` 也记录了上游 commit，但正式 Release 前仍应把这些链接全部改成不可变 revision，保证第三方可以从同一个公开版本复核原文件。

任务包的全局结构、Schema、输入、Golden 和哈希校验已通过 `148/148`；现存 Golden 经各题 Judge 检查为 `148/148`，从空输出目录执行 `solve.sh -> Judge` 的 fresh replay 也是 `148/148`，空交付目录被拒绝 `148/148`。PPT 60 道与 Cross-Office 60 道均保留两次官方 Harbor 0.21 Docker Oracle 成功 trial，每题的两次受测任务快照 checksum 一致，Reward 均为 `1.0`且无异常。Cross-Office 重复产物的语义、OOXML 正文、结构和媒体负载一致；旧任务中少数 Office 包的整体字节差异仅来自 ZIP 时间戳或 `docProps/core.xml` 的 modified 时间。旧 19 道 PPT 还统一复核了空交付和损坏文件共 38 个负例；11 道有原始 PPTX 的任务额外检查了“直接把未修改输入改名交付”，全部低于通过阈值。每题的来源文件、修改边界、真值、replay、负例与视觉检查记录位于自己的 `adaptation/`、`metadata/` 和 `jobs/`。这些题仍是论文候选集；在非参考 Agent Rollout、Judge 梯度校准、泄漏分组和三段式切分完成前，不宣称为已冻结的论文 Release。

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

“能够被 Harbor 识别”和“能够通过官方 `harbor run` 完整回放”是两个不同层次。当前 148 道候选题都通过目录、Schema、输入、Golden 和哈希静态检查；其中 PPT 60 道和 Cross-Office 60 道已进一步完成官方 Harbor 0.21 Oracle Docker 双回放，并在当前工作树的 `jobs/` 下保留可复核的成功 trial、Reward 和实际产物。其余 Word 14 道和 Excel 14 道仍不能仅凭宿主机 fresh replay 宣称为 Harbor-ready。

| 检查项                                  | 当前结果 | 结论                                                                                         |
|-----------------------------------------|---------:|----------------------------------------------------------------------------------------------|
| 官方 Harbor 任务目录与配置静态解析      |  148/148 | 通过                                                                                         |
| 项目级结构、Schema、输入、Golden 和哈希 |  148/148 | 通过                                                                                         |
| 宿主机空目录 fresh replay               |  148/148 | 通过；仍不能代替容器回放                                                                     |
| 空交付拒绝                              |  148/148 | 通过                                                                                         |
| Harbor 容器内输入打包                   |  120/148 | PPT 60/60 与 Cross-Office 60/60 已完成；Word 与 Excel 各 14 道待迁移                     |
| Harbor 标准 Reward 文件                 |  120/148 | 错误答案会写 0 分而不是让 Verifier 异常                                                     |
| 无网络运行与单一产物采集配置            |  120/148 | 使用 `network_mode = "no-network"`，仅显式采集 `/app/output`                                |
| 完整官方 Harbor Oracle 端到端双回放     |  120/148 | PPT 60/60 与 Cross-Office 60/60 均为 Reward 1.0、0 异常；其余 28 道待统一回放          |

PPT 60 道与 Cross-Office 60 道已补齐 `environment/input_files/`、`/app/input_files` 默认入口、标准 Reward、`network_mode = "no-network"` 和单一 `/app/output` 采集路径。迁移过程中真实发现并修复了慢镜像源、容器缺输入、冻结 verifier 缺失、Reward 缺失和重叠 artifacts 等问题。当前只剩 Word 14 道与 Excel 14 道未完成官方 Harbor Oracle 回放。

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

公开任务以如下结构为统一目标。当前 PPT 60 道与 Cross-Office 60 道已经补齐 Harbor 所需的 `environment/input_files/`、标准 Reward、无网络运行和产物采集路径；Word 与 Excel 共 28 道需要完成同样迁移。

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

切分不能按行随机完成。相同输入文件、同源模板、同一客户材料、近重复指令和同一任务家族必须进入同一泄漏组，再以泄漏组为单位分配分区。

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

当前仓库已经完成第一版模块骨架、协议 Schema、多轴分类法、Harbor 题目模板，以及覆盖 42 个任务家族的 148 道真实上游候选题。PPT 与 Cross-Office 已分别达到 60 道阶段目标，并都完成 60/60 官方 Harbor Oracle 双回放。下一阶段从“把参考答案跑通”转为“用真实 Agent 检验难度和 Judge”：

- 对 PPT 60 道使用至少 3 个最新 Agent Harness 做 pass@1 / pass@8 Rollout，依据真实通过率、耗时、成本和失败类型校准 L1-L4，剔除过易、不可解或 Judge 不稳定的题。
- 给 PPT 60 道补充真实 Agent 产生的中等样本和视觉缺陷样本；当前 Oracle、空交付、损坏文件和定向语义负例只是 Judge 校准的第一层。
- 将 Word 与 Excel 剩余 28 道迁移到同一 Harbor 协议，再逐题完成容器回放；之后继续从真实上游来源扩展到每个领域 60 道。
- 仅对 Rollout 显示覆盖不足或方差过大的家族定向补题，不为凑数量批量制造输入、Query 或 Golden。
- 以来源、文件哈希和语义模板为泄漏组，冻结 `feedback / selection / final` 三个分区。
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
