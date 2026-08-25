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

完整定义见 [`docs/task-taxonomy.md`](docs/task-taxonomy.md) 和 [`configs/taxonomy/cowork-office-v0.1.json`](configs/taxonomy/cowork-office-v0.1.json)。

## 当前逐题改造状态

仓库不再保留先前批量生成的 42 道合成题。当前题目必须从许可和来源可追溯的公开任务或真实候选池中逐题选择，并经过“查看原始输入 -> 阅读原始 Query、Rubric 和 Verifier -> 决定改造边界 -> 制作 Golden 与 Judge -> replay 和视觉检查”的人工流程。

目前已完成 **43 道真实上游来源的端到端候选题**，覆盖全部 42 个任务家族。`formatting_fidelity` 有 2 道异质题，其余每个家族各 1 道。

| 主领域       | 家族数 | 候选题数 | 主要来源                                      | 基础验证状态                                      |
|--------------|-------:|-----------:|-----------------------------------------------|---------------------------------------------------|
| 演示文稿     |     12 |         12 | OmegaUse-OfficeVal、Cowork-Bench            | 12 个家族均有原始 Query、完整输入、Golden、Judge 和 replay |
| 文字文档     |      8 |          9 | OmegaUse-OfficeVal、LLM-DOCX-Editing、FORTE、Cowork-Bench | 8 个家族全覆盖；格式保真含 2 道不同任务          |
| 电子表格     |     14 |         14 | OmegaUse-OfficeVal、Cowork-Bench、FORTE、SpreadsheetBench-2 样例 | 14 个家族全覆盖；含可动态回放的三工作簿整合题       |
| 跨办公应用   |      8 |          8 | Cowork-Bench、OmegaUse-OfficeVal             | 8 个家族均含跨文件一致性或工作区回放检查          |

这 43 道题都保留了固定上游版本、许可证、原 Query/任务记录、Agent 可见输入 SHA256 和改造说明。输入来源分布如下：

| 上游公开数据集              | 题数 | 当前输入的来源与处理方式                                                                                         |
|-----------------------------|-----:|------------------------------------------------------------------------------------------------------------------|
| OmegaUse-OfficeVal          |   22 | 上游 PPTX、DOCX、XLSX、PDF 和图片；上游数据卡说明指令经过隐私改写，输入经过 LLM 辅助重建和人工修订              |
| Cowork-Bench                |   15 | 上游任务包、数据库状态、CSV/JSON、PDF、邮件和工作区素材；依赖在线服务的部分被静态化为可离线审计的输入          |
| LLM-DOCX-Editing            |    3 | 上游原始 DOCX 和原始 Query；上游存在通过样例时，保留其可验证的编辑结果                                          |
| FORTE                       |    2 | 上游公开的法律、财务 Office 文件、任务要求和 Rubric                                                            |
| SpreadsheetBench-2 Sample   |    1 | 固定 Hugging Face revision 中的 LBO Excel 公式修复样例                                                         |
| **合计**                    | **43** | **全部登记为 `public_dataset`，每个 Agent 可见输入均记录 SHA256**                                               |

对 Cowork-Bench 的动态服务题，本项目只把上游固定数据库状态和服务结果静态导出，不增加新的订单、金额、评论或业务事实；对 OfficeVal 和 LLM-DOCX 的文件编辑题，保留上游办公文件和要求；对 FORTE 与 SpreadsheetBench-2 样例，保留上游标准答案、Rubric 或修复真值。新增的处理器、审计工作簿和评测器只为离线回放和更严格验收服务，不把自编数据冒充上游题。

这里的“真实上游来源”是指题目确实来自已发布的公开 Benchmark，并且输入、Query 和真值能够追溯到对应来源；它不等于“未经处理的企业生产数据”。当前 43 道题的 `source_type` 都是 `public_dataset`，按保守口径也都标为 `synthetic=true`：OfficeVal 明确说明其任务源自真实办公请求，但指令经过隐私改写、输入经过重建和人工脱敏；其余来源也属于公开 Benchmark 的 fixture、静态快照或重建样例。这个标记说明上游数据的构造性质，不表示输入或 Query 是本项目临时编造的。43 道题都通过当前隐私检查并标记为可公开，其中 38 道记录为经过人工改写，5 道保留上游原始要求。

输入来源和 Golden 来源是两件事。上游存在 Golden 时，本项目保留或按其真值重建；上游只提供输入、任务要求或 Rubric 时，Golden 由制题者逐题阅读原始文件后，依据上游证据人工制作和检查，不能用统一脚本随机生成业务事实。Judge 必须从输入、Rubric 或冻结真值独立重算关键事实，避免只比较文件哈希或只接受参考解法的固定版式。

来源追踪仍有一个发布前风险：21 道 OfficeVal 题的 `source_uri` 仍指向 Hugging Face 的 `main` 路径。当前本地输入已经由 SHA256 固定，`provenance_id` 也记录了上游 commit，但正式 Release 前仍应把这些链接全部改成不可变 revision，保证第三方可以从同一个公开版本复核原文件。

任务包的全局结构、Schema、输入、Golden 和哈希校验已通过 `43/43`；现存 Golden 经各题 Judge 检查为 `43/43`，从空输出目录执行 `solve.sh -> tests/test.sh` 的 fresh replay 也是 `43/43`，空交付目录被拒绝 `43/43`。每题的来源文件、修改边界、真值、replay、负例与视觉检查记录位于自己的 `adaptation/` 和 `metadata/`。这些题仍是论文候选集；在非参考 Agent Rollout、Judge 梯度校准、泄漏分组和三段式切分完成前，不宣称为已冻结的论文 Release。

参考仓库审计见 [`datasets/reference-source-audit.md`](datasets/reference-source-audit.md)，42 个任务家族的候选映射与逐题状态见 [`datasets/task-candidate-map.md`](datasets/task-candidate-map.md)。

## Harbor 兼容性现状

“能够被 Harbor 识别”和“能够通过官方 `harbor run` 完整回放”是两个不同层次。当前 43 道候选题已经使用官方 Harbor 0.21.0 的 `Task.is_valid_dir()` 和 `Task(...)` 完成静态解析，目录和 `task.toml` 均可被识别；项目自己的公开题校验器也能通过 43 道题。但是，目前的 `solve.sh -> tests/test.sh` fresh replay 是在宿主机任务目录中直接执行，不等同于 Harbor 容器回放，因此不能据此宣称这些题已经全部 Harbor-ready。

| 检查项                                  | 当前结果 | 结论                                                                                 |
|-----------------------------------------|---------:|--------------------------------------------------------------------------------------|
| 官方 Harbor 任务目录识别                |    43/43 | 通过                                                                                 |
| 官方 Harbor `task.toml` 解析            |    43/43 | 通过                                                                                 |
| 项目级结构、输入、Golden 和哈希校验     |    43/43 | 通过                                                                                 |
| 宿主机 `solve.sh -> tests/test.sh` 回放 |    43/43 | 通过，但不是正式 `harbor run`                                                        |
| Harbor 容器内输入可见性                 |     0/43 | 未满足；输入仍位于任务根目录 `data/input_files/`，没有进入 `environment/` 构建上下文 |
| Harbor 标准 Reward 文件                 |     0/43 | 未满足；当前测试没有写 `/logs/verifier/reward.txt` 或 `reward.json`                  |
| 完整官方 Harbor 端到端合规              |     0/43 | 当前不能宣称通过                                                                     |

正式 Harbor 运行的两个阻断问题如下：

- Harbor 默认只以 `environment/` 作为环境构建上下文，Oracle 和 Verifier 运行时分别上传 `solution/` 与 `tests/`。当前 43 道题的解法都依赖任务根目录下的 `data/input_files/`，但这些文件没有被打入环境镜像或上传到容器，所以 Agent、参考解法和部分 Judge 在标准容器路径中看不到输入。
- Harbor Verifier 要求测试结束后生成 `/logs/verifier/reward.txt` 或 `/logs/verifier/reward.json`。当前 43 道题的 `tests/test.sh` 只调用 Judge 并把结果写到标准输出或题目产物目录，没有生成 Harbor Reward 文件，因此官方 Verifier 会报告找不到 Reward。

此外，43 道题仍使用已弃用但暂时兼容的 `allow_internet = false`，应迁移为 `network_mode = "no-network"`；当前 artifacts 配置也与 Harbor 默认收集的 `/logs/artifacts` 路径重叠，会产生警告。修复这些问题后，必须使用官方 `harbor run` 逐题执行 Oracle 正例、空交付和关键负例，保存完整 trial 日志，才能把状态从“静态格式兼容”升级为“Harbor 端到端合规”。

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

当前制题目录采用下面的结构。它保留来源审计、Golden 和 Judge 所需材料，但仍是作者侧候选题结构；正式发布为 Harbor 任务前，需要按上一节说明把 Agent 可见输入打入 `environment/`，并让测试生成标准 Reward 文件。

```text
tasks/<领域代码>/<任务编号>/
├── instruction.md
├── task.toml
├── rubric.json
├── environment/
│   └── Dockerfile
├── data/
│   └── input_files/
├── solution/
│   ├── solve.sh
│   └── <Golden 办公文件>
├── tests/
│   ├── judge.py
│   ├── test.sh
│   └── expectations.json      # 可选，声明确定性检查目标
└── metadata/
    ├── task.json              # 多轴分类与论文分析元数据
    └── provenance.json        # 来源、许可、隐私状态与输入哈希
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
/mnt/bn/liuzhou-hl-training/liuzhou/conda_envs/cowork-evolution-benchmark
```

常用检查如下：

```bash
环境路径=/mnt/bn/liuzhou-hl-training/liuzhou/conda_envs/cowork-evolution-benchmark

# 运行项目级公开题校验；它不等同于官方 harbor run
PYTHONPATH=src "$环境路径/bin/python" -m cowork_evo_bench.cli validate-public-tasks tasks

# 运行仓库单元测试
PYTHONPATH=src "$环境路径/bin/python" -m unittest discover -s tests/unit
```

## 当前阶段

当前仓库已经完成第一版模块骨架、协议 Schema、多轴分类法、Harbor 题目模板，以及 42 个任务家族的首轮真实上游选题。下一阶段从“补家族”转为“用真实 Agent 检验题目与 Judge”：

- 先修复 43 道题的 Harbor 输入打包、标准 Reward、网络配置和 artifacts 路径，并使用官方 `harbor run` 完成逐题容器回放。
- 给每道候选题补充真实 Agent 产生的中等样本、视觉缺陷负样本和损坏文件样本，校准确定性、视觉和语义 Judge。
- 使用至少 3 个最新 Agent Harness 做 pass@1 / pass@8 难度 Rollout，并剔除过易、不可解或 Judge 不稳定的题。
- 仅对 Rollout 显示覆盖不足或方差过大的家族定向增加第 2 道异质题，不为凑数量批量制造输入、Query 或 Golden。
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
