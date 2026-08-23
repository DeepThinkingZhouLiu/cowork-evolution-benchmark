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

## 目录结构与模块职责

```text
.
├── docs/                       # 研究协议、架构、数据治理和题目制作规范
├── schemas/                    # 跨模块交换对象的机器可校验契约
├── configs/                    # Benchmark、Harness、评测和 Evolution 的实验配置
├── datasets/                   # 数据来源清单、版本锁和整理记录，不保存私有大文件
├── tasks/                      # Harbor 兼容题目包、公开样例和制题模板
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

正式题目采用 Harbor 目录结构：

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
│   └── test.sh
└── metadata/
    └── provenance.json
```

Agent 运行时只能看到 `instruction.md`、允许公开的 `task.toml` 字段和 `data/input_files/`。`solution/`、完整 `rubric.json`、`tests/` 以及 Judge 凭据属于隐藏评测侧。

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

## Evolution 赛道

- `prompt-only`：只允许修改系统提示词、`AGENTS.md` 和少量示例。
- `skill-memory`：允许修改技能、经验库、检索索引和加载规则。
- `full-harness`：允许修改完整做题 Harness，但仍不能触碰可信控制面。

路径白名单必须同时通过运行时只读挂载和结束后的 Diff 复核落实，不能只在提示词中声明。

## 私有数据管理

真实用户文件、企业材料和密封测试题不提交到 Git。`datasets/` 仅保存来源、许可、下载日期、文件哈希、隐私状态、转换版本和发布权限。代码通过独立的数据根目录读取资产，并使用 `dataset.lock` 固定精确版本。

原始 273 条 Paper2Any 数据只是候选池。进入正式 Benchmark 前仍需完成隐私审查、任务有效性筛选、内容哈希去重、Golden 补全、Rubric/Judge 制作、Judge 校准和分组切分。

## 当前阶段

当前仓库提供第一版模块骨架、协议 Schema、Harbor 题目模板、基础结构校验和评测数据模型。完整 Harbor Runner、办公文件 Judge、Evolution Epoch Controller 和正式 Benchmark Release 将按端到端样例逐步实现。

第一阶段优先跑通：

```text
一条真实 PPT 题
-> Harbor 环境运行
-> 生成 PPTX
-> 确定性检查与视觉 Judge
-> 标准单题结果
-> H0/Candidate 配对评测
```
