# CoworkEvoBench

> **面向真实办公任务的 Harness 自进化基准**
>
> 测量智能体能否从办公任务反馈中改进 Harness，并把学到的流程迁移到新的文件和跨办公工作流。

<p align="center">
  <img src="https://img.shields.io/badge/tasks-240-2563eb" alt="240 tasks" />
  <img src="https://img.shields.io/badge/modalities-4-0f766e" alt="4 modalities" />
  <img src="https://img.shields.io/badge/release-v0.5-7c3aed" alt="v0.5" />
  <img src="https://img.shields.io/badge/python-%3E%3D3.11-3776ab" alt="Python 3.11+" />
</p>

CoworkEvoBench 的核心对象是**可交付的办公文件**和**围绕模型的 Harness**。每道题固定输入、输出合同、运行环境和 Judge；演化方法只能修改允许范围内的 Harness，不能修改题目、评分规则、预算或密封测试集。

## 这个基准测什么

| 模态 | 任务数 | 典型交付物 | 主要能力 |
| --- | ---: | --- | --- |
| Presentation | 60 | PPTX / PDF | 编辑、重建、视觉设计、多源合成 |
| Document | 60 | DOCX | 内容编辑、原生对象重建、证据综合、受控审阅 |
| Spreadsheet | 60 | XLSX | 数据变换、公式、跨表依赖、对账和决策分析 |
| Cross-office | 60 | XLSX、DOCX、PPTX、EML、ICS 等 | 跨文件事实传播、一致性和端到端交付 |

任务来自 11 个公开来源，保留来源 revision、许可证、输入哈希和逐题改造记录。当前 v0.5 共有 59 个任务族、625 个输入文件、532 个必需输出文件，其中 337 个为原生 Office 文件。

## 为什么要把反馈、选择和最终测试分开

```text
固定 H₀
   │
   ├── feedback  90 题：提供轨迹、产物和逐项反馈
   ├── selection 30 题：只提供聚合信号，选择 Candidate
   └── final    120 题：冻结 Harness 后测试
                    ├── 60 题同域泛化
                    └── 60 题跨办公迁移（组合式 OOD）
```

这一区分回答的是科学问题：单一文件类型上的提升，是否真的形成了可以复用的程序性能力。最终任务、参考答案和 Judge 内部实现，在 Harness 冻结前不可访问。

## 评分原则

Judge 检查提交文件本身，而不是只检查运行日志：

- 内容、事实、数字和计算是否正确；
- Office 对象、公式和链接是否保持原生可编辑；
- 非目标内容、来源和隐私约束是否被保留；
- 多个交付文件是否互相一致；
- 版式、层级、间距和视觉质量是否满足任务要求。

能稳定程序化判断的内容使用确定性检查。PPTX/PDF 视觉题额外使用任务专属 rubric，并按 `70% deterministic + 30% visual` 合并。每题都在隔离 Docker 环境中运行，网络关闭，参考答案、Rubric 和 Judge 对 Agent 隐藏。

## v0.5 发布状态

当前冻结版本和历史 Release 入口位于 [`benchmarks/README.md`](benchmarks/README.md)。主线的 v0.5 细节随对应 Release 目录发布。四种模态各 60 题，难度按 `12 / 18 / 18 / 12` 划分为 Easy、Medium、Hard、Expert。

| 检查项 | 状态 |
| --- | ---: |
| 任务目录、输入、Golden、哈希 | 240 / 240 |
| 四模态 Oracle 回放 | 240 / 240 |
| `feedback / selection / final` 分区锁定 | 已锁定 |
| 无网络和单一 `/app/output` 采集 | 240 / 240 |

“静态校验通过”和“官方 `harbor run` 完整回放通过”是不同状态。逐题回放证据保存在任务的 `adaptation/`、`metadata/` 和 `docs/` 中。

## 仓库结构

```text
.
├── tasks/             # 题目包：instruction、输入、环境、Golden、Rubric、Judge
├── benchmarks/        # 冻结 Release、dataset lock、split 和 leakage group
├── src/                # Python 校验、评测、Judge、适配器和演化接口
├── schemas/            # 跨模块 JSON Schema
├── configs/            # Benchmark、评测、Harness 和 taxonomy 配置
├── harnesses/          # 可复现的初始 Harness 定义
├── evolvers/           # Evolution 接入协议和参考实现
├── mutation-policies/  # Candidate Harness 的路径白名单
├── datasets/           # 来源、许可、哈希和审计清单
├── docs/               # 架构、数据治理、任务规范和实验记录
└── tests/              # 单元、集成和协议测试
```

单题包的标准结构见 [`tasks/README.md`](tasks/README.md)。模块职责和数据边界见 [`AGENTS.md`](AGENTS.md)、[`docs/architecture.md`](docs/architecture.md) 和 [`docs/evaluation-spec.md`](docs/evaluation-spec.md)。

## 本地安装与验证

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install -e .

# 单元测试
PYTHONPATH=src python -m unittest discover -s tests/unit

# 公开任务结构、输入和资产校验
PYTHONPATH=src python -m cowork_evo_bench.cli validate-public-tasks tasks

# 检查冻结 Release 的内容锁和三段式划分
PYTHONPATH=src python tools/freeze_benchmark_release.py --check
```

上述命令验证协议和公开资产；它们不替代需要 Docker 和 Harbor 的完整 Agent rollout。

## 研究与开发约定

- 题目、Rubric、Judge、预算和数据划分是评测控制面，Evolution 不得修改。
- `feedback`、`selection`、`final` 必须互斥；同源模板和近重复任务按 leakage group 一起切分。
- 私有办公文件、密封题目、运行产物和模型凭据不提交 Git；大型本地运行目录使用 `.gitignore` 排除。
- 修改任务格式、评分公式或 Release 时，必须同步更新 Schema、验证器、版本号和审计记录。
- `main` 是冻结主线；`lz-dev` 和 `cxy/harbor-migration-43` 是当前保留的开发分支。

## 重要文档

- [Release 与版本索引](benchmarks/README.md)
- [任务分类](docs/task-taxonomy.md)
- [任务规范](docs/task-spec.md)
- [评测协议](docs/evaluation-spec.md)
- [Evolution 协议](docs/evolution-protocol.md)
- [数据治理](docs/data-governance.md)
- [Judge 与视觉评测](docs/evaluation-spec.md)
- [任务包索引](tasks/README.md)

## 许可与来源

每道题在 `metadata/provenance.json` 和 `adaptation/` 中记录来源、许可证、revision、输入哈希和改造边界。仓库中的派生题包遵循相应上游许可证；使用或再发布前请同时检查题包内的许可证文件和第三方声明。
