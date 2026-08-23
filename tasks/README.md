# Harbor 题目目录

本目录保存 CoworkEvoBench 的公开端到端种子题与统一制题模板。当前 42 题共有 95 个输入文件和 59 个 Golden 交付物，均包含可打开输入、可执行参考解法、至少 6 条 Rubric、确定性 Judge、研究元数据和来源哈希；它们用于开发、回归和 Judge 校准，不等同于论文密封测试集。

## 目录责任

| 路径                    | 负责内容                                             | 不负责内容                                      |
|-------------------------|------------------------------------------------------|-------------------------------------------------|
| `templates/`            | Harbor Office 题目模板和字段示例                     | 计入分数或充当正式题                            |
| `ppt/`                  | 12 个演示文稿任务家族的公开种子题                    | Excel 或跨应用任务                              |
| `docx/`                 | 8 个文字文档任务家族的公开种子题                     | 采购合同等真实敏感材料                          |
| `xlsx/`                 | 14 个电子表格任务家族的公开种子题                    | 数据集切分和排行榜汇总                          |
| `cross_office/`         | 8 个多交付物、跨应用或工作区任务家族的公开种子题     | Evolution Epoch 调度                            |
| `catalog.jsonl`         | 逐题机器可读研究元数据                               | 保存 Golden 内容或 Judge 凭据                   |
| `catalog-summary.json`  | 领域、难度和能力覆盖统计                             | 论文最终实验结果                                |
| `CATALOG.md`            | 方便人工检查的逐题目录                               | 作为唯一真相源                                  |

## 单题责任边界

```text
<领域>/<任务编号>/
├── instruction.md              # Agent 可见的用户任务
├── task.toml                   # Harbor 运行资源、离线设置与公开元数据
├── rubric.json                 # 隐藏侧原子评分标准
├── environment/Dockerfile      # 可复现运行环境
├── data/input_files/           # Agent 可见的真实输入文件
├── solution/
│   ├── solve.py                # 参考解法实现
│   ├── solve.sh                # 一键生成 Golden 的入口
│   └── <Golden 文件>           # 隐藏侧参考交付物
├── tests/
│   ├── judge.py                # 候选产物评分实现
│   ├── test.sh                 # 一键 Judge 入口
│   └── expectations.json       # 可选的确定性期望
└── metadata/
    ├── task.json               # 多轴分类与论文分析标签
    └── provenance.json         # 来源、许可、隐私和输入 SHA256
```

Agent 只能读取任务指令、公开配置和 `data/input_files/`。`solution/`、完整 Rubric、测试实现和 Judge 凭据必须由可信评测侧挂载，不能出现在 Agent 工作目录中。

## 当前覆盖

| 主领域         | 家族数 | 每家族题数 | 当前用途                       |
|----------------|-------:|-----------:|--------------------------------|
| `ppt`          |     12 |          1 | 演示生成、编辑、打磨与审计     |
| `docx`         |      8 |          1 | 长文生成、修订、格式与抽取     |
| `xlsx`         |     14 |          1 | 数据、公式、模型、图表与审计   |
| `cross_office` |      8 |          1 | 多文件一致性与工作区编排       |
| 合计           |     42 |          1 | 公开开发种子                   |

逐题列表由 `tools/task_generation/build_catalog.py` 生成，见 [`CATALOG.md`](CATALOG.md)。

## 进入正式发布的门槛

公开种子或候选数据只有同时满足以下条件，才能进入某个冻结的 Benchmark Release：

- 输入、环境、Golden、Rubric 和 Judge 全部可复现，且不存在临时外链。
- 每题至少 6 条客观、原子、完整、可区分的 Rubric，并包含必要的负向惩罚。
- Golden 得到高分；中等缺陷、明显负例和损坏文件能够被正确区分。
- 确定性 Judge 重复运行一致，模型 Judge 的重复标准差满足冻结阈值。
- 输入来源、许可、隐私状态、下载时间和 SHA256 完整可审计。
- 至少 3 个目标 Agent Harness 完成难度 Rollout，过易、不可解和歧义题被剔除。
- 来源相同、文件相同、模板近似和语义近重复任务进入同一泄漏组。
- `feedback`、`selection`、`final` 按泄漏组切分，最终集合在 Harness 冻结前保持密封。

## 本地检查

```bash
环境路径=/mnt/bn/liuzhou-hl-training/liuzhou/conda_envs/cowork-evolution-benchmark

PYTHONPATH=src "$环境路径/bin/python" -m cowork_evo_bench.cli validate-public-tasks tasks
"$环境路径/bin/python" tools/task_generation/build_catalog.py
"$环境路径/bin/python" tools/task_generation/validate_roundtrips.py
```
