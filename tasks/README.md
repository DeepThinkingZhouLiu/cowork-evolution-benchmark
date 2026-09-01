# Harbor 题目目录

本目录保存 CoworkEvoBench 逐题改造的公开候选题与统一制题模板。先前批量生成的 42 道合成题已经删除；这里不允许用通用生成脚本批量编造输入、Query 或 Golden。

## 目录责任

| 路径                    | 负责内容                                             | 不负责内容                                      |
|-------------------------|------------------------------------------------------|-------------------------------------------------|
| `templates/`            | Harbor Office 题目模板和字段示例                     | 计入分数或充当正式题                            |
| `ppt/`                  | 逐题验收后的演示文稿候选题                           | Excel 或跨应用任务                              |
| `docx/`                 | 逐题验收后的文字文档候选题                           | 未审查的敏感材料                                |
| `xlsx/`                 | 逐题验收后的电子表格候选题                           | 数据集切分和排行榜汇总                          |
| `cross_office/`         | 逐题验收后的多交付物或工作区候选题                   | Evolution Epoch 调度                            |

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

| 主领域         | 已完成候选题 | 家族覆盖 | 当前状态                                                        |
|----------------|-------------:|---------:|-----------------------------------------------------------------|
| `ppt`          |           60 |    12/12 | 已达到阶段目标并完成 60/60 官方 Harbor Oracle 回放                 |
| `docx`         |           14 |      8/8 | OfficeVal、LLM-DOCX、FORTE 与 Cowork-Bench 来源齐全                 |
| `xlsx`         |           14 |    14/14 | 全部表格家族已覆盖，包括 SpreadsheetBench-2 公式修复与 FORTE 三工作簿整合 |
| `cross_office` |           13 |      8/8 | 3 道新扩展题均完成双 Harbor Docker replay，其余继续逐题扩展         |
| **合计**       |      **101** | **42/42** | 全部任务包通过公开资产结构、输入、Golden 和 SHA256 校验 `101/101`     |

首轮候选映射见 [`../datasets/task-candidate-map.md`](../datasets/task-candidate-map.md)，PPT 与 Cross-Office 的后续逐题扩展分别见 [`../docs/ppt-task-intake-v0.1.md`](../docs/ppt-task-intake-v0.1.md) 和 [`../docs/cross-office-task-intake-v0.1.md`](../docs/cross-office-task-intake-v0.1.md)。101 道是当前候选集规模，不是已冻结的论文 Release 规模。

全库累计检查结果为：严格 Schema、目录、输入清单与 SHA256 校验 `101/101`；仓库内 Golden 经各题 Judge 验收 `101/101`；从空输出目录重新运行参考解法并交给同一 Judge 的 fresh replay `101/101`；空交付目录被拒绝 `101/101`。公开来源不等于原始生产数据，每题的真实来源性质以自己的 `metadata/provenance.json` 为准。

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
环境路径=REMOVED_LOCAL_PATH

PYTHONPATH=src "$环境路径/bin/python" -m cowork_evo_bench.cli validate-public-tasks tasks
```
