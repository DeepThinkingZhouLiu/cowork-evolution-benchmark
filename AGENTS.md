# Benchmark 仓库协作规范

## 核心边界

- 本仓库是公开、可复现的 Office Benchmark，核心是 240 个 Harbor 任务及其可信评测合同。
- 任务、输入、Golden、Rubric、Judge、数据划分和 Release 锁属于评测控制面；被测 Agent 只能看到任务声明的输入和输出合同。
- 结果必须由可信 Judge 重新计算。Agent 自报的成功、分数和运行日志不能直接成为最终指标。
- Provider 凭据、私有办公文件、密封运行产物和本机绝对路径不得提交到 Git。

## 目录职责

- `tasks/`：单题 Harbor 包，包含 instruction、运行环境、输入、Golden、Rubric、Judge 和来源记录。
- `benchmarks/`：冻结的 Release、split、难度标签、内容锁和泄漏组件。
- `src/cowork_evo_bench/judging/`：确定性 Judge、渲染和可选 VLM 评分。
- `src/cowork_evo_bench/evaluation/`：标准结果模型、聚合指标和配对比较工具。
- `src/cowork_evo_bench/adapters/`：Harbor、外部 Agent Harness 和评分 Provider 的适配接口。
- `src/cowork_evo_bench/tasks/`：任务加载、资产和结构校验。
- `schemas/`：公开任务、Rubric、Judge、结果和来源 Schema。
- `configs/`：评测策略、Judge/VLM 和任务分类配置，不保存真实凭据。
- `datasets/`：来源、许可证、隐私、哈希和转换清单，不保存大型私有数据。
- `docs/`：架构、任务格式、评测、数据治理和逐题审计记录。
- `tools/`：Release 冻结、校验、重评分和结果汇总脚本。

## 任务修改要求

- 新增或修改任务时，先完整读取 `docs/task-spec.md`、`docs/task-taxonomy.md` 和对应任务构造 Skill。
- 一次只处理一题及其直接依赖；必须实际查看输入、来源许可、Golden 和 Judge，不得用通用脚本批量伪造任务。
- 任务必须离线可回放，不能依赖宿主机绝对路径、未声明网络服务或隐藏的本地文件。
- Judge 必须从候选产物独立重算事实、结构、公式、跨文件一致性和输出门槛；可以使用题目专属视觉 Rubric，但不能比较 Golden 字节或内部对象 ID。
- 每条 Rubric 都要有可观察证据、权重和容差；确定性可判断的内容不得交给模型主观打分。
- 每题必须运行 Oracle、空交付、损坏文件、关键事实错误、不完整和中等质量样例；记录命令、分数和剩余风险。

## Release 与评测

- `tasks/` 是题目唯一来源；Release 目录只保存清单、哈希、split 和泄漏组件，不复制题目。
- Release 发布后不可原地改写。任务、Judge、Rubric、阈值或 split 变化必须生成新的版本。
- Feedback、Selection、Final 是评测协议中的访问范围，分别用于逐题诊断、聚合比较和最终报告。
- 四个模块采用各自的结构难度定义，但在同一冻结 Panel、模型、预算、重复次数和最终分数协议下，分别在模块内部排序。
- 基础设施失败、渲染失败和 Provider 失败必须单独标记并重跑，不得静默记为 Agent 零分。

## 验证

修改后至少运行相关单元测试、公开任务校验和当前 Release 检查：

```bash
PYTHONPATH=src python -m unittest discover -s tests/unit -p 'test_*.py'
PYTHONPATH=src python -m cowork_evo_bench.cli validate-public-tasks tasks
PYTHONPATH=src python tools/freeze_benchmark_release.py --check
```

代码和文档默认使用简体中文；提交前运行 `git diff --check`，并确认没有凭据、绝对路径、缓存和运行产物进入 Git。
