# 总体架构

本仓库定位为公开、可复现的 Office Benchmark。它固定任务合同、运行环境、输出要求、可信 Judge、评分策略和 Release 清单，用于比较不同模型与 Agent Harness 在真实办公任务上的表现。

## 数据与评测链路

```text
来源与许可证清单
        ↓
任务包（instruction / inputs / Dockerfile / Golden / Rubric / Judge）
        ↓
冻结 Release（split / 难度 / 内容锁 / 泄漏组件）
        ↓
隔离 Harbor 容器中的 Agent Harness
        ↓
可信 Judge（结构、事实、公式、跨文件、一致性、可选视觉评分）
        ↓
标准单题结果 → 聚合指标与报告
```

- `datasets/` 保存来源、许可证、隐私、转换和哈希清单。
- `tasks/` 是题目唯一来源，每题都是可单独回放的 Harbor 包。
- `benchmarks/` 冻结评测集合和访问边界，不复制题目文件。
- `judging/` 只负责检查候选交付物并输出逐条证据。
- `evaluation/` 只消费标准结果，计算平均分、通过率、配对差异和成本指标。
- `adapters/` 将 Harbor、外部 Harness 和 VLM Provider 接入统一接口，不把具体产品逻辑写入任务和 Judge。

## 可信边界

```text
可信评测侧
├── Release 清单与内容锁
├── 隐藏 Golden、Rubric 和 Judge
├── 运行预算与基础设施状态
├── 逐条判据和最终聚合
└── 结果统计与审计记录

受测侧
├── Agent Harness
├── 任务可见的 instruction 与输入
└── 输出目录中的候选文件
```

受测 Agent 不能读取隐藏 Golden、Judge 实现、视觉 Rubric、Provider 凭据或其他任务的私有状态。候选文件、运行日志和 Agent 自报分数只能作为输入证据，最终结果由可信评测侧重新计算。

## 依赖方向

- `tasks` 可以依赖 `schemas` 和任务专属工具，不能依赖运行结果或本机路径。
- `judging` 可以读取题目隐藏侧内容和候选目录，不能读取 Agent 自报分数作为事实。
- `evaluation` 只能消费标准单题结果，不能直接调用具体模型。
- `adapters` 可以调用外部 Runner 或 Provider，但不能修改题目、Rubric、Judge、Release 或分母。
- 报告由上层实验脚本生成，不能重新定义单题得分。

## 运行时目录

运行状态、渲染图片和 Provider 响应只写入 Git 忽略的 `.coworkevo/`、`artifacts/` 或外部结果目录。它们不是公开任务包的一部分，也不能作为输入提交到 Release。
