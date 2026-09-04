# 总体架构

## 核心判断

CoworkEvoBench 首先是题目与评测的 Benchmark，其次才是 Evolution 实验平台。架构必须保证题目和裁判可以独立演进，也可以在不改变科学问题的情况下替换 Agent Harness 或 Evolution 算法。

## 四层结构

```text
资产层：datasets -> tasks -> benchmarks
评测层：judging -> evaluation -> reporting
被测层：harnesses + evolvers
控制层：controller + adapters + mutation-policies
```

- 资产层回答“评什么”，拥有题目内容、版本、切分和泄漏组。
- 评测层回答“怎样算分”，拥有逐项判据、标准结果、配对统计和报告。
- 被测层回答“谁来做题、谁来改进”，只能通过公开接口参与实验。
- 控制层回答“怎样安全地运行实验”，负责隔离、预算、谱系和可复现性。

## 可信边界

```text
可信控制面
├── Benchmark Manifest
├── final 密封题目
├── Golden、Rubric 与 Judge
├── 预算与资源计量
├── Candidate Diff 校验
└── 配对评测与晋升规则

非可信被测面
├── 初始 Harness 的运行副本
├── Candidate Harness
├── Evolution 算法
└── Agent 生成的文件、日志和自报指标
```

非可信被测面提交的分数、成功声明和修改报告只能作为输入证据，不能直接成为最终结果。最终结果必须由可信评测层重新计算。

## 依赖方向

- `tasks` 可以依赖 `schemas`，不能依赖 `evolution`。
- `judging` 可以读取题目隐藏侧内容，不能读取 Candidate 的内部自评分。
- `evaluation` 只能消费标准单题结果，不能直接调用具体模型。
- `controller` 可以调用 Adapter，但不能把具体 Harbor 或 Agent 产品逻辑写入核心流程。
- `evolvers` 可以依赖公开协议，不能导入 Controller、Judge 或密封数据模块。
- `reporting` 只能读取不可变结果和运行账本，不能重新定义分数。

Codex-VLM 混合评分的可信边界、模块责任和失败语义见
[`codex-vlm-judging.md`](codex-vlm-judging.md)。模型调用只发生在可信评测侧，
不得把 Provider 凭据、隐藏 Golden 或视觉 Rubric 挂载到受测 Harness 容器。

## 运行时目录

所有实验状态写入被 Git 忽略的 `.coworkevo/`：

```text
.coworkevo/
├── sources/<harness-id>/
├── runs/<run-id>/
│   ├── run-manifest.json
│   ├── baseline/
│   ├── epochs/<epoch-id>/
│   │   ├── feedback/
│   │   ├── candidates/
│   │   ├── selection/
│   │   └── decision.json
│   ├── final/
│   └── evolution-ledger.jsonl
├── artifacts/
└── reports/
```

Baseline 与 Candidate 必须是独立工作副本。Updater 只能看到 Candidate，不能看到 Git 控制信息、最终测试题、Judge 凭据或其他 Candidate 的私有状态。
