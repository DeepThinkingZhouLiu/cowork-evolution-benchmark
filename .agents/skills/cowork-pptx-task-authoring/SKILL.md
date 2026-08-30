---
name: cowork-pptx-task-authoring
description: 逐题选择、改造、制作并验证 CoworkEvoBench 的真实来源 PPTX Harbor 任务。新增、修订或审计 tasks/ppt 下任务时使用；不用于批量合成 Mock 题、通用脚本造题或普通演示文稿写作。
---

# CoworkEvoBench PPTX 逐题构造

本 Skill 只负责把一个许可清楚、内容完整的上游 Seed 改造成一个可复现、可评分、可回放的 PPTX 任务。模型负责选题、理解材料、叙事、视觉、Rubric 和错误样例判断；代码只负责下载、哈希、解析、渲染、打包和单题的确定性实现。

## 开始前

1. 在仓库根目录阅读 `AGENTS.md`、`docs/task-spec.md`、`docs/task-taxonomy.md` 和 `docs/supplier-requirement-adaptation.md`。
2. 从 `../../../003-Reference/002-Reference-Code/sources.lock.json` 确认来源 commit 与许可。路径无法解析时，先定位项目的 `003-Reference/002-Reference-Code`，不要猜测来源。
3. 使用 `presentation-skill` 制作或编辑 Golden；读取与当前任务对应的创建、编辑、数据或模板参考，并执行其渲染与视觉 QA。
4. 查看工作树状态。一次只改一题及其直接依赖；不得覆盖用户或其他题目的未提交改动。

## 一题一上下文

在创建任务目录前，必须把以下内容实际读入当前模型上下文，不得只看文件名、搜索摘要或自动生成的统计：

- 上游原始 Query、任务定义和 evaluator/checker；
- 每个输入文件的可读内容；PPTX、PDF、DOCX、XLSX 和图片必须分别渲染或展开检查；
- 许可、NOTICE、commit、输入 SHA256 和上游 synthetic/human-rewritten 状态；
- 计划保留的真实意图、需要澄清的歧义和明确不沿用的上游缺陷。

把这次阅读证据写入该题的 `adaptation/README.md`。未完成上述上下文包时，不得写 Query、Golden、Rubric 或 Judge。

## 来源分级

- `direct`：原 Query、输入和验收依据可直接追溯并允许再分发，可逐题改造。
- `composite`：业务状态来自一个许可清楚的来源，PPT 模板或素材来自另一个许可清楚的来源。必须分别记录来源、哈希和各自贡献，不能声称是上游原题。
- `design_only`：许可、加密、gated access、凭据或运行环境不允许公开复用。只可抽象任务结构，不得复制 Query、输入、答案、隐藏状态或 evaluator 细节。

遇到许可不清、临时链接、输入缺失、隐私风险、答案泄漏或无法确定性验收时，淘汰 Seed，不得自行补故事冒充真实来源。

## 单题制作

1. 写一页适配契约：主交付物必须是 PPTX；列出角色、受众、输入、需要做出的业务判断、允许的辅助应用状态、输出名和不可妥协约束。
2. 保留上游事实和意图。必要的澄清只能消除歧义，不能为了方便 Judge 改成模板填空题。
3. 在 `data/input_files/` 保存 Agent 可见输入，在 `environment/` 的构建上下文中以只读方式复制到固定路径；不得依赖网络或宿主机绝对路径。
4. 按任务内容逐页设计 Golden。所有正文、图表、表格、流程和关键图形应原生可编辑；截图只可作为有来源的图片素材，不能压平整页代替作答。
5. 参考解法可以调用确定性渲染器，但每题的内容、结构、事实选择、图表映射和版式决定必须保存在该题自己的源文件中。禁止用循环批量生成多题 Query、输入、Golden 或 Rubric。
6. 同步设计 `rubric.json` 与 `tests/judge.py`。Judge 评分候选产物，不比较 Golden 文件哈希，也不依赖解题脚本留下的秘密标记。
7. 若题目带邮件、日历、表格或其他模拟应用状态，PPTX 仍是主交付物；状态断言作为独立 Rubric 项收集，不能让 Agent 直接写评分证据。

详细门槛见 [references/quality-gates.md](references/quality-gates.md)。需要设计或校准难度时再读 [references/difficulty-calibration.md](references/difficulty-calibration.md)。

## 完成定义

只有同时满足以下条件，一题才可从 Seed 变为候选题：

- Harbor 包文件齐全，Schema、输入哈希和来源记录通过；
- Golden 可打开、可编辑、内容正确，并完成逐页原图视觉检查；
- Oracle `harbor run` 得满分，空交付、损坏文件、未修改原件和关键错误负例均被拒；
- 至少一个中等质量样例获得介于空交付与 Golden 之间的分数；
- 同一产物连续评分五次的标准差小于 0.05；
- `adaptation/README.md` 记录实际命令、结果、渲染页数、已看页面和未关闭风险。

难度字段在真实 Agent 回放前保持 `unassigned`。设计目标只能写成“候选 L1/L2/L3/L4”，不得冒充实测标签。
