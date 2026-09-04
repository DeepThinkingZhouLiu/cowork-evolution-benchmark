# 单题改造说明

本题逐文件改造自 Cowork-Bench `sf-department-budget-ppt`，固定上游 commit 为 `d943e75bc0fc8e3b27141979300cd8cbcd1e890d`，上游整仓采用 Apache-2.0 许可。

## 保留的上游材料

- `upstream-task.md`：上游俄文 Query，字节不变。

- `upstream-agent-system-prompt.md`：上游 Agent 系统提示，字节不变。

- `upstream-task-config.json`：上游工具与 MCP 声明，字节不变。

- `upstream-preprocess.py`：上游预处理脚本，字节不变。

- `upstream-evaluator.py`：上游评分脚本，字节不变。

- `upstream-input-department_budgets_fy2026.pdf`：上游原始 PDF 输入，字节不变。

- `upstream-groundtruth-budget_vs_actuals.xlsx` 和 `upstream-groundtruth-budget_vs_actuals.pptx`：上游原始 Golden，字节不变，只用于出处和缺陷审计。

- `LICENSE-Cowork-Bench`：上游完整 Apache-2.0 许可。

`source-manifest.json` 记录上述文件和上游 `db/init.sql.gz` 的大小与 SHA256。

## 动态表的离线化

原 Query 要求连接 ClickHouse 逻辑表 `HR_ANALYTICS.PUBLIC.EMPLOYEES`。该动态表的固定数据来自同一 Cowork-Bench commit 的 `db/init.sql.gz`。`extract_snapshot.py` 在校验整个 gzip 哈希后，以二进制行模式定位指定 `COPY` 声明，并原样写出从 `COPY` 声明到 `\.` 终止行的连续解压字节切片。

任务输入 `HR_ANALYTICS__PUBLIC__EMPLOYEES.copy.sql` 保留 18 列、50000 条员工记录，没有投影字段、抽样、重排、修改数值或新增记录。输出 SHA256 为 `c1dfba6c1b0c55b03e47cb46d093202af2befb9593bc7f4cf4796a9781844993`，因此可以逐行与上游 SQL dump 审计。

这些 HR 记录是 Cowork-Bench 的公开模拟企业数据库种子，不是现实公司员工数据，因此来源元数据中 `synthetic` 保持为真。“真实上游输入”表示任务使用上游公开 benchmark 实际发布的原始输入，不表示数据来自现实企业。

## 人工改写边界

新 `instruction.md` 是上游俄文 Query 的中文直译与离线化说明，保留了 CFO 角色、季度管理层会议场景、所有计算口径、部门名、页面顺序、工作表结构和交付文件名。没有新编业务记录、新角色、新业务结论或新 Query 意图。

为了把原题转成可比较的离线 Benchmark，只新增了三类非业务约束：动态 ClickHouse 换成同源 COPY 快照、PPTX 要求 16:9 且保持原生可编辑对象、禁止输入无法支持的额外断言。

## 上游 Golden 缺陷与修正

人工解析、LibreOffice 渲染并逐页审查上游 Golden 后，确认了以下缺陷：

- 上游 PPTX 是 4:3，不是当前管理层汇报常用的 16:9。

- 上游首页写成 `Budget vs Actuals Report - FY2026` 和 `Department Budget Analysis / March 2026`，没有按 Query 使用指定标题 `FY2026 Budget vs Actuals` 和副标题 `Quarterly Leadership Review - March 2026`。

- 上游部门页漏掉 Query 明确要求的 `Headcount` 和 `Avg Salary`。

- 上游部门页的 `Variance Pct` 只显示 1 位小数，Query 要求 2 位。

- 上游页面是默认空白模板和简单项目符号，且上游 Evaluator 对部门页的人数、平均薪资、百分比以及按页数字对应都没有严格校验。

新 Golden 不从旧 Golden 拷贝数字。`solution/solve.py` 每次从当前 PDF 和 COPY 快照重算真值，再交付 Excel 与可编辑 PPTX。Judge 也使用独立实现从输入重算，不与新或旧 Golden 做字节比较。

## 解决方案构建工具

PPTX 根据 `solution/deck_workspace/outline.json` 构建。为使回放不依赖机器全局 Skill，题包在 `solution/vendor/presentation_skill/` 保留 presentation-skill 0.9.0 的必需渲染脚本、精确 SHA256 和 MIT LICENSE。这些脚本只用于本题参考解的可重建性，不是批量制题器。

## 视觉评分适配（visual-v1）

- 模型逐题读取并检查了 `instruction.md`、`task.toml`、`rubric.json`、任务专属 `tests/judge.py`、`tests/expectations.json`、上游适配记录和参考解工作区。输入 PDF 的 1 页预算表已按原尺寸渲染查看；9.5 MB 的 SQL COPY 快照已完整解析 50000 条、18 列员工记录，并按七个俄文部门重新汇总人数、薪资总额和平均薪资，而不是只查看文件名或抽样行。
- Golden `budget_vs_actuals.pptx` 共 10 页，已先查看接触表，再逐页查看原尺寸图：深色 CFO 封面、四项全局总览、七个俄文部门的计划/实际左右对照页，以及最大/最小结余总结页；配套 `budget_vs_actuals.xlsx` 也已核对工作表结构和跨文件真值。
- 本题允许不同的管理层汇报设计，只要求同一批源数据、指定页序与清晰的预算叙事，因此采用 `reference_free`。Golden 仅用于制定视觉标准和稳定性校准，不作为 VLM 的唯一版式参考。
- `V001` 针对 CFO 封面层级，`V002` 针对总览页长金额的表格扫读，`V003` 针对七个部门页的计划/实际跨页视觉映射，`V004` 针对六项指标在投影尺度下的可读性，`V005` 针对总结页两个极值与“全部低于预算”的消歧，`V006` 针对 10 页叙事节奏和交付一致性；它们只补充确定性 Judge 无法判断的版式、阅读顺序与专业表达。
- Golden 的已知视觉风险是第 3–9 页版式高度重复且留白较大，以及第 10 页 `Most Over Budget` 在所有部门实际均低于预算时容易脱离底部状态带被误读。校准时不修改 Golden，也不为制造高分而放宽这两项标准。
- Golden 已用固定 Codex CLI 0.148.0、ZCloud 与 `gpt-5.6-terra` 连续真实评分 5 次，visual reward 为 `1.000000 / 0.984615 / 0.984615 / 1.000000 / 0.984615`，均值 `0.990769`、总体标准差 `0.007537`（小于 0.05）；任务专属确定性 Judge 五次均为 `1.0`，70/30 综合分均值为 `0.997231`。
- 任务特定中等负例只把第 10 页右侧 `Most Over Budget` 标题和 Операции 明细整体左移 2.45 英寸，使其与 НИОКР 最大结余列明显叠压；底部“All 7 departments are under budget”仍保留，PPTX/XLSX 均完整可打开，确定性 Judge 仍为 `1.0`。真实 VLM visual reward 降至 `0.743590`，70/30 综合分为 `0.923077`，严格低于 Golden 均值且高于空白或损坏交付。
- 评分证据保存在 `.coworkevo/vlm-calibration/ppt/cowork-bench-fy2026-leadership-budget-review/`；本地校准产物不提交仓库。
