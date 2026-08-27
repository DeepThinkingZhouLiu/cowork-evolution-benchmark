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
