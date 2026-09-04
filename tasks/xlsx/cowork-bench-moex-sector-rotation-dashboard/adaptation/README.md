这道题逐题改造自 Cowork-Bench `moex-sector-rotation-dashboard`，不是本项目自编的股票分析题。

- 上游 Query、研究门户 HTML、MOEX mock 数据库、Golden 和确定性 Evaluator 都存在。
- 静态输入只投影原 Query 实际使用的 10 个收盘价和 5 份最新年度利润表，不从 Golden 反推市场数据。
- 上游数据是本地 mock 市场快照，不是可用于真实投资决策的实时数据，因此 provenance 标记 `synthetic=true`。
- 上游称它为 dashboard，但 Query 只规定三张表和键值 Summary，没有要求图表或交互控件；本改造不会额外编造隐藏视觉门槛。

逐题验证记录：

- 原始 Golden 经独立 Judge 从输入重新计算后得分 `1.0`。
- `solve.py` 从静态输入回放出的三张表，与原始 Golden 逐单元格完全一致，回放得分 `1.0`。
- 空白、缺文件、错误价格、错误评级、错误收益率、错误财务数、错误 Summary 和错误工作表顺序共 8 个负例均未通过，最高负例得分 `0.641026`。
- LibreOffice 能真实打开、重存和渲染工作簿；视觉抽检确认三张表的数据可读，但上游 Golden 没有做专业看板排版，`Performance` 表头拥挤并额外溢出一页。它可以作为数据联结与计算型 `dashboard` 候选，正式 Release 前仍需决定是否接受上游原始视觉质量，或改换有真正图表/交互要求的上游题。

## 难度档位调整记录（2026-08-31）

本题从 L3/advanced 调整为 L4/expert 的作者结构假设，依据是：市场价格、财报和研究门户三类输入汇入专业行业轮动分析看板，包含多源追溯与图表交付。 难度来自真实工作流中的异构输入、跨文件一致性、状态或专业分析链条，不是单次 Sonnet 5 得分。calibration_status=uncalibrated、review_status=pending_human_review 保持不变，最终需结合多模型结果、失败原因和人工复核。
