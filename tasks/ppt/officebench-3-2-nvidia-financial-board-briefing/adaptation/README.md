# OfficeBench 3-2 逐题适配记录

本题直接读取 OfficeBench `tasks/3-2/subtasks/0.json` 和完整的 10 页 `financial_report.pdf`。原 Query 要求先分析到 `report.xlsx`，再生成 `summary.docx`；原 evaluator 只检查两个文件存在，以及 Word 中含 `NVIDIA`、`financial` 两个关键词，无法验证任何财务事实。

本次适配保留原 PDF，不另编公司、季度或数值。制题者逐页审计 PDF 第 1–7 页后，将交付改成 6 页管理层简报：季度结论、原生 Revenue 图、GAAP 三期表、四业务收入图、GAAP/Non-GAAP 与现金流、Q2 Outlook。Golden 中的推导仅限于可由同比/环比直接支持的表述，并明确标记 `Derived reading`。

Golden 使用任务专属代码逐对象制作，这个代码只服务于本题，不读取其他题目，也不随机生成事实。Judge 将从冻结 PDF 独立提取文本并核对哈希，再检查原生图表、原生表格、GAAP/Non-GAAP 边界、分部、现金流、Outlook、来源页码和可编辑性。

## 实际验证

- Golden、ZIP 等价重封装和非语义性封面变体均为 1.0。Revenue 图错、GAAP 漏行/错值、分部图/增速错、Non-GAAP 错、FCF 错、Outlook 错、缺来源、投资建议、压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，输出 SHA256 稳定为 `875cce0e383842abeacb62b201dcc8caa0776a89a9366595c1a083792272e4cd`。单题结构、公开资产、任务元数据、来源和 Rubric Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `officebench32-oracle-nvidia-briefing-20260829`，trial 为 `officebench-3-2-nvidia-financial__Q8nnSgw`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- PDF 第 1–7 页和 Golden 6 页均已渲染并逐页目视。初版第 2 页图表标题与纵轴重叠，已人工发现、修复并重新完整回放；最终 `visual_qa.py` 返回空问题列表。
