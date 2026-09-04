# OfficeBench 3-2 逐题适配记录

本题直接读取 OfficeBench `tasks/3-2/subtasks/0.json` 和完整的 10 页 `financial_report.pdf`。原 Query 要求先分析到 `report.xlsx`，再生成 `summary.docx`；原 evaluator 只检查两个文件存在，以及 Word 中含 `NVIDIA`、`financial` 两个关键词，无法验证任何财务事实。

本次适配保留原 PDF，不另编公司、季度或数值。制题者逐页审计 PDF 第 1–7 页后，将交付改成 6 页管理层简报：季度结论、原生 Revenue 图、GAAP 三期表、四业务收入图、GAAP/Non-GAAP 与现金流、Q2 Outlook。Golden 中的推导仅限于可由同比/环比直接支持的表述，并明确标记 `Derived reading`。

Golden 使用任务专属代码逐对象制作，这个代码只服务于本题，不读取其他题目，也不随机生成事实。Judge 将从冻结 PDF 独立提取文本并核对哈希，再检查原生图表、原生表格、GAAP/Non-GAAP 边界、分部、现金流、Outlook、来源页码和可编辑性。

## 实际验证

- Golden、ZIP 等价重封装和非语义性封面变体均为 1.0。Revenue 图错、GAAP 漏行/错值、分部图/增速错、Non-GAAP 错、FCF 错、Outlook 错、缺来源、投资建议、压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，输出 SHA256 稳定为 `875cce0e383842abeacb62b201dcc8caa0776a89a9366595c1a083792272e4cd`。单题结构、公开资产、任务元数据、来源和 Rubric Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `officebench32-oracle-nvidia-briefing-20260829`，trial 为 `officebench-3-2-nvidia-financial__Q8nnSgw`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- PDF 第 1–7 页和 Golden 6 页均已渲染并逐页目视。初版第 2 页图表标题与纵轴重叠，已人工发现、修复并重新完整回放；最终 `visual_qa.py` 返回空问题列表。

## 视觉评分适配（visual-v1）

- 模型重新逐题读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、上游子任务、许可/NOTICE、来源与输入清单；10 页输入 PDF 已提取全文并查看全页接触表，其中业务事实集中在第 1–7 页，后 3 页为网页导航与版权尾页。
- Golden 的 6 页 PPTX 已展开原生图表/表格及全部文本，先检查 2 列接触表，再逐页查看 1920×1080 原图。页面依次承担季度结论、Revenue 趋势、GAAP 明细、平台收入、GAAP/Non-GAAP 与现金流、Q2 指引，所有页均保留来源行。
- 本题从财报材料开放式生成管理层简报，专业答案可以采用不同主题、图表和叙事布局，因此使用 `reference_free`；Golden 只用于制定本题 Rubric 和校准，不要求候选复制其几何结构。
- `V001` 聚焦封面四项关键数字与管理层边界，`V002` 聚焦趋势图和 GAAP 指标卡，`V003` 聚焦六行三期 GAAP 表，`V004` 聚焦业务规模图与四张业务语境卡，`V005` 聚焦 GAAP/Non-GAAP 和现金转换，`V006` 聚焦前瞻指引与行动区，`V007` 检查六页叙事、主题与来源系统；事实准确性和原生可编辑性继续由确定性 Judge 负责。
- Golden 的可见风险是封面大标题末行与管理层副标题距离偏近、第 4 页长业务标签和两个小规模柱较难处理、第 6 页行动区文字密度较高；VLM 必须按投影可读性如实评分，不能因其为 Golden 自动给满分。
- Golden 经 `gpt-5.6-terra` 连续五次真实视觉评分为 `1.0000 / 1.0000 / 1.0000 / 0.9847 / 1.0000`，均值 `0.9969`、总体标准差 `0.0061`，满足 `< 0.05` 的稳定性要求。
- 中等视觉负例仅对本题 Golden 做坐标回归：第 3 页 GAAP 表整体下沉并覆盖口径注释，第 4 页四张平台卡整体左移、明显侵入收入图；文件仍可打开，六页、原生表图和事实内容仍保留。首次较轻位移未产生可见区分并如实保留；加大至肉眼可见但未损坏后，视觉分为 `0.7796`，严格低于 Golden 均值且高于空白/损坏产物。
- 所有渲染图、Codex 事件、五次结果和负例结果保存在 `.coworkevo/vlm-calibration/ppt/officebench-3-2-nvidia-financial-board-briefing/`，不提交仓库。
