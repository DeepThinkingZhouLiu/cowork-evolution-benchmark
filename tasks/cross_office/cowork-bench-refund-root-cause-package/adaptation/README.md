本题不是自编任务。它是对 Cowork-Bench `insales-refund-root-cause-excel-word-email` 的单题人工改造，上游固定在 commit `d943e75bc0fc8e3b27141979300cd8cbcd1e890d`，原仓库采用 Apache-2.0。

改造前已逐文件阅读原始俄文 Query、system prompt、task config、preprocess、Evaluator、`Quality_Policy.pdf`、`guide.md`、PDF 生成脚本、两个冻结 Golden 和仓库 LICENSE。这些上游材料已保存在本目录，便于独立审计。

对 Benchmark 做了以下必要适配：

- 将原题依赖的动态 InSales/PostgreSQL 表静态化为 `refunds.csv`、`refund_orders.json`、`products.csv` 和 `product_reviews.csv`。`build_static_snapshot.py` 完整记录了从上游 `db/init.sql.gz` 到输入文件的投影方法；脚本不生成新记录、数值或文本。
- 保留全部 13 条退款、所有关联订单、全部 82 个商品和全部 396 条评论。PDF 和 guide 是上游原文件的逐字节副本。
- 将原题对真实发信服务的依赖替换为两份 `.txt` 离线草稿；收件人、主题关键字和正文事实要求不变。
- 要求交付可回放处理脚本和 JSON 中间结果，使 Judge 能对变异输入重新执行，拒绝固定答案。
- 上游 preprocess 会将且仅将订单 39、49、110 的单笔全额退款金额对齐为订单总额，对应退款从 `25.44 -> 66.45`、`27.87 -> 88.03`、`887.57 -> 1221.32`。上游冻结 `Refund_Analysis.xlsx` 仍使用对齐前数字，因而总金额 3972.91 和严重度 8/3/2 已与上游 preprocess/Evaluator 的动态真值冲突。本题保留这份过期 Golden 仅供审计，实际 Judge 对 preprocess 后的输入重算，真值为总金额 4407.83 和严重度 10/1/2。

本题输入来自上游公开 Benchmark 种子，本身带有上游合成数据属性；它的“真实来源”指任务和数据均从用户指定的公开仓库逐题选取，而不是由本项目新编。

## v2 连续评分改造（2026-09-04）

- `rubric.v1.json` 与 `tests/judge_v1.py` 原样封存原评分协议；本题原先没有独立 `negative_cases.py`，因此不存在可封存的 v1 负例文件。`SCORING_VERSION=v1|v2` 可显式复现，默认使用 v2。
- v2 将处理器原始回放、金额/评分变异回放、JSON 六个区域、Excel 的 13×7 退款字段、24×7 商品字段和九项汇总、Word 六章节及事实、两封邮件事实拆成任务专属原子单元连续计分；局部错误不再经 hurdle 与 penalty 重复扣成整项 0 分。
- 只有六个核心交付缺失、为空、损坏/不可解析，处理脚本语法无效，或拿冻结输入逐字节冒充交付时硬归零。无来源金额、百分比、行业基准、客户流失、ROI 与供应商缺陷率只由单一证据边界处罚处理；额外输出单独轻罚。
- Golden 和从全部冻结输入 fresh replay 的 v2、v1 Reward 均为 `1.0`。13 个缺失、空、损坏或输入冒充致命反例全部为 `0`；30 个局部与分层反例覆盖单退款、单商品、汇总、排序、Word、邮件、处理器回放、跨载体冲突、证据越界和额外文件，得分范围为 `0.603322–0.999682`。其中单事实错误保留高部分分，多组件仍可解析但内容严重不完整的反例明显降至 `0.603322`。
- 本题核心交付为 Python、JSON、XLSX、DOCX 与 TXT，不含 PPTX 或 PDF 视觉交付；按统一协议只使用确定性 Judge，不强加 VLM，也不存在待解决的 VLM 基础设施失败。
