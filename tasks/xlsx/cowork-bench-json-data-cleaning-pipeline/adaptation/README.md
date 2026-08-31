本题只改造 Cowork-Bench `fetch-terminal-data-pipeline`，没有根据分类名编造输入、Query 或业务数字。

- 上游仓库：`0717376/cowork_bench`
- 上游 commit：`d943e75bc0fc8e3b27141979300cd8cbcd1e890d`
- 上游路径：`tasks/finalpool/fetch-terminal-data-pipeline`
- 许可：Apache-2.0
- 原始任务：读取两份内部 API JSON，按 order/product 主键保留首条去重，统一商品名大小写，生成两份清洗 JSON、三表 Excel、memory 记录和同内容 Google Sheet。

人工逐文件阅读并保留了原始俄文 Query、system prompt、preprocess、Evaluator、task config、两份输入、清洗脚本、两份清洗 JSON、Excel、memory Golden 和仓库 LICENSE。`data/input_files/` 的三份文件与上游逐字节一致；动态 Google Sheet 副作用只被替换为可离线回放的 `Pipeline_Dashboard.xlsx`。

上游 Golden 有一个确定性冲突：原始 `inventory_api_response.json` 的仓库是 `Склад А/Б/В`，原 Query 和 `clean_data.py` 只要求修改 `product_name`，但冻结 `cleaned_inventory.json` 把仓库额外翻译成 `Warehouse A/B/C`。本题保留该 Golden 供审计，但本地 Golden 与 Judge 均严格保留第一次出现记录中的俄文 warehouse，不接受这项无指令改写。

可从输入直接复算的真值：销售原始 20 条，按 `order_id` 保留首条后 18 条；库存原始 7 条，按 `product_id` 保留首条后 5 条。商品聚合依次为 Widget Alpha `35 / 525 / 15`、Widget Beta `34 / 850 / 25`、Gadget Pro `17 / 850 / 50`、Sensor X1 `23 / 805 / 35`、Connector Z `75 / 750 / 10`；总销量 184，总收入 3,780 卢布。Sensor X1 与 Connector Z 为 LOW，其余为 OK。

来源边界：这些 API 记录是 Cowork-Bench 公开合成 fixture，不是真实公司的生产数据；“真实来源”指题目、Query、输入与规则完整取自用户指定且许可清楚的开源仓库，而不是本项目现场编造。

## 难度档位调整记录（2026-08-31）

本题从 L3/advanced 调整为 L4/expert 的作者结构假设，依据是：两份 JSON 清洗、可回放脚本、两个 Excel 报告和进度状态形成跨文件数据管线。 难度来自真实工作流中的异构输入、跨文件一致性、状态或专业分析链条，不是单次 Sonnet 5 得分。calibration_status=uncalibrated、review_status=pending_human_review 保持不变，最终需结合多模型结果、失败原因和人工复核。
