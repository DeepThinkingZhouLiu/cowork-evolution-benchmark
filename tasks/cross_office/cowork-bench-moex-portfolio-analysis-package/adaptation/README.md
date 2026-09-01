# Cowork-Bench MOEX 投资组合分析任务逐题适配记录

## 为什么选择这道题

上游要求把投资组合、MOEX 行情数据库和基金政策合并成 Excel 分析与 Word 备忘录，再通过邮件提交。它的难点是行情日期边界、公式化财务计算、政策约束和附件版本必须同时正确；任一上游错误都会传播到三份交付物，适合 `extract_analyze_present` 家族。

结构设计目标为 L3 / Hard：三份异构来源、115 行行情、日期窗口选择、五类公式、政策证据边界和三文件强一致性共同构成深耦合。按照项目规则，正式难度仍为 `unassigned`，等待冻结多 Harness Rollout 后统一校准。

## 实际读入上下文的上游材料

- `docs/task.md`：完整 Query、30 日窗口、行业、市值和交付要求。
- `docs/agent_system_prompt.md`、`task_config.json`：原 Excel、Word、邮件、MOEX 数据库和文件工作区约束。
- `preprocess/main.py`：上游文件、数据库状态和初始化流程。
- `evaluation/main.py`：原评分逻辑；已完整阅读并识别其静态关键词、弱公式检查和错误 Golden 依赖。
- `initial_workspace/Portfolio.xlsx`、`Investment_Policy.pdf`：逐表、逐页读取的原始输入。
- `db/init.sql.gz`：逐行审计 `moex.stock_info` 与 `moex.stock_prices` 中五个目标代码的 115 条价格记录。
- `groundtruth_workspace/Portfolio_Analysis.xlsx`、`Investment_Memo.docx`：逐表、逐段审计的上游 Golden。
- Cowork-Bench 根目录 `LICENSE`，固定 commit `d943e75bc0fc8e3b27141979300cd8cbcd1e890d`。

上述 Query、系统提示、配置、预处理、Evaluator、错误 Golden 和许可证的精确快照均保存在本目录。

## 上游 Golden 审计结论

上游 Golden 与任务输入发生根本冲突：它使用 `GOOGL / AMZN / JPM / JNJ / XOM`、USD 和美国交易所，而冻结 Portfolio 的真实代码是 `SBER.ME / GAZP.ME / LKOH.ME / MGNT.ME / MTSS.ME`，政策要求 RUB 与 MCX。因此本题没有复制上游答案，而是从冻结输入独立重算。Judge 还专门拒绝这些错误美股代码和 USD 数据回流。

## 冻结行情如何物化

- `Portfolio.xlsx` 与 `Investment_Policy.pdf` 是上游原文件的逐字节副本。
- `moex_market_snapshot.csv` 从上游 `db/init.sql.gz` 中逐行提取目标五个代码的 `stock_info` 与 `stock_prices`，共 115 行；未改写公司、行业、市值、日期或收盘价。
- 最新日固定为 2026-05-26；“不晚于 2026-04-26 的最新交易日”固定为 2026-04-24。任务无网络，禁止用实时行情替换。

## 人工真值台账

| Symbol  | Quantity | Baseline 2026-04-24 | Current 2026-05-26 | Current value | 30d return | Sector                 |
|---------|---------:|--------------------:|-------------------:|--------------:|-----------:|------------------------|
| SBER.ME |      100 |            127.7280 |           120.8171 |    12,081.7100 |  -5.410638% | Financial Services     |
| GAZP.ME |      200 |            211.2980 |           208.6306 |    41,726.1200 |  -1.262383% | Energy                 |
| LKOH.ME |       10 |          3,754.0441 |         3,727.0965 |    37,270.9650 |  -0.717828% | Energy                 |
| MGNT.ME |       15 |          4,216.7609 |         4,182.1923 |    62,732.8845 |  -0.819798% | Consumer Defensive     |
| MTSS.ME |      150 |            250.1673 |           255.0623 |    38,259.3450 |   1.956691% | Communication Services |

独立汇总：总市值 `192,071.0245 RUB`；最佳 `MTSS.ME +1.956691%`；最差 `SBER.ME -5.410638%`；Energy 为 `78,997.085 RUB / 41.129101%`；5/5 持仓均为 RUB、MCX、允许行业且市值严格大于 1000 亿 RUB。政策没有持仓或行业权重上限，不能自行编造。

## 因果依赖图

```text
Portfolio.xlsx + 115 行 MOEX 快照 + Investment_Policy.pdf
-> 日期边界、发行人字段与五个持仓公式
-> Summary 行业聚合与政策结论
-> Word 投资备忘录
-> EML 正文 + 当次 Excel/Word 字节一致附件
```

## 对上游的必要改造

- 保留上游真实 Portfolio、政策和数据库事实，不新增证券、行情或政策阈值。
- 将动态 MOEX 数据库的目标子集静态化，使 Harbor 可以无网络确定回放。
- 重做公式化 Excel、可编辑 Word 和带真实附件的 EML，不沿用错误上游 Golden。
- Judge 从三份输入独立重算全部边界、持仓、行业、政策和邮件附件，不比较 Golden 哈希。
- 对无来源因果、自动买卖结论、虚构政策阈值、美股代码和 USD 泄漏设置专门的 hurdle 与 penalty。

## 许可、隐私与泄漏

来源为 Apache-2.0 Cowork-Bench fixture。任务的公司、证券、行情和政策是上游合成工作区状态，保守标记 `synthetic=true`；本适配没有引入个人数据。任何复用相同五个持仓、115 行行情或政策的变体必须进入 `cowork-bench-moex-portfolio-analysis` leakage group，不能跨 split。

## 验证结果

- 协议、输入 SHA256、Task Metadata、Provenance 与 Rubric Schema 全部通过。
- Golden 与两次空目录 fresh replay 的 Judge Reward 均为 `1.0`；三份产物逐字节一致。
- 14 个定向负例全部低于 `0.8`，最高负例得分为 `0.4`。
- XLSX 与 DOCX 经 LibreOffice 渲染为 6 页和 2 页，8 页已逐页目视检查，无截断或公式错误。
- 官方 Harbor 0.21.0 Docker Oracle 独立回放两次，均为 Reward `1.0`、0 异常；容器内三份 `/app/output` 产物逐字节一致。
