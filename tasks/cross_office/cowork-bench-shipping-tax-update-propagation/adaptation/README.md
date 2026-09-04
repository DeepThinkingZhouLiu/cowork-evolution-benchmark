本题是对 Cowork-Bench 单题 `insales-shipping-tax-excel-word` 的逐文件人工改造，不是批量生成任务。

来源和许可：

- 上游仓库：`https://github.com/0717376/cowork_bench.git`
- 固定 commit：`d943e75bc0fc8e3b27141979300cd8cbcd1e890d`
- 原题路径：`tasks/finalpool/insales-shipping-tax-excel-word`
- 许可证：Apache-2.0，完整副本见 `LICENSE-Cowork-Bench`
- 原 Query、system prompt、task config、email config、preprocess、evaluator 和两份原 Golden 均原样保存在本目录。

静态化方法：

- 原题通过 WooCommerce/InSales 本地 MCP 读取动态 PostgreSQL 数据。
- `build_static_snapshot.py` 直接解析上游 `db/init.sql.gz` 中的六个 `COPY wc.* FROM stdin` 块，不改写、不补行、不生成业务值。
- 六个原始 COPY 块的精确文本保存为 `upstream-db-source-snapshot.sql`，便于逐行审计。
- 输入总计 27 行、1704 字节，每个文件的 SHA256 见 `../metadata/input-manifest.json`。
- 上游未声明这些店铺配置来自真实商户；因此本题诚实标记为“上游 Benchmark seed fixture”，只能证明没有由本次改造新编记录。

人工改写范围：

- 将俄文原 Query 人工改写为中文离线 Harbor 题，保留“配送区+税率 -> Excel+Word+运营邮件”的原业务目标。
- 用静态 CSV 取代动态 MCP 读取，用离线邮件草稿取代真实发送。
- 增加可回放 processor 和机器可读 JSON，便于 Judge 从输入重算。
- 将 `Methods_Count` 指定为从方法文本动态计数的 Excel 公式，并增加将 Reduced Rate 从 5 改为 7 的变异回放。

上游 Golden 的已知问题：

- 原 Word 将当前数据库中的俄文配送区名称改成 `California`、`Domestic US`、`International`，与同一 Golden 中 Excel 和当前源快照不一致。
- 原 Word 宣称两条 CA 税率“可能反映该州高订单量”，但原题输入没有订单量。
- 因此原 Golden 仅作溯源材料保留，本题 Golden 不从它复制数据，而是从本题六份输入独立重算。

## v2 连续评分闭环

- `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/regression_cases_v1.py` 逐字节保留第一版严格口径；`SCORING_VERSION=v1|v2` 可显式回放，默认 v2。
- v2 仅在 processor、JSON、Excel、Word 或邮件缺失、为空、损坏/不可解析时硬归零；其他错误按 3 个配送区、5 个地域、6 个方法、settings、3 个税类、7 条税率、3 条公式、Word 表格、邮件事实和 5->7 跨文件传播逐原子单元评分。
- 任务专属负例共 46 个：17 个缺失/空/损坏/原样输入致命反例均为 0；29 个单事实、公式、更新传播、来源边界、中度缺失与严重但可解析反例落在 `0.447107–0.999566`。其中整块配送信息缺失为 `0.748950`，严重不完整但仍可解析的包为 `0.447107`，与单字段错误有明确分层。
- Golden、fresh replay 与 v1 Golden 均为 `1.0`。本题输出只有 JSON/XLSX/DOCX/TXT，没有 PPTX/PDF，因此不强加 VLM，最终分即 deterministic reward。
