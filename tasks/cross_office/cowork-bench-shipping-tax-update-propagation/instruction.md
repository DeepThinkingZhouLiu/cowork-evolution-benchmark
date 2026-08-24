你在运营团队中负责审核店铺的配送和税率配置。请基于工作区中的上游数据快照，生成一套可回放的配置分析交付物。Excel、Word、JSON 和邮件草稿必须由同一批输入推导，其中任一税率或配送设置变更后，所有相关交付物都必须同步更新。

输入目录包含从上游 Cowork-Bench 数据库逐行静态化的六个 CSV：

- `shipping_zones.csv`：配送区。
- `shipping_zone_locations.csv`：每个配送区的地域代码和类型。
- `shipping_zone_methods.csv`：每个配送区的已配置方法及启用状态。
- `shipping_methods.csv`：配送方法目录，用于校验 `method_id`。
- `tax_classes.csv`：全部税类，包括可能暂无税率记录的税类。
- `tax_rates.csv`：全部已配置税率及其适用范围。

按以下口径处理：

- 配送区按 `order`、`id` 升序；地域和配送方法各按其 `id` 升序。
- 只将 `enabled=t` 的配送方法计入 `Methods_Count`；方法标题用英文分号加空格 `; ` 连接。
- 地域用 `<type>:<code>` 表示，多个地域用 `; ` 连接。不得重命名、翻译或推断输入中的区域代码。
- 税类按 `name` 字母序；同一税类的税率按 `id` 升序。每条税率必须保留。税类存在但没有税率记录时，仍在 Excel 中保留一行，`State`、`Rate`、`Compound`、`Shipping` 留空。
- 不得把区域名、销量、收入、客户需求、市场基准、节省金额、ROI 或预测补成输入未提供的事实。建议必须明确基于已观察到的配置差异。

请在输出目录交付以下五个核心文件：

- `shipping_tax_update_processor.py`：可回放处理器，必须支持 `python shipping_tax_update_processor.py --input-dir <输入目录> --output-dir <输出目录>`，且从指定输入重新生成其余四个交付物。
- `shipping_tax_results.json`：机器可读分析结果，顶层键为 `shipping_zones`、`tax_classes`、`summary`，并保留输入中的配送方法设置、税率名称和适用范围。
- `Shipping_Tax_Analysis.xlsx`：原生可编辑 Excel。
- `Operations_Report.docx`：原生可编辑 Word 报告。
- `operations_email.txt`：离线邮件草稿，不得实际发送。

Excel 必须恰好包含两张工作表：

- `Shipping Zones`：列依次为 `Zone_Name`、`Methods_Count`、`Regions`、`Methods`。`Methods_Count` 必须是从同行 `Methods` 单元格动态计数的原生 Excel 公式，不得填入硬编码常量。
- `Tax Rates`：列依次为 `Tax_Class`、`State`、`Rate`、`Compound`、`Shipping`，按上述空税类规则展开。
- 两张表都应有专业表头、冻结首行、筛选或原生表格、合理列宽和数字格式；不得将表格压平为图片。

Word 报告必须准确包含 `Executive Summary`、`Shipping Coverage`、`Tax Structure`、`Recommendations`、`Source Boundary` 五个章节。配送和税率明细必须用原生可编辑表格展示。报告需写出所有配送区、地域、方法数和方法标题，以及所有税类、税率名称、适用范围、税率、`Compound` 和 `Shipping` 状态，并明确指出无配置记录的税类。

邮件草稿的收件人必须是 `operations@store.com`，主题必须恰为 `Shipping and Tax Configuration Report`。正文需概括配送区、方法数、税类、已配置税率数、最高已配置税率和建议。
