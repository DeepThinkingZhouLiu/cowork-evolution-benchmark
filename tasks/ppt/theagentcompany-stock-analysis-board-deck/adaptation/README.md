# TheAgentCompany 股票分析董事会演示逐题适配记录

本题直接阅读了上游 `ds-stock-analysis-slides` 的完整 Query、checkpoints、evaluator 和预置文件来源。`Original_Request.md` 是上游请求的逐字节副本；`stock_data_raw.csv` 从 TheAgentCompany OwnCloud 1.0.0 镜像 `sha256:326406d88345899db22fc1081aa8fa15415c542a71b80486c0a2cebf59f56f63` 中原样提取，包含 3740 行原始记录，并非为本 Benchmark 编造的数据。

逐题审计证明：原始数据中恰有 5 行 `FakeDate_0` 至 `FakeDate_4` 记录同时缺失任务所需字段。删除缺失行、按日期升序排列，并恢复 `Unnamed: 0` 与 `Volume` 的整数类型后，得到的 3735 行结果与上游 evaluator 内冻结的清洗参考逐值、逐列和逐类型一致。原始数据没有证券代码，因此本题不自行猜测或声称具体股票名称。

Golden 演示文稿是针对这一道题人工设计的 5 页可编辑交付：收盘价月度走势、收盘价与 50/200 日均线、7×7 相关矩阵、日收益率直方图和成交量直方图。四个图表均为原生可编辑 PowerPoint 图表，相关矩阵为原生可编辑表格；讲稿逐页对应，并只描述能够从冻结数据重算的事实。Judge 独立读取清洗 CSV、PPTX 内嵌图表序列和表格单元格以及逐页讲稿，不依赖 Golden 文件哈希放行。

## 实际验证

- Golden、ZIP 等价重封装和非语义变体均为 1.0。空/损坏/缺文件交付，以及篡改单个清洗值、交换日期、修改收盘价点、均线点、相关系数符号、直方图区间或讲稿波动率等 23 个定向负例均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，三个产物哈希稳定：PPTX 为 `3994c9b535ec19fb779a147c56804a27c3ed5837ae1ff839dc87847e9ae4ed5a`，清洗 CSV 为 `1b8940d298205bed6700b1f7aea0321385c5aff0599f3fcec173300c9bb3fcbf`，讲稿为 `58749f5bb5bddc1504b5585a47f6772dfd734a879e3aa3027733cc65ca718a07`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `theagentcompany-stock-analysis-oracle-20260829`，trial 为 `theagentcompany-stock-analysis-b__cucRRuZ`；奖励 1.0、0 异常，三个容器产物哈希均与本地 Golden 一致。
- 5 页均已通过 LibreOffice 渲染并逐页目视检查；图表、相关矩阵、标题、说明和页脚没有重叠、裁切或整页压平。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本适配记录、上游 Query/checkpoints/evaluator 及清洗参考、`Original_Request.md`、真实 3,740 行 `stock_data_raw.csv`、`build_golden.py` / `build_fixtures.py` / `solve.py`、清洗后 3,735 行 CSV、完整五页讲稿与 Golden PPTX。又独立重算了 5 条 FakeDate 缺失记录、179 个月末点、$30.95/$408.46 起止价、MA50/MA200、7x7 相关矩阵、3,734/3,735 条样本的两个 30-bin 直方图和讲稿中全部董事会数字。
- Golden 先渲染为五页接触表，再逐页放大检查长期 Close 折线、Close/MA50/MA200 三线图、7x7 热力表、日收益率直方图和 Volume 直方图；右侧数字卡、Board read、图例、轴标和来源页脚均完整。任务允许不同的董事会视觉系统和图表布局，只要准确、可编辑且证据有界，所以使用 `reference_free`，Golden 只用于制订 Rubric，不作为唯一布局模板。
- 本题专属视觉 Rubric 分别检查收盘价长期趋势与非预测边界、Close/MA50/MA200 三线及非买卖信号、7x7 相关热力表与非因果边界、日收益 30-bin 尾部风险、Volume 30-bin 右偏长尾、五步董事会叙事/来源链和全局数据演示完成度；清洗数据逐值、图表内嵌数列、相关单元格、直方图分箱计数、讲稿数字和原生可编辑性仍由确定性 Judge 负责。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.0000, 0.9450, 0.9795, 0.9730, 0.9865]`，均值 `0.9768`，总体标准差 `0.0182`，小于 `0.05`；对应综合分为 `[1.0000, 0.9835, 0.9939, 0.9919, 0.9960]`。VLM 的轻微扣分主要落在两张 30-bin 直方图的密集 x 轴标签可读性，但不影响主峰、尾部和关键数字的理解。
- 中等可见负例使用本题逐项构造的 `flattened-first-slide`：只将第 1 页长期 Close 趋势图替换为仍可打开的空白整页图片，其他 4 页、清洗 CSV 和讲稿保持；五步董事会叙事缺首页在接触表中清楚可见，但不属于整份空白或损坏交付。其确定性分 `0.0000`、视觉分 `0.7040`、综合分 `0.2112`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮 5 次 Golden 和 1 次负例的 VLM 调用均成功完成，没有基础设施失败或被错误计作 0 分的样本。
