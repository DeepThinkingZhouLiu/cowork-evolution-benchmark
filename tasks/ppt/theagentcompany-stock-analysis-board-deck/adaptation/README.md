# TheAgentCompany 股票分析董事会演示逐题适配记录

本题直接阅读了上游 `ds-stock-analysis-slides` 的完整 Query、checkpoints、evaluator 和预置文件来源。`Original_Request.md` 是上游请求的逐字节副本；`stock_data_raw.csv` 从 TheAgentCompany OwnCloud 1.0.0 镜像 `sha256:326406d88345899db22fc1081aa8fa15415c542a71b80486c0a2cebf59f56f63` 中原样提取，包含 3740 行原始记录，并非为本 Benchmark 编造的数据。

逐题审计证明：原始数据中恰有 5 行 `FakeDate_0` 至 `FakeDate_4` 记录同时缺失任务所需字段。删除缺失行、按日期升序排列，并恢复 `Unnamed: 0` 与 `Volume` 的整数类型后，得到的 3735 行结果与上游 evaluator 内冻结的清洗参考逐值、逐列和逐类型一致。原始数据没有证券代码，因此本题不自行猜测或声称具体股票名称。

Golden 演示文稿是针对这一道题人工设计的 5 页可编辑交付：收盘价月度走势、收盘价与 50/200 日均线、7×7 相关矩阵、日收益率直方图和成交量直方图。四个图表均为原生可编辑 PowerPoint 图表，相关矩阵为原生可编辑表格；讲稿逐页对应，并只描述能够从冻结数据重算的事实。Judge 独立读取清洗 CSV、PPTX 内嵌图表序列和表格单元格以及逐页讲稿，不依赖 Golden 文件哈希放行。

## 实际验证

- Golden、ZIP 等价重封装和非语义变体均为 1.0。空/损坏/缺文件交付，以及篡改单个清洗值、交换日期、修改收盘价点、均线点、相关系数符号、直方图区间或讲稿波动率等 23 个定向负例均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，三个产物哈希稳定：PPTX 为 `3994c9b535ec19fb779a147c56804a27c3ed5837ae1ff839dc87847e9ae4ed5a`，清洗 CSV 为 `1b8940d298205bed6700b1f7aea0321385c5aff0599f3fcec173300c9bb3fcbf`，讲稿为 `58749f5bb5bddc1504b5585a47f6772dfd734a879e3aa3027733cc65ca718a07`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `theagentcompany-stock-analysis-oracle-20260829`，trial 为 `theagentcompany-stock-analysis-b__cucRRuZ`；奖励 1.0、0 异常，三个容器产物哈希均与本地 Golden 一致。
- 5 页均已通过 LibreOffice 渲染并逐页目视检查；图表、相关矩阵、标题、说明和页脚没有重叠、裁切或整页压平。
