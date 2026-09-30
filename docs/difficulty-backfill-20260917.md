# 240题难度回填与主线整理（2026-09-17）

## 统一入口

开发分支为 `lz-dev`，本机目录为 `.worktrees/origin--lz-dev--主开发线`；完整内容同步至 `main`。当前Release为 `benchmarks/cowork-evo-240-v0.5/`。历史工作树、实验产物和旧Release保留，日常使用当前目录的 `tasks/` 即可。

## 回填依据

| 模块         | 题数 | 原始来源                                      | 分档处理                       |
|--------------|-----:|-----------------------------------------------|--------------------------------|
| PPT          |   60 | difficulty-labels-v0.5.csv                    | 保留已有逐题排名与档位         |
| Word         |   60 | difficulty-labels-v0.5.csv                    | 保留升级后已有逐题排名与档位   |
| Excel        |   60 | difficulty-labels-v0.5.csv                    | 保留升级后已有逐题排名与档位   |
| Cross-Office |   60 | difficulty-labels-v0.5.csv                    | 按三Harness混合分均值复算排名  |

上表文件均位于 `benchmarks/`。统一清单 `benchmarks/difficulty-labels-v0.5.csv` 保存240道题的路径、等级、标签、名次、三Harness分数、Panel均分、原报告路径及SHA256。未调用Agent或VLM重新评分。

每个模块都按已有方案分为Easy 12、Medium 18、Hard 18、Expert 12。比例是预先约定的，具体题目归属按域内实测排名确定；等级不能解释为跨模块相同的绝对得分阈值。PPT与Cross-Office的三处边界无同分；Word的Hard/Expert零分边界保留原报告的预计工作量、稳定ID次序，同分跨档不表示实测能力差异。Excel历史负向罚分原样保留。

标签同时写入 `task.toml` 的 `[metadata]` 与 `metadata/task.json`，机器值统一为 `L1/easy`、`L2/medium`、`L3/hard`、`L4/expert`。两份配置各自原有的来源说明保留，`notes` 追加 `calibration_status=pilot_single_panel` 和统一清单路径。此前Word/Excel的大写标签不符合Schema，以及 `officebench-1-19-agreement-terms` 的TOML/JSON难度不一致，本次一并修复。

## 证据边界

当前证据是Sonnet 5配Claude Code、OpenCode、OpenHands的单轮Panel，不是三个独立模型。标签已固定用于当前版本分析，独立跨模型与重复运行验证仍未完成。PPT/Cross-Office沿用既有混合分，Word/Excel沿用代码分；没有改变Judge、Rubric、输入、Golden或历史得分。

## Release与分区

v0.5继承v0.4的全部逐题分区，不重新洗牌：PPT/Word/Excel各30/10/20，Cross-Office为0/0/60，合计90/30/120。仅重新生成难度字段、split内容哈希和dataset.lock。历史v0.4曾只更新内容锁、未同步split哈希，本次保留旧文件并用新版本解决一致性问题。

泄漏检查覆盖输入SHA256、provenance与声明的leakage group，不等同于自动证明所有语义近重复均不存在。难度校准用分数只用于标签；本次不调整训练、验证、测试归属。

## 验证

- `tools/sync_difficulty_labels.py`：核对240题与来源逐一对应、三Harness配对、排名、均分、标签和来源SHA256；重复执行不产生新修改。
- 项目资产校验：检查240题目录、配置、输入哈希、Golden及隐藏侧入口。
- Cross-Office：对照本地180份逐题JSON与原CSV，153正常完成、27任务失败、0基础设施失败；复核70/30合成分。
- 配置差异审计：题目目录只修改240份TOML与240份元数据JSON中的难度及校准说明。
- Release校验：核对任务树、逐题内容锁、split互斥、90/30/120数量、PPT分层以及跨分区泄漏组件。
- 回归测试：40/40通过，包含来源说明保留、基础设施失败拒绝与不完整Panel拒绝检查。

最终静态资产240/240通过，TOML/JSON标签240/240一致，split标签240/240与统一清单一致；v0.5 Release与Benchmark Schema校验通过，任务分区迁移数0，跨分区来源/输入共享组件0。可复核记录见 `benchmarks/cowork-evo-240-v0.5/migration.json`。

本次是元数据回填，不重新运行Docker Oracle或付费模型；原题目作者侧回放证据继续保留。
