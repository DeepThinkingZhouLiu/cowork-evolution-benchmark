# AutomationBench 4044 逐题适配记录

本题直接阅读 `finance.po_three_way_match` 的 Query、完整 initial state 和 assertions，并逐字段静态化 Invoice Batch、Purchase Orders、Receiving Reports 和 tolerance policy。没有新编供应商、金额、数量、日期或状态。

制题审计发现一个重要的上游语义缺口：政策只定义“纯价格偏差在 5% 以内”时进入 Price Variance Review，但 VI-002 的偏差是 7.14%，上游 assertion 仍指定 Review。本适配没有把这个矛盾藏起来：Golden 保留上游状态，同时要求显式标出政策缺口；Judge 两者独立检查。

## 实际验证

- Golden、ZIP 等价重封装和非语义性封面变体均为 1.0。来源行错、VI-001 边界错、VI-002 状态错、VI-003 数量/状态错、图表错、遗漏政策缺口、收件人错、矛盾批准声明、来源缺失、整页压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，输出 SHA256 稳定为 `760721adac961f2bcb16c2a4c91e1c3d13dd046bc797238c73613b6633b7101b`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `automationbench4044-oracle-three-way-match-20260829`，trial 为 `automationbench-4044-three-way-m__WcfKJsw`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 6 页已逐页目视；来源对齐表、决策表、异常卡、原生金额图和 AP 通知页无重叠裁切，`visual_qa.py` 返回空问题列表。
