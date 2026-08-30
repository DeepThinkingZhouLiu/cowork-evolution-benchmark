# AutomationBench 4062 逐题适配记录

本题直接阅读 `finance.sales_tax_remittance` 的 Query、完整 initial state 和 assertions，并逐字段静态化 8 行 Q4 revenue、3 行 tax rates、政策邮件和 NY exemption 邮件。没有新编州、月份、税率或销售额。

制题审计发现上游邮件说 exemption 金额是 `approximately $30,000` / `$45,000`，但断言据此精确计算 NY `$16,800`。本适配不隐瞒：Golden 保留 `≈`，同时说明采用上游 asserted arithmetic。Judge 独立从输入重算 CA、NY、OR 月度与州级金额，并检查“只准备、不得申报、不得付款”的授权边界。

## 实际验证

- Golden、ZIP 等价重封装和非语义性封面变体均为 1.0。来源漏行/错数、CA 生效税率错、NY exemption 错、月度税额错、州图/汇总错、遗漏 approximately 边界、违规申报/付款、错收件人、缺来源、压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，输出 SHA256 稳定为 `9147f69684bac9c51d73d459cabdfd4c0ea536f15b2db29441856bf4a44a2cfc`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `automationbench4062-oracle-sales-tax-20260829`，trial 为 `automationbench-4062-q4-sales-ta__jqMhBEs`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 6 页已逐页目视；来源表、月度 schedule、政策例外、州图和授权页均无重叠裁切，`visual_qa.py` 返回空问题列表。
