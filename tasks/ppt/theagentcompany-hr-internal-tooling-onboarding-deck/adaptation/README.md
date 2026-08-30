# TheAgentCompany HR Internal Tooling 逐题适配记录

本题直接阅读了上游 `hr-internal-tooling-slides` 的完整 Query、checkpoints 和 evaluator。为避免把网页任务缩成我们自己编的文本，又从官方 `servers-owncloud:1.0.0` 镜像中逐字节提取了 `TAC_overview.md` 和 `Personnell_File.odt`，并从上游 evaluator、RocketChat NPC 目录、Plane README、OwnCloud 目录树和 tagline 文件静态化工具快照。

逐题审计发现：真实冻结镜像的 `/Documents` 有 9 个顶层业务目录，而 legacy evaluator 遗漏 `Marketing` 与 `Q1 Planning and Allocation`，还把实际的 `Human Resources Team` 写成 `Human Resource Team`。本适配以冻结镜像为真值，并强制演示暴露这个语义差异。原题的 OwnCloud 上传与 RocketChat 消息被转换为可离线重放的 `Delivery_Record.json`，明确标记没有执行在线副作用，不伪造发送成功。

## 实际验证

- Golden、ZIP 等价重封装和非语义封面变体均为 1.0。缺交付文件、TAC 产品错、频道/default 错、repo 漏、目录错、evaluator 差异漏、Plane metric 错、收件人/路径错、跨文件冲突、来源缺失、整页压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0。PPTX SHA256 稳定为 `f2aa5b41e75a8b1a548d481be4aa27e43992d636a2fa438e4d2355a8a57ceed7`，Delivery JSON SHA256 稳定为 `8b00b6777153089a8698abf85c2b5f996f0c52d9015d7572722c3fbacb35d77c`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `theagentcompany-hr-tooling-oracle-20260829`，trial 为 `theagentcompany-hr-internal-tool__obspnwj`；奖励 1.0、0 异常，两个容器产物哈希均与本地 Golden 一致。
- 7 页已逐页目视；TAC overview、15 个频道、14 个 repo、9 个目录、Plane 指标和离线交付页无重叠裁切，`visual_qa.py` 返回空问题列表。
