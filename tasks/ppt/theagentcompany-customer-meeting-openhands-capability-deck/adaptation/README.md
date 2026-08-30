# TheAgentCompany OpenHands Customer Meeting 逐题适配记录

本题直接阅读了上游 `pm-prepare-meeting-with-customers` 的完整 Query、checkpoints、evaluator 和 RocketChat NPC 场景。`Original_Request.md` 与 `RocketChat_Recipient_Context.json` 是上游文件的逐字节副本；`OpenHands_Evaluation_Snapshot.json` 只冻结上游 evaluator 真实列出的 3 个领域和 16 个 benchmark 名称，没有自编题目、数据或能力分数。

逐题审计发现：上游证据只能证明这些 benchmark entry 在评测目录中的存在和分组，不能支持分数、通过率、模型版本、排名或客户效果。因此 Golden 将“evaluation inventory”和“measured performance”显式分开，Judge 也会拒绝无来源的性能声称。原任务的 RocketChat 发送动作转换为 `Delivery_Record.json`，明确标记未执行在线副作用。

## 实际验证

- Golden、ZIP 等价重封装和非语义封面变体均为 1.0。空/损坏/单文件交付、Software/Web/Misc 任一 benchmark 漏项、跨领域串项、伪造 95% 通过率、证据边界缺失、收件人/渠道/网络状态错、跨文件冲突、来源缺失、整页压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0。PPTX SHA256 稳定为 `171ed28bbc45f0086240c61f9dafc4031e8f5063ebb77058264982988258245e`，Delivery JSON SHA256 稳定为 `61ef42069eff2a959d9a74afdd985c7d4fd314d304e4ad38ae97e0056497d439`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `theagentcompany-openhands-customer-oracle-20260829`，trial 为 `theagentcompany-customer-meeting__Rq8Yqdj`；奖励 1.0、0 异常，两个容器产物哈希均与本地 Golden 一致。
- 6 页已逐页通过 LibreOffice 渲染和目视检查；8/2/6 项的页内布局、来源页脚和交付页无重叠、裁切或整页压平。
