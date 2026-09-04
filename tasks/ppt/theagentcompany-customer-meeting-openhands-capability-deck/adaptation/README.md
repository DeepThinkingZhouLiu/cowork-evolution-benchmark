# TheAgentCompany OpenHands Customer Meeting 逐题适配记录

本题直接阅读了上游 `pm-prepare-meeting-with-customers` 的完整 Query、checkpoints、evaluator 和 RocketChat NPC 场景。`Original_Request.md` 与 `RocketChat_Recipient_Context.json` 是上游文件的逐字节副本；`OpenHands_Evaluation_Snapshot.json` 只冻结上游 evaluator 真实列出的 3 个领域和 16 个 benchmark 名称，没有自编题目、数据或能力分数。

逐题审计发现：上游证据只能证明这些 benchmark entry 在评测目录中的存在和分组，不能支持分数、通过率、模型版本、排名或客户效果。因此 Golden 将“evaluation inventory”和“measured performance”显式分开，Judge 也会拒绝无来源的性能声称。原任务的 RocketChat 发送动作转换为 `Delivery_Record.json`，明确标记未执行在线副作用。

## 实际验证

- Golden、ZIP 等价重封装和非语义封面变体均为 1.0。空/损坏/单文件交付、Software/Web/Misc 任一 benchmark 漏项、跨领域串项、伪造 95% 通过率、证据边界缺失、收件人/渠道/网络状态错、跨文件冲突、来源缺失、整页压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0。PPTX SHA256 稳定为 `171ed28bbc45f0086240c61f9dafc4031e8f5063ebb77058264982988258245e`，Delivery JSON SHA256 稳定为 `61ef42069eff2a959d9a74afdd985c7d4fd314d304e4ad38ae97e0056497d439`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `theagentcompany-openhands-customer-oracle-20260829`，trial 为 `theagentcompany-customer-meeting__Rq8Yqdj`；奖励 1.0、0 异常，两个容器产物哈希均与本地 Golden 一致。
- 6 页已逐页通过 LibreOffice 渲染和目视检查；8/2/6 项的页内布局、来源页脚和交付页无重叠、裁切或整页压平。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本适配记录、上游 Query/checkpoints/evaluator、三个冻结输入、`build_golden.py` / `build_fixtures.py` / `solve.py`、Golden PPTX 和 `Delivery_Record.json`；核对了 3 个领域、8/2/6 共 16 个 benchmark、四类禁止性能声称与离线交付字段。
- Golden 与两个备选负例均已按 6 页原尺寸渲染并逐页查看：Golden 使用深海军蓝封面/收尾、浅灰内容页和蓝/青/紫领域色，三组 entry 卡片、来源页脚、evidence box 与离线交付卡均完整，无重叠裁切。由于任务允许不同的客户化叙事和视觉方案，只要证据边界与清单正确，视觉评分采用 `reference_free`，Golden 仅用于制订 Rubric，不作为唯一布局模板。
- 专属视觉 Rubric 分别检查封面 3/16/0 定位、8/2/6 总览、Software 8 项、Web 2 项、Misc 6 项、证据/离线交付边界和客户会议完成度；精确字符串、跨领域排他性、Delivery JSON 一致性、来源计数和原生可编辑性继续由确定性 Judge 负责。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；对应综合分均为 `1.0000`。
- 中等可见负例使用本题逐项构造的 `missing-web-benchmark`：只把第 4 页第二张卡的 `MiniWoB` 改成 `Entry omitted`，WebArena、页头、2 entries 声明、其他五页和 Delivery JSON 均保持；文稿可正常打开，视觉版式仍完整但领域清单清楚缺一项，不属于空白或损坏交付。其确定性分 `0.0000`、视觉分 `0.9140`、综合分 `0.2742`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮 6 次 VLM 调用均成功完成，没有基础设施失败或被错误计作 0 分的样本。
