# TheAgentCompany HR Internal Tooling 逐题适配记录

本题直接阅读了上游 `hr-internal-tooling-slides` 的完整 Query、checkpoints 和 evaluator。为避免把网页任务缩成我们自己编的文本，又从官方 `servers-owncloud:1.0.0` 镜像中逐字节提取了 `TAC_overview.md` 和 `Personnell_File.odt`，并从上游 evaluator、RocketChat NPC 目录、Plane README、OwnCloud 目录树和 tagline 文件静态化工具快照。

逐题审计发现：真实冻结镜像的 `/Documents` 有 9 个顶层业务目录，而 legacy evaluator 遗漏 `Marketing` 与 `Q1 Planning and Allocation`，还把实际的 `Human Resources Team` 写成 `Human Resource Team`。本适配以冻结镜像为真值，并强制演示暴露这个语义差异。原题的 OwnCloud 上传与 RocketChat 消息被转换为可离线重放的 `Delivery_Record.json`，明确标记没有执行在线副作用，不伪造发送成功。

## 实际验证

- Golden、ZIP 等价重封装和非语义封面变体均为 1.0。缺交付文件、TAC 产品错、频道/default 错、repo 漏、目录错、evaluator 差异漏、Plane metric 错、收件人/路径错、跨文件冲突、来源缺失、整页压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0。PPTX SHA256 稳定为 `f2aa5b41e75a8b1a548d481be4aa27e43992d636a2fa438e4d2355a8a57ceed7`，Delivery JSON SHA256 稳定为 `8b00b6777153089a8698abf85c2b5f996f0c52d9015d7572722c3fbacb35d77c`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `theagentcompany-hr-tooling-oracle-20260829`，trial 为 `theagentcompany-hr-internal-tool__obspnwj`；奖励 1.0、0 异常，两个容器产物哈希均与本地 Golden 一致。
- 7 页已逐页目视；TAC overview、15 个频道、14 个 repo、9 个目录、Plane 指标和离线交付页无重叠裁切，`visual_qa.py` 返回空问题列表。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本适配记录、上游 Query/checkpoints/evaluator、四个冻结输入（包括解包读取 `Personnell_File.odt`）、`build_golden.py` / `build_fixtures.py` / `solve.py`、Golden PPTX 和 `Delivery_Record.json`；在上下文中核对了 mission/vision、6 产品、15 频道、14 仓库、9 目录、Plane 六步流程/四指标以及 Chen Xinyi 离线交付语义。
- Golden 按原尺寸渲染为 7 页，先查看全套接触表，再逐页放大检查封面、TAC Overview、RocketChat、GitLab 用法、GitLab Repos、OwnCloud 和 Plane/离线交付页。视觉系统以深海军蓝封面/收尾与浅灰卡片内容页组成；逐页检查证实 RocketChat 页部分长频道名与用途栏距较紧，因此不把 Golden 机械视为视觉满分。任务允许不同的新员工培训叙事和版式，所以使用 `reference_free`，Golden 只用于制订 Rubric，不作为唯一布局模板。
- 本题专属视觉 Rubric 分别检查封面四工具操作地图、TAC mission/vision/六产品、RocketChat 5/4/6 分组与 default 频道、GitLab 三个工作流、14 仓库注册表、OwnCloud 九目录与 legacy evaluator 差异、Plane 流程/四指标/离线交付边界及全套完成度；精确字符串、清单排他性、Delivery JSON 一致性、来源计数和原生可编辑性继续由确定性 Judge 负责。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、the configured VLM provider 真实独立评分 5 次，视觉分为 `[0.915, 0.850, 0.915, 0.915, 0.915]`，均值 `0.9020`，总体标准差 `0.0260`，小于 `0.05`；对应综合分为 `[0.9745, 0.9550, 0.9745, 0.9745, 0.9745]`。VLM 的主要扣分证据是 RocketChat 页长频道名和用途文本的局部拥挤，与逐页人工视觉检查一致。
- 中等可见负例使用本题逐项构造的 `flattened-rocketchat`：只将第 3 页 RocketChat 的 15 频道信息图替换为仍可打开的空白整页图片，其他 6 页与 `Delivery_Record.json` 保持；这使 5/4/6 频道地图和 default channels 的缺失在接触表中明显可见，但不属于整份空白或损坏交付。其确定性分 `0.0000`、视觉分 `0.7650`、综合分 `0.2295`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮 5 次 Golden 和 1 次负例的 VLM 调用均成功完成，没有基础设施失败或被错误计作 0 分的样本。
