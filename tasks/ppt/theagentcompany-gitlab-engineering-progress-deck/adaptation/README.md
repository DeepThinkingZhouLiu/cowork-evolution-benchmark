# TheAgentCompany GitLab Engineering Progress 逐题适配记录

本题直接阅读了上游 `pm-present-gitlab-info-as-ppt` 的完整 Query、checkpoints、evaluator 和参考答案，从 evaluator 中逐字段冻结 14 个 GitLab 仓库的顺序、描述、Issue 总数和 Merge Request 总数，没有自编仓库或业务数字。`Original_Request.md` 是上游 task.md 的逐字节副本。

逐题审计发现上游 evaluator 有三个会造成假阳性的缺口：名称、描述、Issue 和 MR 每类只要超过一半仓库正确即可得分；数字允许 `+/-5`；同一仓库的四个字段可以从不同页面分别命中。本 Harbor 适配改为固定 15 页，14 个仓库按页序一一对齐，描述和两类数值逐页精确核对。对上游为空的三个描述，显式要求标记元数据缺口，不允许脑补。

## 实际验证

- Golden、ZIP 等价重封装和非语义封面变体均为 1.0。损坏/空交付、仓库页调序、漏仓库页、描述错误、空描述被脑补、Issue 只错 1、MR 只错 1、封面总数错、计数口径错、单页误称 open only、来源缺失、整页压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，PPTX SHA256 稳定为 `ad38f21f0a2be3b6295e6911bf5e583963333b1cb661457864ece532d9790771`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `theagentcompany-gitlab-progress-oracle-20260829`，trial 为 `theagentcompany-gitlab-engineeri__cvuWnhs`；奖励 1.0、0 异常，容器产物哈希与本地 Golden 一致。
- 15 页均已通过 LibreOffice 渲染；封面、空描述页、高数值页和最长描述页已逐页目视，无重叠、裁切或整页压平。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本适配记录、上游 Query/checkpoints/evaluator、两个冻结输入、`build_golden.py` / `build_fixtures.py` / `solve.py` 和 Golden PPTX；在上下文中逐仓库核对了 14 个名称、顺序、描述、Issue/MR 精确数字、三个空描述仓库和 all-states 计数语义。
- Golden 按原尺寸渲染为 15 页并通过两张接触表逐页检查；封面、api-server、opensearch、risingwave 又单页放大复核。主体为深海军蓝 CTO 封面和 14 个浅灰仓库信息卡，14/10/4,990/794 概览、三个 `METADATA GAP · NOT INFERRED` 缺口页、绿/橙活动条、冻结口径与来源页脚均完整，无重叠裁切。由于只要准确呈现冻结数据，合格答案可采用不同的 CTO 叙事和版式，因此使用 `reference_free`，Golden 只用于制订 Rubric，不作为唯一布局模板。
- 本题专属视觉 Rubric 分别检查 CTO 封面与四个总数、1+14 页及冻结仓库顺序、逐页仓库身份/描述/指标层级、三个空描述缺口、Issues/MR 活动条编码、all-states/冻结来源可见性和 CTO 汇报完成度；精确字符串、数字、页序、来源计数和原生可编辑性仍由确定性 Judge 负责。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；对应综合分均为 `1.0000`。
- 中等可见负例使用本题逐项构造的 `flattened-repo-slide`：只将第 11 页 opensearch 仓库卡替换为仍可打开的空白整页图片，其他 14 页保持；接触表中缺页清楚可见，但不属于整份空白或损坏交付。其确定性分 `0.0000`、视觉分 `0.5650`、综合分 `0.1695`，严格低于 Golden 均值且高于空白/损坏交付。首个 `invented-empty-description` 探针因 15 页接触表小字不足以让 VLM 稳定发现单页描述变异，评为 `1.0000`，已明确拒绝用于负例校准。
- 本轮 5 次 Golden 和 2 次负例探针的 VLM 调用均成功完成，没有基础设施失败或被错误计作 0 分的样本。
