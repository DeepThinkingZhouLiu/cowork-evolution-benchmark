# TheAgentCompany GitLab Engineering Progress 逐题适配记录

本题直接阅读了上游 `pm-present-gitlab-info-as-ppt` 的完整 Query、checkpoints、evaluator 和参考答案，从 evaluator 中逐字段冻结 14 个 GitLab 仓库的顺序、描述、Issue 总数和 Merge Request 总数，没有自编仓库或业务数字。`Original_Request.md` 是上游 task.md 的逐字节副本。

逐题审计发现上游 evaluator 有三个会造成假阳性的缺口：名称、描述、Issue 和 MR 每类只要超过一半仓库正确即可得分；数字允许 `+/-5`；同一仓库的四个字段可以从不同页面分别命中。本 Harbor 适配改为固定 15 页，14 个仓库按页序一一对齐，描述和两类数值逐页精确核对。对上游为空的三个描述，显式要求标记元数据缺口，不允许脑补。

## 实际验证

- Golden、ZIP 等价重封装和非语义封面变体均为 1.0。损坏/空交付、仓库页调序、漏仓库页、描述错误、空描述被脑补、Issue 只错 1、MR 只错 1、封面总数错、计数口径错、单页误称 open only、来源缺失、整页压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，PPTX SHA256 稳定为 `ad38f21f0a2be3b6295e6911bf5e583963333b1cb661457864ece532d9790771`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `theagentcompany-gitlab-progress-oracle-20260829`，trial 为 `theagentcompany-gitlab-engineeri__cvuWnhs`；奖励 1.0、0 异常，容器产物哈希与本地 Golden 一致。
- 15 页均已通过 LibreOffice 渲染；封面、空描述页、高数值页和最长描述页已逐页目视，无重叠、裁切或整页压平。
