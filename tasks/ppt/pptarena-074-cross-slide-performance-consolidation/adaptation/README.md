# PPTArena Case 74 逐题适配记录

本题逐页读取了 PPTArena Case 74 的 Query、`style_target`、原 5 页 PPTX 和上游 Ground Truth，没有用通用生成器批量编题。原第 2、3 页分别含 5 条 Team A / Team B 分数，原第 4 页是应删图表，原第 5 页含 3 张真实二进制图片和 3 条图注。

上游 Ground Truth 的事实和布局满足题意，因此 Judge 将其作为独立合法变体。本地 Golden 只对汇总表做可追溯精修：共同类别恢复原顺序 Manufacturing -> Finance -> IT；Other 行内 Team A 按 Logistics -> HR，Team B 按 Sales -> Marketing；全部统一为 en dash，并使用真段落分行，不改变任何分数或图片。

Judge 不比较 Golden 哈希。它从原数据独立编码 10 个 `团队×类别 -> 分数` 事实，既接受 Other 汇总行，也接受独有类别单独建行，同时拒绝漏项和串队。图片按二进制 SHA256 区分 Manufacturing / Finance / IT，再检查从左到右顺序、等尺寸、顶对齐、图注居中与空间关联。

## 实际验证

- 本地 Golden、ZIP 等价重封装和排版不同的上游答案均为 1.0；原 5 页未修改文稿、错分数、漏独有类别、串队、Title Case 失败、图片不等大/不对齐、图注未居中、删图、压平和答案标记均为 0。
- fresh replay 连续 5 次均为 1.0；Schema 和输入/Golden 资产校验通过。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `pptarena074-oracle-consolidation-20260829`，结果 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 本地 Golden 的 3 页已逐页渲染检查；表格文字清晰，三张图等大顶对齐，图注居中且无裁切。`visual_qa.py` 只报告原模板封面和标题占位符高度较大，该特征在冻结输入中已存在，本题不改封面版式。
