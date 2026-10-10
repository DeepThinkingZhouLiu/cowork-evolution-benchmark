# PPTArena Case 74 逐题适配记录

本题逐页读取了 PPTArena Case 74 的 Query、`style_target`、原 5 页 PPTX 和上游 Ground Truth，没有用通用生成器批量编题。原第 2、3 页分别含 5 条 Team A / Team B 分数，原第 4 页是应删图表，原第 5 页含 3 张真实二进制图片和 3 条图注。

上游 Ground Truth 的事实和布局满足题意，因此 Judge 将其作为独立合法变体。本地 Golden 只对汇总表做可追溯精修：共同类别恢复原顺序 Manufacturing -> Finance -> IT；Other 行内 Team A 按 Logistics -> HR，Team B 按 Sales -> Marketing；全部统一为 en dash，并使用真段落分行，不改变任何分数或图片。

Judge 不比较 Golden 哈希。它从原数据独立编码 10 个 `团队×类别 -> 分数` 事实，既接受 Other 汇总行，也接受独有类别单独建行，同时拒绝漏项和串队。图片按二进制 SHA256 区分 Manufacturing / Finance / IT，再检查从左到右顺序、等尺寸、顶对齐、图注居中与空间关联。

## 实际验证

- 本地 Golden、ZIP 等价重封装和排版不同的上游答案均为 1.0；原 5 页未修改文稿、错分数、漏独有类别、串队、Title Case 失败、图片不等大/不对齐、图注未居中、删图、压平和答案标记均为 0。
- fresh replay 连续 5 次均为 1.0；Schema 和输入/Golden 资产校验通过。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `pptarena074-oracle-consolidation-20260829`，结果 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 本地 Golden 的 3 页已逐页渲染检查；表格文字清晰，三张图等大顶对齐，图注居中且无裁切。`visual_qa.py` 只报告原模板封面和标题占位符高度较大，该特征在冻结输入中已存在，本题不改封面版式。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本适配记录、上游 Case 与元数据、`build_golden.py` / `build_fixtures.py` / `solve.py`、冻结输入、上游 Ground Truth 和本地 Golden；核对了两队 10 个团队×类别事实、三页删并关系、三张媒体哈希/顺序与三条图注。
- 冻结输入 5 页、本地 Golden 3 页和 `misaligned-image` 负例均已原尺寸渲染并逐页查看：Golden 的封面副标题改为 Title Case，第 2 页使用蓝灰 5×3 原生表格整合两队，第 3 页三图按 Manufacturing / Finance / IT 等大顶对齐且图注居中；无内容裁切或附带页。
- 本题要求精确删页、表格整合与图片保持，因此使用 `reference`。专属视觉 Rubric 检查三页叙事、摘要表事实的可读分组、表格版式、三图顺序/顶线/裁切、图注关联和全稿一致性；精确事实、媒体哈希、图片几何、可编辑性与输出卫生继续由确定性 Judge 负责。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、the configured VLM provider 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；对应综合分均为 `1.0000`。
- 中等可见负例使用本题逐项构造的 `misaligned-image`：只将第 3 页 IT 图片下移 0.4 英寸，使其顶线明显低于前两图并覆盖 `IT: Data Servers` 图注；封面、摘要表、另外两图和所有文字仍在，文件可正常打开，不属于空白或损坏交付。其确定性分 `0.0000`、视觉分 `0.7600`、综合分 `0.2280`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮 6 次 VLM 调用均成功完成，没有基础设施失败或被错误计作 0 分的样本。
