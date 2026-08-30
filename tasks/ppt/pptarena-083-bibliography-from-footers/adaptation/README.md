# PPTArena Case 83 逐题适配记录

本题逐页读取了 PPTArena Case 83 的 Query、`style_target`、原 23 页 PPTX 和上游 Ground Truth，没有用通用生成器批量编题。

原文稿在第 4、5、6、7、11、12、13、14、15、16、18、19 和 21 页出现来源文字；其中 Office、Vahidy、Price-Haywood 和 Patrick 重复出现，去重后得到 9 条。上游 Ground Truth 在最终 Thank you 页前插入 Bibliography，按 Agarwal、Ali、Garg、Office、Patrick、Price-Haywood、Richardson、Vahidy、Wosik 排序，并保留了 5 处重要斜体。本地逐页审计认定该答案正确，因此不重新编造引用或内容。

Judge 不比较 Golden 文件哈希，而是从冻结输入独立检查 9 条准确文本、去重、作者排序、段落分隔、5 处斜体和 13 处原页脚引用的完整保留；同时检查原 23 页对象、备注、15 个媒体、新页位置、可编辑性和渲染页数。

## 实际验证

- 本地 Golden、ZIP 等价重封装和上游答案均为 1.0；原稿未修改、漏引用、重复引用、错序、丢失斜体、错位插页、改/删原页脚、改原页、删媒体、压平和答案标记均为 0；多一个调试文件的交付为 0.96。
- fresh replay 连续 5 次均为 1.0；单题结构、输入/Golden 资产和相关单元测试均通过。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `pptarena083-oracle-bibliography-20260829`，trial 为 `pptarena-083-bibliography-from-f__2n9vxyK`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 24 页已渲染为接触表并逐页目视；Bibliography 文字清晰、无裁切，倒数第 2 页位置正确。`visual_qa.py` 只报告原模板存在的空闲文本框；第 2 页为上游原稿的动画分步大纲，LibreOffice 静态渲染会同时显示各动画状态而重叠，该现象在冻结输入中已存在，本题根据“原页不改”约束不擅自修复。
