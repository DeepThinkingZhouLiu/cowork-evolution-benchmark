# PPTArena Case 83 逐题适配记录

本题逐页读取了 PPTArena Case 83 的 Query、`style_target`、原 23 页 PPTX 和上游 Ground Truth，没有用通用生成器批量编题。

原文稿在第 4、5、6、7、11、12、13、14、15、16、18、19 和 21 页出现来源文字；其中 Office、Vahidy、Price-Haywood 和 Patrick 重复出现，去重后得到 9 条。上游 Ground Truth 在最终 Thank you 页前插入 Bibliography，按 Agarwal、Ali、Garg、Office、Patrick、Price-Haywood、Richardson、Vahidy、Wosik 排序，并保留了 5 处重要斜体。本地逐页审计认定该答案正确，因此不重新编造引用或内容。

Judge 不比较 Golden 文件哈希，而是从冻结输入独立检查 9 条准确文本、去重、作者排序、段落分隔、5 处斜体和 13 处原页脚引用的完整保留；同时检查原 23 页对象、备注、15 个媒体、新页位置、可编辑性和渲染页数。

## 实际验证

- 本地 Golden、ZIP 等价重封装和上游答案均为 1.0；原稿未修改、漏引用、重复引用、错序、丢失斜体、错位插页、改/删原页脚、改原页、删媒体、压平和答案标记均为 0；多一个调试文件的交付为 0.96。
- fresh replay 连续 5 次均为 1.0；单题结构、输入/Golden 资产和相关单元测试均通过。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `pptarena083-oracle-bibliography-20260829`，trial 为 `pptarena-083-bibliography-from-f__2n9vxyK`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 24 页已渲染为接触表并逐页目视；Bibliography 文字清晰、无裁切，倒数第 2 页位置正确。`visual_qa.py` 只报告原模板存在的空闲文本框；第 2 页为上游原稿的动画分步大纲，LibreOffice 静态渲染会同时显示各动画状态而重叠，该现象在冻结输入中已存在，本题根据“原页不改”约束不擅自修复。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本适配记录、上游 Case 与元数据、`build_golden.py` / `build_fixtures.py` / `solve.py`、冻结输入、上游 Ground Truth 和本地 Golden；逐页核对了 13 处原页脚、9 条去重引用、指定作者/机构顺序和 5 个斜体词组。
- 冻结输入 23 页、本地 Golden 24 页和 `missing-citation` 负例均已原尺寸渲染为接触表并逐页查看，另单独放大检查第 23 页：Golden 的 9 条蓝色项目在标题和底部品牌区之间完整可读，Agarwal→Wosik 排序与斜体层级清楚；Bibliography 位于 Summary 和 Thank you 之间。原第 2 页动画静态叠字在输入和 Golden 中相同，明确不计为本题新增缺陷。
- 本题要求精确引用文本、插页位置和原稿保持，因此使用 `reference`。视觉 Rubric 检查末页顺序、九条作者前缀与独立项目、长引用换行/可读性、五处斜体、UCLA 蓝橙品牌体系和原稿视觉保持；逐字文本、段落对象、字体属性、页脚签名、媒体哈希和原生可编辑性继续由确定性 Judge 负责。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；对应综合分均为 `1.0000`。
- 中等可见负例使用本题逐项构造的 `missing-citation`：仅删除 Bibliography 第 3 条 Garg 文献，使页面剩 8 个独立项目并在 Ali 与 Office 之间形成缺口；其余 23 页、页面品牌、斜体和结束页均保留，文件可正常打开，不属于空白或损坏交付。其确定性分 `0.0000`、视觉分 `0.7950`、综合分 `0.2385`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮 6 次 VLM 调用均成功完成，没有基础设施失败或被错误计作 0 分的样本。
