# 逐题适配记录

本题逐文件读取了 PPTArena Case 67 的原始 Query、`style_target`、原始 PPTX、上游 Ground Truth、来源 URL 和许可说明，没有用通用生成器批量编题。

上游答案有两处不宜直接作为本 Benchmark 真值：一是 6 个隐藏标题只是 `Slide N Title`，没有概括页面内容；二是第 2 页 `Idea 1`–`Idea 5` 表头仍保留 `DM Sans` 直接字体覆盖，与“所有文字使用主题默认字体”冲突。此外，第 3、4 页上游阅读顺序并未稳定遵循内容组的左到右顺序。

本地 Golden 以冻结上游文件为可追溯基线，逐页人工命名 6 个语义标题，并逐对象审计 6 页共 119 个原对象的逻辑顺序。字体处理只删除文字属性中的直接字体族节点，保留字号、粗体、斜体、颜色、对齐和其他可见格式；未新增业务事实或视觉素材。

Judge 不比较文件哈希，而是独立检查唯一标题与无障碍元数据、逐页逻辑顺序、主题 Arial 与零直接字体族覆盖、原对象 ID/几何/文本/非字体样式、媒体和备注，并执行 LibreOffice 渲染检查。上游未修正版被纳入负例，确保 Judge 会拒绝残留字体覆盖和泛化标题。

官方 Harbor 0.21 Oracle 容器回放 job 为 `pptarena067-oracle-accessibility-20260829`，trial 为 `pptarena-067-accessibility-readi__3X722BP`；奖励 1.0、0 异常。容器产物与本地 Golden 的 SHA256 均为 `78fdf673fddd833cd5b0e70f4ac2605c05693dd91b11254471b0bec43c729b6e`。
