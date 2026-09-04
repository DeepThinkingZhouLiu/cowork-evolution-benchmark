# 改造说明

本题逐题改造自 OmegaUse-OfficeVal 的 `officeval_055`，精确来源版本为 `ffbeecb8752447c8e40b594a0eeb1db7236ecb36`。

- `upstream-task.json`：上游真实 Query 和 PNG 地址，原样保留。
- `upstream-rubric.json`：上游逐项评分要求，原样保留。
- `upstream-verifier.py`：上游逐题 Verifier，原样保留；`../tests/` 中亦保存了可离线执行的副本。
- `LICENSE-OmegaUse-OfficeVal`：上游 Apache-2.0 许可证。

本 Benchmark 没有自编 Query 或输入数据。真实输入 `website_outage_editable_flow.png` 来自上游公开数据，SHA256 为 `80bca602124e0cd54203937d2d46817062e7c61849cc05c911e64345be1c6add`。本地只去掉了上游文件名在扩展名前的一个空格，文件内容未变。

Golden 是针对这一道真实题的逐对象改造结果。我先人工查看原图，核对五个阶段的编号、图标、标题、正文、箭头、情绪折线和底部主题，再按上游 Query 把第二阶段改为 `check`，并替换 `finish` 正文。交付物中没有嵌入原图；每个文字框、形状、图标部件、箭头和折线都可单独编辑。

`solve.py` 只复现本题的人工改造结果：它先校验这一个真实输入的 SHA256，然后生成本题固定的可编辑交付物，不是通用制题器，也不批量 mock 其他题。

上游 Verifier 有两个已确认的实现问题。第一，它在计数四条主 action 箭头时没有排除左上角图例中的黑色箭头，因此会得到五条并误判；第二，它仅用 OOXML 连接线的正数 `ext.cy` 判断上下方向，却忽略了真正表示向上折线的 `flipV`。原始 Verifier 保留不改；本地 Judge 排除图例区域，并按 `flipV` 判断视觉方向。

参考交付物已通过 LibreOffice 实际渲染、人工视觉检查、Golden 正例、错误文本、缺失阶段、缺失箭头、缺失 emotion 折线、空白 PPTX 和整页截图等负例验收。

## 视觉评分适配（visual-v1）

- 模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、`tests/expectations.json`、完整 `tests/judge.py`、输入清单、来源记录、上游 Query/Rubric、本题改造说明和整份专用 `solve.py`。
- 输入 `website_outage_editable_flow.png` 已按原始分辨率查看：五列分别是发现停机、检查确认、故障高峰、恢复方案和结束，页面同时用四条黑色 action 箭头与六段橙色 emotion 虚线讲叙事件和情绪两条进程。Golden 已真实渲染为 16:9 单页图像，并目视检查全部 63 个可编辑对象的可见结果。
- 这是对指定图像的精确重建与文本定点修改，因此本题使用 `reference`；VLM 同时查看候选与 Golden，但不得把上游图中已被用户明确要求替换的 `assessment` 或旧 finish 段落当成正确答案。
- `V001` 检查五列四层节点的宏观布局；`V002` 检查五个可编辑图标、告警/恢复徽章和四条 action 箭头；`V003` 检查五个标题条与正文框，特别是 check 和更新后的 finish；`V004` 检查左上角双线型图例；`V005` 检查六段 emotion 起伏、四个情绪词和主题；`V006` 检查整页双线叙事的字体、卡片、线条、阴影和配色一致性。精确文本、对象数、可编辑性与页面尺寸继续由确定性 Judge 负责。
- Golden 的可见风险包括：左上图例之外的页面上方留白较大；第 5 列的长 finish 段落字号小于其余列；concerned 和 hopeful 两个斜体标签与橙色虚线较近，局部接触。VLM 必须按这些实际渲染现象评分，不得因其为 Golden 自动给满分。
- Golden 保留确定性得分 `1.0`；Terra 连续 5 次视觉得分均为 `1.0000`，均值 `1.0000`，总体标准差 `0.0000`，满足 `std < 0.05`。
- 中等可见负例只改动本题的指定对象：将第 3 列 `problem peak` 标题条下移到正文框后方，造成标题可见缺失；将第 4 列浅蓝图标圆下移，使云端图标与圆心分离；同时上移蓝色 theme 文字，压缩它与 emotion 路径的间隔。文件仍可正常打开，五列主体和双线叙事仍可识别；该负例视觉得分 `0.7716`，严格低于 Golden 均值且高于空白/损坏产物；70/30 合成得分为 `0.9315`。
- 校准证据保存在忽略目录 `.coworkevo/vlm-calibration/ppt/officeval-055-image-to-editable-outage-flow/`，包含 5 次 Golden、候选/参考渲染图和中等负例结果。
