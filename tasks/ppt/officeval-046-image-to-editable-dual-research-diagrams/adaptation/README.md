本题一对一改造自 OmegaUse-OfficeVal `officeval_046`。

- 原始 Query、中英文 Rubric、上游 Verifier 和两张 PNG 已原样保留，输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题 Golden 是在实际阅读两张图后逐对象重建的 30×20 cm 双页 PPTX。
- 第 1 页的箱体、顶部分栏、左侧指标、主流程、两个控制模块、膜分离设备与结果栏均由原生对象组成。
- 第 2 页的四阶段、侧边模块、中心风险等级、8 个周边指标、条形图、三糴折线图与底部输出也均为原生对象。
- 交付 PPTX 不含任何位图；两页分别含 132 和 108 个可编辑原生对象。
- 上游 Verifier 对 OOXML 对象分类和启发式命名强依赖，对该等价可编辑重建出现大量漏判；Harbor Judge 因此独立检查内容、层级、对象数、几何关系和真实渲染。

Golden 已经 LibreOffice 渲染为 2 页 PDF 并逐页目视检查；本地 Judge 得分为 1.0。
