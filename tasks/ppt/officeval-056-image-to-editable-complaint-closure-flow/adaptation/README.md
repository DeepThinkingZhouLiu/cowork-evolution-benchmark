本题一对一改造自 OmegaUse-OfficeVal `officeval_056`。

- 原始 Query、中英文 Rubric、上游 Verifier 和参考 PNG 已原样保留，输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题在实际阅读参考图后，逐对象重建上方主流程、重大问题判断、下方整改链、复核判断、结案回流、返工回路和相关记录框。
- 耳机、文档、柱状图、人物、文件夹、放大镜、清单、齿轮、盾牌和剪贴板均由多个原生形状组成。
- 所有 26 个可编辑文本段的 OOXML 字体均设为“思源手写体”。渲染环境如未安装该字体，LibreOffice 会使用本地回退字体，但 PPTX 中指定的字体名保留不变。
- 交付页面共含 107 个原生可编辑对象，不含任何位图。

Golden 已经 LibreOffice 渲染为 1 页 PDF 并目视检查；本地 Judge 得分为 1.0。
