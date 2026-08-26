本题一对一改造自 OmegaUse-OfficeVal `officeval_060`。

- 原始 Query、中英文 Rubric、上游 Verifier 和 5 页源 PPTX 已原样保留，输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题实际打开并检查了全部 5 页，只修复第 4 页；第 1、2、3、5 页的渲染结果与源文件逐像素一致。
- 第 4 页背景图仍复用源 PPTX 中的同一媒体文件，通过上下对称裁剪适配 16:9 全出血页面，没有拉伸或重编码图片。
- 除 `ppt/slides/slide4.xml` 外，Golden 与源 PPTX 的全部 ZIP 包成员逐字节相同。

Golden 已经 LibreOffice 渲染为 5 页 PDF 并逐页检查；本地 Judge 得分为 1.0。
