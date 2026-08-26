本题一对一改造自 OmegaUse-OfficeVal `officeval_078`。

- 原始 Query、中英文 Rubric、上游 Verifier、10 页 PPT 和真实 CRM 工作簿已原样保留；输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题实际检查第 8 页全部 14 个原对象及右侧空白区域，并读取工作簿的 3 张工作表和记录维度后，设计了可编辑的 Excel 绿色圆角按钮。
- 按钮位于 x=27.0 cm、y=7.4 cm，尺寸 4.6 × 2.4 cm；不遮挡表格、标题、页脚和右下圆形。
- 超链接绑定在整个按钮形状上，目标严格为同目录相对文件名 `CRM_pipeline_restyled_reordered.xlsx`；随附工作簿与输入逐字节一致。
- 原 PPT 只有 `ppt/slides/slide8.xml` 和 `ppt/slides/_rels/slide8.xml.rels` 两个包成员变化，其他页面和原对象保持不变。

Golden 已经 LibreOffice 渲染为 10 页 PDF 并逐页检查；上游 Verifier 得分 10/10，本地独立 Judge 得分为 1.0。
