# R035 适配记录：OmegaUse-OfficeVal officeval_088

状态：`pending_human_review` / `uncalibrated`。作者初始难度假设为 L3/advanced。

## 来源与证据

- 固定仓库 commit：`ffbeecb8752447c8e40b594a0eeb1db7236ecb36`。
- 上游任务：`officeval_088`；原始任务 JSON 和 Rubric 保存在 `adaptation/`。
- 输入为五份产品说明 DOCX、一个 Excel 模板和一个 Word 模板；逐文件哈希见 `metadata/input-manifest.json`。
- 输入图片均从对应 DOCX 的唯一 `word/media` 原图裁剪；不使用网络图片。
- 许可文件保存在 `adaptation/LICENSE-OmegaUse-OfficeVal`；当前仅为受限候选，需人审确认资产发布权利。

## 输出压缩

上游原任务要求 Excel 和 Word 两份输出。本 Excel benchmark 版本将同一自然工作流压缩到一个 `officeval_088_product_table.xlsx`：`Products` 工作表保存五行产品事实、来源文件和五张原生嵌入图片；`Top3 Analysis` 工作表用跨表公式按指导价格引用最高三项。没有删掉 Word 输出名后假装完成，而是明确把高价分析作为同一 XLSX 的第二工作表。

## 独立事实与图像审计

产品名称、型号、指导价格和参数从五份 DOCX 逐项提取；价格排序为云溪 4500、统帅 3999、小天鹅 3039、米家 999、美的 846。五个裁剪框在逐张源图检查后固定，排除海报营销文案。Golden 由 `solution/solve.py` 从可见输入重建，不复制上游 Word/Excel Golden。

## Judge 与验证

Judge 独立检查五行事实、Top3 排序/跨表公式、五张图片锚点和像素裁剪、原生公式、工作表结构、输入哈希和 LibreOffice 渲染；额外非 XLSX 文件与未修改模板会失败或扣分。当前已完成包结构和参考解法，待执行完整 Golden/fresh replay、图像变异和 Claude Code + Sonnet 5 rollout。

## 未关闭风险

- OfficeVal 题内素材的衍生发布范围需人工确认。
- 图片裁剪使用了逐题审核的固定坐标，若源文档版本变化必须重新计算哈希并复核。
- `Top3 Analysis` 的高价摘要已压缩为 XLSX，不再交 Word；这属于用户指定 Excel-only 合同，需要人审确认语义足够保真。
