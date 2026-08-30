# OfficeVal 093 改造记录

来源为 OmegaUse-OfficeVal 固定提交 `ffbeecb8752447c8e40b594a0eeb1db7236ecb36` 的 `officeval_093` 任务、Rubric 和 verifier；上游任务要求在观察表 `Sheet0!C32` 按参考图加入两个可编辑自由五边形。输入工作簿和参考 PNG 已人工查看并记录 SHA256；许可证与原 verifier 保存在本目录。

本题采用直接资产改造。用户意图、工作表内容和图形几何要求均保留；外部图片只作为可见参考，不进入交付物。参考解法在 OOXML DrawingML 中创建一个 C32 锚定的图形组，组内包含两个独立 `custGeom` 五边形，使用无填充深灰色实线并保留原单元格内容。Judge 调用保留的几何解析器检查输入可打开性、自由多边形、镜像交叠、尺寸、轮廓、填充和文本可读性，未比较 Golden 哈希。

作者假设为 L1/basic，状态保持 `uncalibrated`、`pending_human_review`。验证命令：`python solution/solve.py --input-dir data/input_files --output-dir solution`，随后以 `tests/judge.py --artifact-dir solution` 检查；还需在人工作业环境中确认 Excel/LibreOffice 对 DrawingML 组内自由形状的可视编辑一致性。
