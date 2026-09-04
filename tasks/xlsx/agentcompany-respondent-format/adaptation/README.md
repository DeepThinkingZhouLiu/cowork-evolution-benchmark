# C082 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=intermediate` 是 L2 作者假设，不是最终难度或发布决定；`calibration_status=uncalibrated`。

## 来源、输入和许可

- 来源仓库：TheAgentCompany，固定 commit `98b68ef82a47690c316f42fddb05baafaab56851`。
- 上游任务路径：`workspaces/tasks/ds-format-excel-sheets/`；原始 Query、检查点和 evaluator 分别保留为 `upstream-task.md`、`upstream-checkpoints.md`、`upstream-evaluator.py`。
- 上游输入 `respondent_counts.xlsx` 由 Git 在该 commit 直接跟踪；本地 solver-visible 副本重命名为 `respondent_counts_input.xlsx` 以避免与输出文件同名。SHA256：`be71871d79156a2106b085cf29a20e08ee39bf305c82022003f9533d33ffc043`，大小 8,598 bytes。
- 仓库根部 `LICENSE` 为 MIT，副本保留在 `LICENSE-TheAgentCompany`。仓库没有为该工作簿提供单独版权或数据条款；本候选将其按 MIT 仓库内直接跟踪的 synthetic benchmark fixture 记录，最终发布权利仍交由人工复核。
- 工作簿仅包含调查表文字和数值，无宏、公式、外部链接、嵌入媒体或个人记录；标题提及加拿大 respondent counts，但本题不把外部统计来源当作新增事实。

## 上游意图和改造边界

上游要求从 OwnCloud `Documents/Data Analysis/respondent_counts.xlsx` 的 `unformatted` 工作表创建 `formatted`，将指定顶部表头填充 `#87CEEB`，并使有值单元格水平居中，同时不修改 `unformatted`。本地任务保留同一工作流，把 OwnCloud 下载/上传副作用投影为固定输入和 `/app/output/respondent_counts.xlsx` 文件交付。

保留项：原始工作表内容、21 行 11 列尺寸、18 个合并区域、行高/列宽、字体、边框和输入填充；`formatted` 复制这些内容后只允许表头填充和水平对齐变化。目标填充坐标来自上游 evaluator 的明确列表：`A3、B3、B5、C3、C4、D4、D5、E5、F3、F4、G4、G5、H5、I3、I4、J4、J5、K5`。未添加外部数据、解释文本、排序、公式或隐藏要求。

## 独立真值和证据到要求映射

| 要求 | 独立证据 | Judge 检查 |
|---|---|---|
| `unformatted` 不变 | 输入工作簿本身的值、合并和样式 | R002 对输入逐格重算/比较 |
| `formatted` 内容复制 | 输入可见单元格和合并区域 | R003 |
| 指定表头 skyblue | 上游 Query/evaluator 的 18 坐标与 `#87CEEB` 规则 | R004 精确检查 `FF87CEEB` |
| 所有有值单元格居中 | 上游 Query 的“all cell values”要求 | R005 |
| 其它格式和可编辑性 | 输入样式、无宏/媒体事实和专业表格要求 | R006/R007 |
| 离线交付 | 原始 OwnCloud 文件副作用投影 | R001/R008 |

Judge 不读取 Golden 或上游期望文件；它从 solver-visible 输入重新比较值、结构和样式，并检查 LibreOffice 是否能渲染输出。目标单元格的背景色与非目标样式分开检查，避免把固定 Golden 当作唯一答案。

## 参考解法和验证

参考解法 `solution/solve.py` 使用 `openpyxl.copy_worksheet` 创建副本，仅修改 `formatted` 的非空单元格水平对齐和 18 个目标填充；输入哈希先验证。已运行：

```bash
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/xlsx/agentcompany-respondent-format
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/xlsx/agentcompany-respondent-format
```

正例与 fresh replay：`solution/solve.sh` 生成的 `respondent_counts.xlsx` 均由 `tests/test.sh` 判为 `1.000000`；Judge 逐项 R001–R008 全部通过。LibreOffice 7.3.7.2 无头打开并导出 PDF 成功，输出可读。

负例校准已覆盖：空输出/损坏文件=`0.000000`；把输入直接改名为输出/缺少 `formatted`=`0.230000`；单个目标颜色错误=`0.800000`；单个有值单元格未居中=`0.850000`；修改 `unformatted`=`0.800000`。后三类虽保留部分分数，但均因对应核心 hurdle 失败而判为未通过。Judge 重复运行结果确定性一致。

## 难度、泄漏和待审事项

- 作者假设：L2/intermediate。一个输入工作簿、一次工作表复制、18 个精确表头样式目标和全表对齐/结构保持构成中等单应用格式化任务；预计熟练办公人员 20 分钟内完成。大行数或外语不作为难度依据。
- leakage group：`tac-respondent-format`。不得与同一 TheAgentCompany `respondent_counts.xlsx`、同一格式化模板或同一来源 fixture 的变体跨切分。
- 待人工决定：MIT 对仓库内二进制 fixture 的再发布范围、是否接受 OwnCloud 到本地文件的适配、最终难度和是否纳入正式 60 题。Codex 不标记接受、校准或发布。
