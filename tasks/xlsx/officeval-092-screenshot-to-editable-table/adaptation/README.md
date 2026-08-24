# 逐题改造记录：OmegaUse-OfficeVal officeval_092

## 为什么选择这道题

- 原始任务来自经过脱敏的真实办公请求，不是本项目自编业务场景。
- 原始 Query、输入 XLSX、Rubric 和逐题 Verifier 都可以追溯。
- 任务同时考查截图理解、表格恢复、公式、格式保持和打印分页，适合 `document_to_table` 家族。
- 上游记录的人类完成时间为 101 分钟，具备真实工作量信号。

## 原始来源

- 仓库：`baidu-frontier-research/OmegaUse-OfficeVal`
- 精确 commit：`ffbeecb8752447c8e40b594a0eeb1db7236ecb36`
- 上游任务：`data/tasks/officeval_092.json`
- 上游 Rubric：`data/rubrics/officeval_092.json`
- 上游 Verifier：`verifiers/officeval_092_verifier.py`
- 输入资产：Hugging Face `task_files/officeval_092/智慧农业示范园建设项目报价表.xlsx`

以上文本和许可证均在本目录或 `tests/upstream_verifier.py` 中保留原始副本。

## 实际检查过的输入

- 工作簿包含“智慧农业报价表”和“费用类别”两个工作表。
- 主工作表为 125 行、10 列，内嵌一张 `1274×523` JPEG 截图。
- 截图内容是“五、环控与照明物资清单”和“六、物联网设备预算清单”。
- 原文件经 LibreOffice 打印为 9 页，截图仍作为不可编辑对象存在，多张表格被分页切断。
- 下载文件 SHA256 为 `1f09aeac32d262643acb6976cb9540e31cc2602da7eb1716bb2d1f18ae66e959`。

## CoworkEvoBench 改造内容

- 原始 Query 不改写，只补充输入文件名、固定输出文件名、离线要求和不得破坏无关内容的边界。
- 直接使用原始 XLSX，不重新制造输入数据。
- Golden 逐格恢复截图中的两张表，使用可重算公式，并保留其他工作表和业务内容。
- 上游 Verifier 原样保留；Harbor Judge 继续复用其单元格检查，并补充其当前未覆盖的打印渲染、行高、截图删除和辅助表保留检查。
- 本题具有独立的逐题参考解法，不共享任何批量题目生成器。

## 已完成的逐题验证

- 参考解法从原始输入独立 replay 成功，Golden SHA256 为 `769ae967b74c32e60abc031c471d7c030f0231bf30f2c14248e8e127f20334a4`。
- 上游逐题 Verifier 得分为 `13/13`；本项目扩展 Judge 得分为 `1.0`。
- Golden 通过 LibreOffice 实际导出为 13 页 A4 PDF，并逐页人工检查。检查时曾发现汇总页数值显示为 `###`，修正列宽后再次渲染，全部数值和两张恢复表均可见，且没有表格跨页。
- 原始未作答文件作为错误交付时得分为 `0.057143`，不能通过；空目录和损坏 XLSX 得分均为 `0`，不能通过。
- 正例、空目录、原始未作答文件和损坏文件产生的 Judge 结果均通过 `schemas/judge-result.schema.json`。
- `metadata/task.json`、`metadata/provenance.json`、`rubric.json`、`task.toml`、输入文件大小和 SHA256 均完成结构或一致性校验。

## 尚未完成的发布门槛

- 需要用至少一个非参考 Agent Harness 独立作答，确认 Query 对通用 Agent 可解，而不是只对参考解法可解。
- 需要在真实 Agent 输出中选取“部分完成”的中等缺陷样本，继续校准 Judge 的分数梯度。当前空白、未作答和损坏负例只能证明拒绝下界，不能替代难度校准。
- `openpyxl` 保存时会移除它不支持的扩展和条件格式扩展；当前逐页渲染未见可见损坏，但正式发布前仍需用 Excel 或 WPS 再做一次人工兼容性抽检。
