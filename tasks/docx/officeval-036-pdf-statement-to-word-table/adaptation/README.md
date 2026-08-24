# 逐题改造记录：OmegaUse-OfficeVal officeval_036

## 为什么选择这道题

- 原始 Query、两页 PDF 银行对账单和 4.3MB Word 模板都来自真实办公任务数据，不是本项目编写或合成的。
- 任务不只是把 PDF 文字复制到 Word，还必须分清 Withdrawals 和 Deposits、保持空白金额列、核对逐笔余额与期末合计，并保持 50 行统一行高和可编辑性。
- 上游记录人类完成时间为 180 分钟，是真正需要长上下文和细致自检的 `document_extraction` 任务。

## 原始来源

- 仓库：`baidu-frontier-research/OmegaUse-OfficeVal`
- 精确 commit：`ffbeecb8752447c8e40b594a0eeb1db7236ecb36`
- 上游任务：`data/tasks/officeval_036.json`
- 上游 Rubric：`data/rubrics/officeval_036.json`
- 上游 Verifier：`verifiers/officeval_036_verifier.py`
- 输入资产：Hugging Face `task_files/officeval_036/`

## 实际检查过的输入

- PDF 为 2 页 Letter 尺寸银行对账单，已查看两页渲染图，并逐行阅读可抽取文本。
- PDF 中共 48 行目标：1 行 `Balance forward`、46 笔收支、1 行 `Closing totals`。期末汇总为 Withdrawals `$54,891.73`、Deposits `$82,302.02`、Balance `$71,122.77`。
- Word 模板原有 3 个表格、一张 15–17 × 9–11 厘米图片和 2 行既有收支记录。结合上游“PDF 48 行”和“输出表共 50 行”两条 Rubric，可确认是保留 2 行模板记录后追加 48 行 PDF 记录。
- PDF SHA256 为 `34e4e7818a08dfe4adb6a57b3214c3ea62498d6ce09b3e22d27ce41fae70ed0b`；Word 模板 SHA256 为 `5293b5f65ef9d8ae4f9e86e9dd1b22da210be638fdd033fbba7a2dce71abdf24`。

## 真实标准答案是怎么来的

- `tests/expected-records.json` 不是编造的 Golden 数据：前 2 行直接来自 Word 模板，后 48 行按 PDF 两页的顺序逐行转录。
- 转录后独立汇总 46 笔交易的支出和收入，分别精确等于 PDF 期末的 `$54,891.73` 和 `$82,302.02`，证明没有把单笔金额放错列。
- 逐笔余额、首行 `$43,712.48` 和末行 `$71,122.77` 均与 PDF 一致。

## CoworkEvoBench 改造决策

- 原始 Query 逐字保留，只增加输入文件名、固定输出路径、模板既有 2 行的保留边界和可编辑性要求。
- 重建的仅是第一个收支记录表的数据行；模板图片字节、标题、其他表格和页脚都保留。
- 50 个数据行固定为 1.80 厘米，禁止跨页拆分；表头使用 Word 自动重复表头，没有在页面上贴截图。
- 上游 Verifier 原样保留；本项目 Judge 额外做 50 行逐字段对比、固定行高、模板图片 SHA256、渲染文字完整性和统一结果 schema 检查。
- `solution/solve.py` 只记录这道题已经人工转录的 48 行和这个 Word 模板的样式修改，不是通用 PDF 抽取器或批量制题脚本。

## 已完成的逐题验证

- Golden SHA256 为 `d975cb33a955a451e79c0565bfd150b3922e466f303a5be91a4a02e241b1d7b2`。
- 上游逐题 Verifier 的 8 个正向分项全部通过，三个罚项均未触发，得分 `18/18`；本项目扩展 Judge 得分为 `1.0`。
- `solve.sh` 从原始 Word 模板重新生成的独立 replay 再次得到上游 `18/18` 和本项目 Judge `1.0`；DOCX 内部时间元数据会使文件级 SHA256 不同，但 50 行事实、OOXML 结构与渲染检查全部一致。
- Golden 经 LibreOffice 实际渲染为 6 页，已逐页查看。前 5 页的 50 行文字和金额没有裁切，每页顶部自动重复表头；第 6 页保留原模板图片和空续表表头。
- 原始未作答模板得分 `0.0`，单笔余额错误得分 `0.428571`，10 行行高不一致得分 `0.714286`，缺失 5 行得分 `0.035714`，损坏 DOCX 和空交付均为 `0.0`；这些负例全部低于 `0.75` 通过阈值。
- Golden、replay 与六类负例的 Judge 输出均通过统一 `judge-result.schema.json` 校验。

## 尚未完成的发布门槛

- 需要由至少一个非参考 Agent Harness 独立作答并进行难度 Rollout。
- 正式发布前仍建议用 Microsoft Word 或 WPS 对 3.68 厘米列宽、Arial 替代字体和重复表头做一次兼容性抽检。
