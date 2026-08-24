# 逐题改造记录：OmegaUse-OfficeVal officeval_090

## 为什么选择这道题

- 原始任务来自经过脱敏和输入重建的真实办公请求，Query 和 61 工作表 XLSX 都不是本项目编写或合成的。
- 任务需要在 60 个子表中找到正确的“砼体积”和“总重量合计”，然后与汇总表逐行对齐；不能误取“砼重量”或某一条钢筋重量。
- 子表并不是一个模板的机械复制，实际存在多种布局。它能考查表格结构理解、取数、公式、保留性编辑和交付自检。

## 原始来源

- 仓库：`baidu-frontier-research/OmegaUse-OfficeVal`
- 精确 commit：`ffbeecb8752447c8e40b594a0eeb1db7236ecb36`
- 上游任务：`data/tasks/officeval_090.json`
- 上游 Rubric：`data/rubrics/officeval_090.json`
- 上游 Verifier：`verifiers/officeval_090_verifier.py`
- 输入资产：Hugging Face `task_files/officeval_090/子表数据提取到汇总表.xlsx`

## 实际检查过的输入

- 工作簿共 61 个工作表：1 个“汇总”和 60 个子表；子表为 Sheet1–Sheet10、Sheet12–Sheet61，原始就不存在 Sheet11。
- 汇总表 C3:C62 已给出逐行子表名称，F/G 两列为待填写目标，H 列原有静态数值需替换为公式。
- 逐表打开后确认，Sheet3、Sheet6–10、Sheet12–39、Sheet41 使用 D20/I10；Sheet5 和 Sheet42–61 使用 E21/I11。
- 四个异构表必须单独处理：Sheet1 使用 H19/H2，Sheet2 使用 E20/A10，Sheet4 使用 B19/B10，Sheet40 使用 F19/J6。
- 输入 SHA256 为 `fad1409b1b0d8cb9f7ef08a482d1c2102eef277b6ca0a68753e9b88ad2bc12ed`。

## CoworkEvoBench 改造决策

- 原始 Query 逐字保留，只增加输入文件名、固定输出路径、离线边界和未目标单元格保护要求。
- 对 56 个规则子表使用直接工作表引用，保留可追溯性；四个异构子表使用逐格核对后的数值。
- H3:H62 全部替换为可编辑公式，F/G 使用 `0.000`，H 使用 `0.00`。
- 上游 Verifier 对四个异构表的标签定位存在错误，例如它把 Sheet1 的砼体积源解析为文字标题 A18，而真实数值在 H19。本项目 Judge 依照实际单元格比对，不继承这个误判。
- `solution/solve.py` 是这道题的独立 replay 记录，其映射来自逐表检查，不是换个数据就能批量生成任务的工具。

## 已完成的逐题验证

- 参考解法可从原始 XLSX 独立 replay，当前 Golden SHA256 为 `656c3d62d86669344f9c005a1a9956570cafb6a1b9a516797831ccc11e0b6e97`。
- 上游逐题 Verifier 得分为 `25/25`，本项目扩展 Judge 得分为 `1.0`。
- Golden 经 LibreOffice 实际打开、重算和渲染；前六个打印页能读取汇总表首尾行的编号、砼体积、总重量和含量。
- 将 60 行 F/G/H 全部粘贴成正确静态数值的部分作答得分为 `0.708333`，低于 `0.75` 通过线，说明 Judge 能区分“值看起来对”与“保留正确取数公式”。
- 原始未作答工作簿得分为 `0.083333`；故意改错 10 行总重量的样本得分为 `0.208333`，并触发大面积取数错误罚项。
- 改动一个子表真实数值的样本得分为 `0.541667`，触发子表保护罚项；损坏文件和空目录都得分 `0`。
- Golden 与上述五类负例的 Judge 结果均通过 `schemas/judge-result.schema.json`。

## 尚未完成的发布门槛

- 需要由至少一个非参考 Agent Harness 独立作答并进行难度 Rollout。
- 正式发布前仍建议用 Microsoft Excel 或 WPS 对直接引用和打印分页做一次兼容性抽检。
