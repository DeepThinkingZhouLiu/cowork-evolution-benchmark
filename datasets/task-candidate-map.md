# 42 个任务家族的 57 道真实候选映射

## 如何读这张表

这是 42 个家族的人工选题与改造记录。目前全部家族均已至少有 1 道 `已完成基础验证` 的端到端任务；首轮 43 道完成家族覆盖，本轮又从 OmegaUse-OfficeVal 逐题增加 14 道异质题，共 57 道。一个上游任务原则上只归入一个主家族，避免为凑覆盖重复包装同一道题。

状态含义：

- `已完成基础验证`：逐题检查、改造、Golden replay、视觉检查和基础负例已经完成，仍待非参考 Agent Rollout。
- `待逐题审查`：只完成元数据级筛选，尚未下载或完整检查输入。
- `需静态化`：上游依赖数据库、网页、邮件、日历或其他动态服务，需要先固化离线输入。
- `许可待确认`：仓库或单个附件的再发布权限还不清楚，不能进入公开任务包。
- `缺口`：参考区还没有足够贴合的候选，完成已知来源后再按要求进行 Web Search。

## 演示文稿

| 任务家族                    | 首选上游候选                                             | 状态             | 选择理由或待确认点                                  |
|-----------------------------|----------------------------------------------------------|------------------|-----------------------------------------------------|
| `material_to_deck`          | OmegaUse `officeval_055`                                 | 已完成基础验证   | 真实参考 PNG 逐对象重建为可编辑流程图；0 嵌入图片，Golden/replay Judge `1.0` |
| `outline_to_deck`           | OmegaUse `officeval_076`                                 | 已完成基础验证   | 真实长 Query + 模板生成 21 页可编辑 Deck，9 页动画主序列，本地 Judge `1.0` |
| `data_to_deck`              | OmegaUse `officeval_066`                                 | 已完成基础验证   | 真实 Excel 追加可编辑表格与折线图，上游 `22/22`       |
| `template_based_generation` | OmegaUse `officeval_075`                                 | 已完成基础验证   | 5 页参考 Deck 逐角色迁移到 39 页课件，本地 Judge `1.0` |
| `content_editing`           | OmegaUse `officeval_063`                                 | 已完成基础验证   | 4 张真实标注图指导多页编辑和原生双轴图表，本地 Judge `1.0` |
| `visual_editing`            | OmegaUse `officeval_069`                                 | 已完成基础验证   | 真实标注截图指定 10 个局部框体样式，上游 `30/30`       |
| `deck_polishing`            | OmegaUse `officeval_070`                                 | 已完成基础验证   | 单页拥挤文字的布局和可读性修复                       |
| `deck_consolidation`        | OmegaUse `officeval_068`                                 | 已完成基础验证   | 10 页 PDF + 删页后 15 页 PPT 合并为全可编辑 25 页，本地 Judge `1.0` |
| `audience_adaptation`       | Cowork-Bench `sf-department-budget-ppt`                   | 已完成基础验证   | 上游 PDF + 50,000 行 HR COPY 完整静态化；面向领导层重组为 10 页 16:9 可编辑汇报，Golden/fresh/变异 replay 均 `1.0` |
| `localization`              | OmegaUse `officeval_074`                                 | 已完成基础验证   | 46 页、507 个真实中文对象逐一英文化，图片、背景和版式保持，本地 Judge `1.0` |
| `presentation_package`      | Cowork-Bench `ppt-clickhouse-executive`                  | 已完成基础验证   | 2,579 条 Q4 订单 + 1,451 条客户维度静态化；Excel、7 页可编辑 PPTX 与同版 PDF 跨文件 Judge `1.0` |
| `deck_audit`                | OmegaUse `officeval_041`                                 | 已完成基础验证   | 23 页逐页备注、3 张真实图片、Excel 表格与动画跨附件修订，本地 Judge `1.0` |

## 文字文档

| 任务家族                  | 首选上游候选                                             | 状态             | 选择理由或待确认点                                  |
|---------------------------|----------------------------------------------------------|------------------|-----------------------------------------------------|
| `material_to_report`      | FORTE `Legal-020`                                        | 已完成基础验证   | 规则 + 6 份委托书逐项核查，八条上游真值全部落实；Golden/replay Judge `1.0` |
| `outline_to_document`     | Cowork-Bench `rzd-msk-spb-team-trip-gcal-word`           | 已完成基础验证   | 上游 8 个 RZD 车次 + 三档票价 + 会议约束静态化；原始 Golden/replay Judge `1.0` |
| `template_filling`        | OmegaUse `officeval_006`                                 | 已完成基础验证   | 5 页 PDF 的 21 道题与答案逐页录入，63 个分数和 14 个根式为 OMML，上游 `11/11`、本地 Judge `1.0` |
| `editing_rewriting`       | LLM-DOCX-Editing `D3`                                    | 已完成基础验证   | 五项成交条件面向非法律读者改写；canonical 与异文均 Judge `1.0` |
| `redlining_review`        | LLM-DOCX-Editing `F6`                                    | 已完成基础验证   | 5 个真实 `w:ins` + 4 个 `w:del`；Golden/replay/独立 Claude 输出 Judge `1.0` |
| `formatting_fidelity`     | OmegaUse `officeval_018` + `officeval_022`               | 已完成基础验证   | 保留 2 道异质题：长报告格式统一，以及封面/目录样式跨文档迁移；均完成 replay 和逐页检查 |
| `merge_compare`           | LLM-DOCX-Editing `S4`                                    | 已完成基础验证   | 上游公开通过结果作 Golden；7 段更新、25 段保持，Golden/replay Judge `1.0` |
| `document_extraction`     | OmegaUse `officeval_036`                                 | 已完成基础验证   | PDF 48 行加模板 2 行逐项核对，上游 `18/18`           |

## 电子表格

| 任务家族                       | 首选上游候选                                             | 状态                 | 选择理由或待确认点                              |
|--------------------------------|----------------------------------------------------------|----------------------|-------------------------------------------------|
| `document_to_table`            | OmegaUse `officeval_092`                                 | 已完成基础验证       | 真实截图恢复、公式、排版和打印分页              |
| `data_cleaning`                | Cowork-Bench `fetch-terminal-data-pipeline`               | 已完成基础验证       | 上游 20 条销售 + 7 条库存 API JSON 原样保留；按主键保留首条、商品名规范、三表 Excel 与离线 Dashboard Golden/replay Judge `1.0` |
| `table_transformation`         | OmegaUse `officeval_090`                                 | 已完成基础验证       | 60 子表逐表确认取数位置，上游 `25/25`            |
| `formula_authoring`            | OmegaUse `officeval_082`                                 | 已完成基础验证       | 126 个逐格双态控件、380 个公式、两组 LibreOffice 行为回放，上游 `16/16` |
| `model_building`               | Cowork-Bench `sf-regional-sales-forecast`                | 已完成基础验证       | 上游 2,000 客户+20,000 订单聚合为 15 条历史；Excel+Word Golden/replay Judge `1.0` |
| `scenario_analysis`            | Cowork-Bench `insales-inventory-reorder-forecast`        | 已完成基础验证       | 上游 DB 投影 82 产品+150 订单；77 行分析与 18 行补货逐格对齐 Golden，replay Judge `1.0` |
| `pivot_analysis`               | OmegaUse `officeval_091`                                 | 已完成基础验证       | 销售单取数、kg 转吨、57 客户六类汇总，上游 `9/9` |
| `charting`                     | OmegaUse `officeval_086`                                 | 已完成基础验证       | 基于原表制作可编辑的复合饼图                    |
| `dashboard`                    | Cowork-Bench `moex-sector-rotation-dashboard`            | 已完成基础验证       | 10 个价格 + 5 份年报 + 门户 HTML 已静态化；Golden/replay Judge `1.0`，但上游无图表且打印排版较弱 |
| `template_completion`          | OmegaUse `officeval_087`                                 | 已完成基础验证       | 两个真实工作簿逐项迁移，上游 `36/36`            |
| `spreadsheet_repair`           | SpreadsheetBench-2 公开样例 `Debugging/09_01`     | 已完成基础验证       | 选用单独声明 CC-BY-SA-4.0 的数据集 revision；人工核对 4 类、29 个错误公式，Golden/fresh/新增同类错误 replay 均 `1.0`；未复制无许可的代码仓文件 |
| `spreadsheet_audit`            | Cowork-Bench `sf-sla-breach-audit`                       | 已完成基础验证       | 上游 31,588 条工单 + 4 条 SLA 策略完整静态化；Excel 与离线邮件草稿 Golden/replay Judge `1.0` |
| `multi_workbook_integration`   | FORTE `Finance-018`                                      | 已完成基础验证       | 原 Query、3 份往来明细与 2 份上游 Golden 原样保留；31 笔未付、2 笔未收、0 僵尸账，内容变异 replay 可动态识别新增款项与僵尸账 |
| `analysis_to_report`           | Cowork-Bench `customer-satisfaction-analytics-framework` | 已完成基础验证       | 上游 50 份问卷 + 20 条工单；修正原 Golden 虚构的 245/847 样本量与无来源推断，4 份交付 Golden/replay Judge `1.0` |

## 跨办公应用

| 任务家族                     | 首选上游候选                                                  | 状态         | 选择理由或待确认点                                |
|------------------------------|---------------------------------------------------------------|--------------|---------------------------------------------------|
| `extract_analyze_present`    | Cowork-Bench `pw-scholarly-trend-analysis-excel-ppt`           | 已完成基础验证 | 上游 4 篇论文 + HTML 面板静态化；处理脚本、JSON、Excel 和 5 页 PPT Golden/replay Judge `1.0` |
| `research_to_deliverables`   | Cowork-Bench `fetch-arxiv-lit-review-word-gsheet`              | 已完成基础验证 | 上游 5 条论文记录 + 趋势 API 已静态化；排除 quant-ph 噪声并拒绝旧 Golden 的无来源引用数，Excel、Word、离线 Tracker replay Judge `1.0` |
| `shared_source_multi_output` | OmegaUse `officeval_088`                                       | 已完成基础验证 | 5 份真实产品说明 -> 5 行 Excel + 高价前三 Word；8 张来源裁图、双输出一致性 Judge `1.0` |
| `update_propagation`         | Cowork-Bench `insales-shipping-tax-excel-word`                  | 已完成基础验证 | 将上游运费含税变更传播到 Excel 模型与 Word 说明；静态化原始店铺数据，两交付物一致性与变异 replay 通过 |
| `cross_file_consistency`     | Cowork-Bench `insales-refund-root-cause-excel-word-email`      | 已完成基础验证 | 上游 13 退款、13 订单、82 商品与 396 评论已静态化；修正旧 Golden 金额冲突，Excel、Word、邮件 replay Judge `1.0` |
| `template_driven_package`    | Cowork-Bench `terminal-pw-kulinar-excel-word-gcal`             | 已完成基础验证 | 上游 22 行价格 + 50 条 Kulinar 配方 + PDF/预算说明；脚本、JSON、Excel、Word、ICS Golden/replay Judge `1.0` |
| `review_and_revision`        | OmegaUse `officeval_080`                                       | 已完成基础验证 | 保留 26 MB PPTX、5 MB 修改要求 DOCX、12 条意见与 4 张内嵌图；18 页修订稿 + 响应表，Golden/fresh/变异 replay 均 `1.0` |
| `workspace_orchestration`    | Cowork-Bench `terminal-canvas-excel-word-teamly-email`         | 已完成基础验证 | 上游 Canvas 2 课程、4,648 enrollment、26 作业、32,424 submission 已静态化；脚本、JSON、四表 Excel、Word、Tracker 与邮件 replay Judge `1.0` |

## 本轮新增的 14 道异质题

| 领域       | OfficeVal 题号 | 归入正式家族                | 状态             | 逐题改造和验证摘要                                                     |
|------------|---------------:|-----------------------------|------------------|------------------------------------------------------------------------|
| 文字文档   |            003 | `formatting_fidelity`       | 已完成基础验证   | 95 行会议记录表清理与横线式表格；Golden、fresh replay 和负例通过       |
| 文字文档   |            010 | `formatting_fidelity`       | 已完成基础验证   | 压平流程图重建为 30 个图形和 39 条连接线；零图片替代并完成视觉检查     |
| 文字文档   |            029 | `document_extraction`       | 已完成基础验证   | 图片表格重建为可编辑 Word 表格；上游 Verifier 矛盾已留档               |
| 文字文档   |            033 | `editing_rewriting`         | 已完成基础验证   | 1,435 个零售价逐项乘以 1.1；独立重算并保护非目标内容                   |
| 文字文档   |            040 | `formatting_fidelity`       | 已完成基础验证   | 样式、段落、页面和水印逐项设置；上游标题误报已留档                     |
| 演示文稿   |            043 | `material_to_deck`          | 已完成基础验证   | 两张中英名片图重建为两页可编辑 PPT；仅二维码保留图片                   |
| 演示文稿   |            046 | `material_to_deck`          | 已完成基础验证   | 两张研究框架图重建为 132/108 个原生对象；逐页视觉检查                  |
| 演示文稿   |            051 | `material_to_deck`          | 已完成基础验证   | 异常检测模型图重建为 166 个原生对象；拒绝整页截图替代                  |
| 演示文稿   |            054 | `material_to_deck`          | 已完成基础验证   | 工业水回用图重建为 114 个原生对象；依据输入纠正上游第 8 步错误         |
| 演示文稿   |            056 | `material_to_deck`          | 已完成基础验证   | 客诉闭环图重建为 107 个原生对象；字体、节点和连接关系受检              |
| 演示文稿   |            060 | `visual_editing`            | 已完成基础验证   | 只修第 4 页图片布局；其余四页逐像素一致，源媒体不变                    |
| 演示文稿   |            064 | `template_based_generation` | 已完成基础验证   | 12 张照片顺序、裁切、四种画框和模板对象保持均受检                      |
| 跨办公应用 |            059 | `workspace_orchestration`   | 已完成基础验证   | 在 Word 中加入相对 PPT 链接；PPT 逐字节保持，上游实体名错误已留档      |
| 跨办公应用 |            078 | `workspace_orchestration`   | 已完成基础验证   | 在 PPT 第 8 页加入可编辑 Excel 链接按钮；工作簿逐字节保持              |

## 后续处理顺序

42 个家族的首题和本轮 14 道定向异质题已齐全，下一阶段不是继续为凑数量批量扩题，而是通过真实 Agent Rollout 决定哪些家族还需要第 2 道或更多异质题：

1. 用至少 3 个非参考 Agent Harness 跑 pass@1 / pass@8，检查题目是否过易、不可解或依赖参考解法。
2. 对失败样本做 Judge 梯度校准，补齐真实的部分完成、格式损坏、事实错误与跨文件冲突样本。
3. `cowork-evo-240-v0.1` 已按来源、输入哈希和显式泄漏组冻结 `feedback / selection / final` 三段式切分；后续语义近重复审计若发现新增泄漏，只能发布新版本，不能原地修改 v0.1。
4. 只对 Rollout 显示覆盖不足或方差过大的家族定向增加第 2 道题；新来源仍必须先完成许可、隐私、输入和真值审计。
