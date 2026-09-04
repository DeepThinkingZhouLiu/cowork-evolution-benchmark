# 恢复 TEAM 收入模型并修复 DCF 级联公式

打开 `08_03_input.xlsx`，审计并修复 TEAM 财务模型中的删除行、断裂引用和级联估值错误。先恢复 `Revenue Build` 的年度收入汇总及增长行，再恢复 `Operating Model + DCF` 的 Unlevered Free Cash Flow 行并修复依赖公式。

`Comps + WACC` 中的 TEAM 数据应通过 INDEX/MATCH 与 `TEAM US` 行标签匹配；`Revenue Build` 的历史 FY 总收入应按年度表头从 `IS Annual` 匹配，不能按固定行序硬编码。保留 13 张工作表的顺序和其他内容。

输出必须是可编辑、可重算的原生工作簿，指定模型区域不得残留 `#REF!`。将唯一交付文件保存为 `/app/output/team_model_08_03_repaired.xlsx`。
