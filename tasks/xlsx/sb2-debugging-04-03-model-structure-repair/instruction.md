# 恢复财务模型结构并修复级联引用

打开 `04_03_input.xlsx` 并彻底修复模型。文件存在删除行引发的 `#REF!`、损坏的定义名称、缺失财务科目和级联计算错误。请按依赖顺序先恢复结构，再修复依赖公式，同时保留不相关内容和 13 张工作表的顺序。

可核验的历史数据：FY2019 Restructuring Charges 为 -5.7M；Merger & Related Restructuring Charges 在 FY2018 为 -0.6M、FY2019 为 -1.36M。模型单位行中的 `Mil` 为 1,000,000。

输出必须是无损、可编辑、可重算的原生工作簿，指定模型区域不得残留 `#REF!`。将唯一交付文件保存为 `/app/output/model_04_03_repaired.xlsx`。
