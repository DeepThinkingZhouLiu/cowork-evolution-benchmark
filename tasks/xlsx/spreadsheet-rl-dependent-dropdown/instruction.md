# 构建账户类别级联下拉表

打开 `account_form_input.xlsx`。根据 `Account` 表中的 Category/Subcategory 数据，补全工作簿中的级联选择器：

1. 新建 `Lists` 工作表，将去重后的 Category 横向放在第 1 行，并在每个 Category 下方按首次出现顺序列出其 Subcategory。
2. 创建 `Categories` 命名区域，以及每个类别对应的、将空格替换为下划线的 Subcategory 命名区域。
3. 在 `Form!B2` 设置类别列表验证，在 `Form!C2` 设置依赖 `B2` 的子类别列表验证；`Account` 与现有 `Form` 内容必须保持不变。

交付一个可编辑的原生 XLSX 文件 `/app/output/account_form_completed.xlsx`。