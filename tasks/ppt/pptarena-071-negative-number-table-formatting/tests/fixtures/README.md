# 固定校准样例

这些文件只用于校准 PPTArena Case 71 的 Judge，不是新题或 Golden。

- valid-digits-only：只将 6 处负数的数字设为红色粗体，验证 Judge 不依赖 Golden 的 run 分段。
- unmodified：未编辑的冻结输入。
- skip-one：漏掉 `Range: -3 to 7 items` 中的负数。
- red-not-bold：`-5.6%` 为红色但不是粗体。
- bold-black：`-$15,250` 为粗体但改成黑色。
- format-positive：误将正数 `3.2%` 设为红色粗体。
- format-all：将所有表格文字都设为红色粗体。
- changed-text：擅自改写单元格文字。
- moved-table：将第 3 页表格水平移动 0.1 英寸。
- flattened：将第 2 页原生表格替换为整页截图。
- hardcoded-marker：在页面中夹带 `solution/Golden` 答案路径。
- extra-output：主交付正确，但多交付 `debug.txt`。
- corrupt 和 empty：损坏文件与空交付。

生成逻辑只针对这一道冻结输入，见 `tests/build_fixtures.py`。
