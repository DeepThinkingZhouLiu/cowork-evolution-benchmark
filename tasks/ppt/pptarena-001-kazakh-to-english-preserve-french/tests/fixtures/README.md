# 固定校准样例

这些文件只用于校准 PPTArena Case 1 的 Judge，不是新题或 Golden。

- valid-repackage：只改变 ZIP 压缩方式，Office 语义不变，证明 Judge 不依赖 Golden 文件哈希。
- unmodified：完全没有翻译的冻结输入。
- upstream-uncorrected：上游人工 Ground Truth，仍含 `meal, a diet`、`classification`、`Cafe` 和 `To drink` 四处语义问题。
- missed-translation：将 Coffee 恢复为未翻译的 `кофе`。
- translated-french：误将表格左列法语 `Le pain` 改成英语。
- wrong-conjugation：将语法上的“动词变位”错译为 classification。
- wrong-water：漏掉 `Su ishu` 中的 water 语义。
- wrong-style：将 Feedback 标题误改为红色。
- moved-target：将第 17 页主文本框右移 1 英寸。
- deleted-shape：删除第 17 页主文本框。
- flattened-slide9：将第 9 页可编辑表格替换为整页截图。
- hardcoded-marker：夹带 `solution/Golden` 答案路径。
- extra-output：主交付正确，但多交付 `debug.txt`。
- corrupt 和 empty：损坏文件与空交付。

生成逻辑只针对这一道冻结输入与已审计 Golden，见 `tests/build_fixtures.py`。
