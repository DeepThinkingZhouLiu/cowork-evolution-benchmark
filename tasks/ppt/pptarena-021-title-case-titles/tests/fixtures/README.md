# 固定有效变体、负例与中等例

这些文件只用于校准 PPTArena Case 21 的 Judge，不是新题或 Golden。

- valid-strict-questions：第 2、4 页采用 Who Are We? / What Do We Do?；这是上游规则明确允许的第二种正确答案。
- unmodified：未修改的冻结输入。
- missed-one-title：第 11 页仍为 results of workshops。
- lowercase-dsi：第 12 页错误写成 Growth of dsi。
- wrong-question-case：第 2 页错误写成 Who Are we?，不属于两种允许写法。
- wrong-minor-word：第 14 页错误写成 Results Of Growth。
- wrong-title-style：第 8 页标题被改成红色。
- changed-body：第 2 页正文被改写。
- deleted-title：删除第 10 页语义标题文本框。
- flattened：将第 10 页替换为上游 Golden 的整页截图。
- hardcoded-marker：在文件中夹带 solution/Golden 泄漏标记。
- extra-slide：错误增加第 16 页空白页。
- moved-title：大小写正确，但第 10 页标题下移 0.08 英寸，用作中等样例。
- corrupt 与 empty：损坏文件和空交付。

生成逻辑只修改这一道题的固定输入，见 tests/build_fixtures.py。
