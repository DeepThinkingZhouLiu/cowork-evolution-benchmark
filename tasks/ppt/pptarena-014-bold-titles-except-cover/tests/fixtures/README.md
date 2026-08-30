# 固定负例与中等例

这些文件只用于校准 PPTArena Case 14 的 Judge，不是新题或 Golden。

- unmodified：未修改的冻结输入。
- skip-one-title：第 17 页标题未设为粗体。
- skip-hidden-titles：只漏掉隐藏的第 8、9、23、31 页标题。
- bold-cover：错误地把第 1 页封面主标题也设为粗体。
- bold-everything：把正文和封面一并设为粗体。
- wrong-title-style：将第 12 页标题改为红色。
- changed-body：改写第 20 页正文。
- changed-hidden-status：把第 8 页从隐藏改为显示。
- flattened：将第 10 页替换为上游 Golden 的整页截图。
- hardcoded-marker：在文件中夹带 solution/Golden 泄漏标记。
- moved-title：标题粗体正确，但第 17 页标题下移 0.08 英寸，用作中等样例。
- corrupt 与 empty：损坏文件和空交付。

生成逻辑只修改这一道题的固定输入，见 tests/build_fixtures.py。
