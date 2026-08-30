# 固定负例与中等例

这些文件只用于校准 PPTArena Case 15 的 Judge，不是新题或 Golden。

- unmodified：未修改的冻结输入。
- skip-cover-title：漏改第 1 页主标题 CSE1300。
- skip-one-title：漏改第 9 页主标题。
- incomplete-multirun-title：第 6 页多运行标题只有部分文字为粗体。
- changed-cover-subtitle：错误地取消第 1 页课程副标题的粗体。
- bold-everything：把正文和课程副标题一并设为粗体。
- wrong-title-style：将第 8 页标题改为红色。
- changed-body：改写第 10 页正文。
- flattened：将第 4 页替换为上游 Golden 的整页截图。
- hardcoded-marker：在文件中夹带 solution/Golden 泄漏标记。
- moved-title：标题粗体正确，但第 10 页标题下移 0.08 英寸，用作中等样例。
- corrupt 与 empty：损坏文件和空交付。

生成逻辑只修改这一道题的固定输入，见 tests/build_fixtures.py。
