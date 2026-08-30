# 固定负例与中等例

这些文件只用于校准 `pptarena-003-source-footer-edit` 的 Judge，不是新题或 Golden。

- `unmodified`：未修改的冻结输入。
- `wrong-text`：在第 5 页写入错误 URL。
- `wrong-slide`：将正确引用写到第 4 页。
- `changed-original-text`：修改第 5 页的原有关键数值。
- `flattened`：将第 5 页替换为上游 Golden 的整页截图。
- `duplicate`：重复写入两次正确引用。
- `hardcoded-marker`：在文件中夹带 `solution/Golden` 泄漏标记。
- `off-slide`：内容正确，但文本框位于页面下方，用作中等样例。
- `corrupt` 与 `empty`：损坏文件和空交付。

生成逻辑只修改这一道题的固定输入，见 `tests/build_fixtures.py`。
