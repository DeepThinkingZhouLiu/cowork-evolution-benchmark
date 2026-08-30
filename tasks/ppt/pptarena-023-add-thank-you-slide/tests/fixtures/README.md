本目录不手工存放大量 PPTX；由 `../build_fixtures.py` 基于冻结输入、人工 Golden 和已目视审核的末页截图现场生成。

- `valid-alternate`：从原始 14 页出发，用空白版式和 44pt 可编辑文本框新建等价末页，不复制 Golden 的主题布局。
- `empty` / `corrupt` / `unmodified`：验证缺文件、损坏文件和未完成任务不会过关。
- `wrong-title` / `duplicate-title`：验证标题内容和唯一性。
- `inserted-at-start` / `extra-slide`：验证新页必须恰好位于第 15 页。
- `changed-original-body` / `changed-original-chart`：验证原 14 页的文字和图表数据不得改动。
- `flattened-final`：用整页截图伪装末页，验证可编辑性。
- `hardcoded-marker` / `invisible-title`：验证答案路径泄漏和不可见文字不会被接受。
- `off-center-title`：只移动末页标题，验证分项评分会产生严格介于 0 与 1 的中间分。

`golden-slide15.jpg` 仅用于构造“压平末页”反例，它不参与 Judge 正向评分。
