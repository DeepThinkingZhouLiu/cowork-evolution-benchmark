本目录不手工存放大量 PPTX；由 `../build_fixtures.py` 基于冻结输入、人工 Golden 和已目视审核的第 3 页截图现场生成。

- `valid-alternate`：从冻结输入独立新建一个 8.4×4.75 英寸的原生饼图，使用右侧图例和百分比标签，不复制 Golden 的坐标、图例选择或数据标签形式。
- `empty` / `corrupt` / `unmodified`：验证缺文件、损坏文件和未修改的柱形图不会过关。
- `wrong-chart-type` / `wrong-value` / `wrong-category` / `second-chart`：验证图表类型、数据和唯一性。
- `changed-slide3-title` / `changed-other-chart` / `extra-slide`：验证目标页非图表内容和其他页面不得变更。
- `flattened-slide3`：用整页截图伪装第 3 页，验证原生可编辑性和防压平。
- `hardcoded-marker`：验证答案路径泄漏不会被接受。
- `no-labels-or-legend` / `tiny-labels`：验证无法辨识 8 个类别或标签过小会触发核心门槛。
- `undersized-chart`：只破坏布局尺寸，验证 Judge 会给出严格介于 0 与 1 的中间分。

`golden-slide3.jpg` 仅用于构造“压平目标页”反例，它不参与 Judge 正向评分。
