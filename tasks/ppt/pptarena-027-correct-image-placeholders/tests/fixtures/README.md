# Case 27 测试样例

本目录的实体样例由单题脚本 `../build_fixtures.py` 从冻结输入和本地 Golden 重建，不纳入正式任务包。

- `valid-224-centered` 和 `valid-232-centered` 是两个几何不同于 Golden 的合法答案。
- `repacked-content-types` 是功能完整但包骨架被无害重写的中等样例。
- 其余样例分别破坏图片语义对应、等大对齐、标签中心、图片裁剪、文字、主题、版式、媒体、原生可编辑性或交付卫生。
