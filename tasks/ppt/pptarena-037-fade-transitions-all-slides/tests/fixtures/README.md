本目录不手工存放大量 PPTX；由 `../build_fixtures.py` 基于冻结输入和人工 Golden 现场生成。

- `valid-default`：从冻结输入出发，在 9 个幻灯片 XML 中各写入不指定速度的标准 Fade，不复制 Golden 的 `slow` 选择。
- `valid-through-black-fast`：从冻结输入出发，在全部 9 页使用 `fast` + Through Black Fade，验证 Judge 不锁死时长和变体。
- `repacked-content-types`：只对 `[Content_Types].xml` 增加无害空白，语义、转场和渲染都不变，验证字节级保持分项产生严格介于 0 和 1 的中间分。
- `empty` / `corrupt` / `unmodified`：验证缺文件、损坏文件和未添加转场不会过关。
- `missing-one` / `one-cut` / `all-cut` / `duplicate`：验证 9/9 页、效果类型和转场节点唯一性。
- `auto-advance` / `disable-click` / `sound`：验证不允许夹带自动换页、禁用点击换页或转场声音。
- `changed-text` / `changed-notes` / `added-animation` / `changed-image`：验证文字、讲者备注、动画时序和媒体保护。
- `flattened-slide`：把第 1 页替换为整页图片，验证原生可编辑性和防压平。
- `hardcoded-marker` / `extra-slide`：验证答案路径泄漏和页数保护。
