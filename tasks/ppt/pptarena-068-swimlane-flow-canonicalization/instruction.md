请编辑 `/app/input_files/SwimlaneFlowCanonicalization_TestA.pptx`，将第 1 页上杂乱的流程图改造成一张规范、可编辑的三泳道图。

具体要求：

- 将页面标题改为 `Canonical Process • 3 Swimlanes`。
- 建立三条水平泳道，从上到下分别标注 `Intake`、`Processing`、`Output`。泳道要覆盖页面主要宽度，边界和标签清晰。
- 保留原流程中的 5 个步骤框、`OK?` 判定菱形和图片，并按以下归属重新排列：
  - `Intake` 泳道：`Step 1: Intake`、`Step 2: Validate`、`Step 3: Transform`。
  - `Processing` 泳道：`OK?` 判定菱形和原图片。
  - `Output` 泳道：`Step 4: Reject`、`Step 5: Export`。
- 5 个矩形步骤框必须等宽等高，并保留原有文字、填充色和核心视觉风格。
- 在每个步骤框附近添加圆形编号徽章，编号必须与步骤号一一对应：步骤 1–5 分别对应徽章 `1`–`5`。
- 删除原有直线连接线，用 5 条原生可编辑的折线连接符表示 `Step 1 -> Step 2 -> Step 3 -> OK?`，以及 `OK? -> Step 4`和 `OK? -> Step 5`。每条连接符末端都要有箭头，且不得穿过其他步骤框。
- 在页面下方加入 `Legend`，说明泳道、统一步骤框、折线箭头和 1–5 徽章的含义。
- 整页应无重叠、无裁切，保持原生可编辑，不得用整页截图或覆盖图层代替图形和连接符。

输出文件必须命名为 `/app/output/SwimlaneFlowCanonicalization_GroundTruthA.pptx`。请仅在 `/app/output` 中写入最终交付文件。
