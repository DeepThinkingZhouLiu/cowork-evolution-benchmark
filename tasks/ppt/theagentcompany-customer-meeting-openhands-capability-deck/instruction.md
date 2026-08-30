请基于下列冻结输入，为 OpenHands 潜在客户会议制作能力范围演示文稿：

- `/app/input_files/Original_Request.md`
- `/app/input_files/OpenHands_Evaluation_Snapshot.json`
- `/app/input_files/RocketChat_Recipient_Context.json`

交付物必须为：

- `/app/output/openhands_intro.pptx`
- `/app/output/Delivery_Record.json`

`openhands_intro.pptx` 必须是恰好 6 页、16:9 的业务演示文稿：

- 第 1 页为客户会议封面，明确这是一份基于冻结 GitLab evaluation tree 的 OpenHands 能力清单，总计 3 个领域、16 个 benchmark entry。
- 第 2 页总览 3 个领域及数量：Software Engineering 8、Web Browsing 2、Misc. Assistance 6。
- 第 3 页仅介绍 Software Engineering 领域，完整列出冻结快照中的 8 个 benchmark。
- 第 4 页仅介绍 Web Browsing 领域，完整列出冻结快照中的 2 个 benchmark。
- 第 5 页仅介绍 Misc. Assistance 领域，完整列出冻结快照中的 6 个 benchmark。
- 第 6 页说明评测来源、证据边界和离线交付。

上游证据只能支持“benchmark 名称 + 领域归属”。演示文稿不得编造分数、通过率、成功率、模型版本、排名、比较优势或客户效果；必须明确说明“evaluation inventory 不等于已测性能”。

原任务要求通过 RocketChat 把文件名发给 Huang Jie。本 Harbor 任务为离线可重放环境，不得伪造已发送成功；`Delivery_Record.json` 必须准确记录：

- 文件名 `openhands_intro.pptx`，原路径 `/workspace/openhands_intro.pptx`。
- 收件人 `Huang Jie`，渠道 `RocketChat direct message`，消息内容仅为 `openhands_intro.pptx`。
- 用途 `customer meeting materials`。
- `network_side_effect` 必须为 `not_executed_offline`。

所有页面必须保留来源，所有文本和图形必须原生可编辑，不得整页压平。成品要能在 LibreOffice Impress 中打开和渲染，无重叠、裁切或超出页面。
