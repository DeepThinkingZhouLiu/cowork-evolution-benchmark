# Evidence ledger v0.3

- `E1`：上游 Query 与 evaluator，`adaptation/source-task.json`，542 字节，SHA256 `9e5504efbf2f9e3e245f753fd92481cb94f23d9e97aeff6c8988fcdd413c5ae7`。
- `E2`：上游 Apache-2.0 许可，`adaptation/LICENSE-OfficeBench`，11,357 字节，SHA256 `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`。
- `E3`：受保护的大纲主文档，`data/input_files/sample_syllabus.docx`，188,874 字节，SHA256 `9efefdcac29942e59edf9537aa5cfade576ea55b408dde34205792fd773c8505`。
- `E4`：研究诚信增补来源，`data/input_files/academic_integrity.docx`，16,079 字节，SHA256 `260bd988b35c3a3577238d9f15f6a15010a39e5de3a167cc73b65bbaad3291cc`。
- `E5`：受保护内容，E3 的 187 个现有段落、34×4 课程表、标识图片、外部链接、样式/run 格式、首节页面几何和页眉页脚关系。
- `E6`：一般研究数据不当行为，E3 P96 的 `fabrication, falsification or sabotage of research data`。
- `E7`：Turnitin 验证与承诺语境，E3 P101 的 `verified using Turnitin technology`，是新增政策的前锚点。
- `E8`：后继章节边界，E3 P104 的 `Special Needs`，是新增政策必须位于其前的后锚点。
- `E9`：校外/社区适用范围，E3 P184–P185，明确政策和行为准则延伸至学校工作人员、社区组织及客户。
- `E10`：研究数据范围，E4 P1，涵盖从学生、教师、同事取得的口头、书面、观察、记录和作品数据。
- `E11`：研究设计四项控制，E4 P4–P7：目的与问题/数据关联、无偏见选择、风险收益、全时匿名。
- `E12`：输出合同，`sample_syllabus_policy_review.docx`，只交付一个可编辑 DOCX。
- `E13`：Golden 仅用于作者侧可用性验收；Judge 从 E3/E4 独立重算内容和结构，不读取 Golden 字节、哈希或 solver 私密标记。
