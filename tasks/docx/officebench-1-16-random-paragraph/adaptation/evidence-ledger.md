# Evidence ledger

- E1 上游 Query/Evaluator：`adaptation/source-task.json`，OfficeBench `1-16/0` @ `b978b808667c32b52ce19a67ce1def1de9ae02b7`。
- E2 许可：`adaptation/LICENSE-OfficeBench`，Apache-2.0，SHA256 `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`。
- E3 唯一事实权威：`random_paragraph.txt`，SHA256 `4439b9fb95c87edce812407d9961abb9d2928fb1a21668e3c336fb44e9285f55`。
- E4 受保护 Word 模板：`cafe_profile_editorial_template.docx`，SHA256 `f53dfaf5aa724f5ef644a1b9e4ef16f8b7d840c484a0785b3770e3924faa916b`。
- E5 批准编辑票据：`cafe_profile_approved_copy_ticket.docx`，SHA256 `7f728aacce45dd7884fc3297979ac7c9dab8cc5bf420e1b445bfffad7e6066ed`。
- E6 输入字节一致：`data/input_files` 与 `environment/input_files` 的三份同名文件逐字节一致，Docker 固定复制到 `/app/data/input_files` 并链接为 `/app/input_files`。
- E7 句级真值：S1–S5 的源/批准词数为 `18/14, 21/15, 21/13, 18/10, 25/16`，汇总 `103 -> 68`, delta `-35`。
- E8 批注真值：S2 `locals and tourists alike`，S3 `Freshly brewed coffee and baked pastries`，S5 `time seemed to slow`；正文逐行来自 E5。
- E9 导航契约：批准修订书签 `review_S1`–`review_S5`，台账内链一对一指向。
- E10 评分独立性：Judge 只读 E3–E5 与候选 DOCX，不读 Golden 哈希、solve 标记或候选自报结果。
