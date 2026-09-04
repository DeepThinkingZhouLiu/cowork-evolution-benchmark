# Evidence ledger

- E1 Original Query: `adaptation/source-task.json` field `task`.
- E2 Original evaluator: `adaptation/source-task.json` field `evaluation`.
- E3 Input bytes: `metadata/input-manifest.json`; SHA256 `19f2b63d8ce3010d55b28058226b0a4db2ee38e5567ab1f9ef5ae4dedad7092a`.
- E4 Adaptation: the original answer container was `data/answer.txt`; this package requests only an editable `final_pass_count.docx` DOCX and does not add a second deliverable.
- E5 Truth: the Judge re-reads the workbook and recomputes the answer; it does not compare Golden bytes.
