# Evidence ledger

- E1 Original Query: `adaptation/source-task.json` field `task`.
- E2 Original evaluator: `adaptation/source-task.json` field `evaluation`.
- E3 Input bytes: `metadata/input-manifest.json`; SHA256 `5a495d6f9d3c319207dc1b9ed002672a42eee30c3f25741b31ca4e7df399faeb`.
- E4 Adaptation: the original answer container was `data/answer.txt`; this package requests only an editable `prerequisite_check.docx` DOCX and does not add a second deliverable.
- E5 Truth: the Judge re-reads the workbook and recomputes the answer; it does not compare Golden bytes.
