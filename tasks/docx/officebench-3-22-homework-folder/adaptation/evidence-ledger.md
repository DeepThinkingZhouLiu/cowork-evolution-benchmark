# Evidence ledger

- E1 Query: `adaptation/source-task.json` field `task`
- E2 Evaluation: `adaptation/source-task.json` field `evaluation`
- E3 Input bytes: `metadata/input-manifest.json`; recomputed package hashes: 1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf
- E4 Required source terms: none
- E5 Forbidden source terms: none
- E6 Deliverable contract: homework.docx, homework/homework.docx
- E7 Golden: `solution/` is audit-only; the Judge recomputes from E1-E3 and does not use Golden bytes as truth.
- E8 Same-type multi-output check: both `homework.docx` and `homework/homework.docx` are required and must contain the same OCR result.
