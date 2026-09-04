# Evidence ledger

- E1 Query: `adaptation/source-task.json` field `task`
- E2 Evaluation: `adaptation/source-task.json` field `evaluation`
- E3 Input bytes: `metadata/input-manifest.json`; recomputed package hashes: a5c8026dc09f38411c6b8783dad74f7f66f08387eeb21b9d8eb7f28a90dd5b6e
- E4 `score.xlsx`, Sheet1 `A1:C21`: 20 non-empty student rows; independently recomputed averages are `80.05` for `midterm1` and `78.30` for `midterm2`, with signed difference `-1.75`.
- E5 Forbidden source terms: none
- E6 Deliverable contract: midterm_report.docx
- E7 Golden: `solution/` is audit-only; the Judge recomputes count, means, and difference from the workbook and does not use Golden bytes as truth.
