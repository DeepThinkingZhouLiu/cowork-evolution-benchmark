# Evidence ledger

| ID | Evidence | SHA256 / finding | Adaptation use |
|----|----------|------------------|----------------|
| E1 | `adaptation/source-task.json` / upstream `tasks/2-21/subtasks/0.json` | `44c5906595c29ecf33fcecb81728e814710e1eab70bf9c3fb3e6a0b0216eb07c`; byte-identical | Original Query, target filename history, 15-keyword evaluator arguments |
| E2 | Upstream `utils/evaluate.py` | `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90` | Exact case-insensitive containment semantics; audit only |
| E3 | Upstream/package/environment workbook copies | `d7764b78e2931fbcda87f2a45ef140693c3d6977e0235b264aa868cf949bab8a`; byte-identical | Sole factual source and runtime-offline input |
| E4 | `adaptation/LICENSE-OfficeBench` / upstream `LICENSE` | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`; byte-identical | Apache-2.0 permission and attribution |
| E5 | Old v0.2 Golden and solution copies | `2395845887826d5f3c2e02460e0545504145c4efbba95d6d8645fbe593a7a15e`; byte-identical | One-page course-name list inspected and superseded, not used as truth |
| E6 | Workbook `Sheet1!A1:D16` | 15 complete rows; no formulas/merges/hidden data; row facts and calculations recorded in README | Ledger, term/subject/grade/continuity truth and source locators |
| E7 | Public v0.3 grade-point rule | A/A-/B+/B = 4.0/3.7/3.3/3.0, every course counted once | Unweighted point/GPA calculations; explicitly not a source-workbook claim |
| E8 | Task-specific Golden | generated only from E3 and E7 by `solution/solve.py` | Audit/reference output; never read by Judge |
| E9 | Task-specific deterministic Judge | independently recomputes from E3 and E7 | Per-row facts, calculations, Word controls, structure, safety, and rendering |

No host path, network fetch, temporary URL, provider secret, Golden hash, self-reported candidate value, or solver-only marker participates in scoring.
