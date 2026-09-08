# Evidence ledger

| ID | Evidence | SHA256 / finding | Adaptation use |
|----|----------|------------------|----------------|
| E1 | `adaptation/source-task.json` / upstream `tasks/2-20/subtasks/1.json` | `9dcb7ebe091e7e470ec408dd706f0d9fb97c9526f4feaaa9a3b9dba01ad5c755`; byte-identical | Query, target filename history, original keyword evaluator arguments |
| E2 | `adaptation/upstream-evaluator.py` from upstream `utils/evaluate.py` | Upstream full file `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90` | Exact case-insensitive containment semantics; audit only |
| E3 | Three input workbook copies | `c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7`; byte-identical | Sole factual source and runtime offline input |
| E4 | `adaptation/LICENSE-OfficeBench` / upstream `LICENSE` | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`; byte-identical | Apache-2.0 permission and attribution |
| E5 | Old v0.2 Golden | `44a1029192d7f6798993c351e906713321483ee99b68c01a00a5250e9a996e30` | Inspected only; three plain paragraphs, superseded rather than used as truth |
| E6 | Workbook cells `A1:J10` | 9 student rows; three exact CS matches at rows 2, 3, 7 | Queue selection, individual evidence, source locators, aggregates |
| E7 | Task-specific Golden | generated only from E3 by `solution/solve.py` | Audit/reference output; never read by Judge |
| E8 | Task-specific deterministic Judge | recomputes from E3 | Selection, privacy, order, calculations, Word objects, structure and rendering |

No host path, network fetch, temporary URL, provider secret, Golden hash, self-reported candidate value, or solver-only marker participates in scoring.
