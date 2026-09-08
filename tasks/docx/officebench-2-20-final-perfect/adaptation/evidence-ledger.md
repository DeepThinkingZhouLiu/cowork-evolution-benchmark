# Evidence ledger

| ID | Evidence | SHA256 / finding | Adaptation use |
|----|----------|------------------|----------------|
| E1 | `adaptation/source-task.json` / upstream `tasks/2-20/subtasks/2.json` | `aaad40b592a9cf8f9a940276c03b0e3c7de8d9e55642535e51a2df5d2c9e68c9`; byte-identical | Query, target filename history, original keyword-evaluator arguments |
| E2 | `adaptation/upstream-evaluator.py` from upstream `utils/evaluate.py` | Upstream full file `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90` | Exact case-insensitive containment semantics; audit only |
| E3 | Upstream/package/environment input workbook copies | `c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7`; byte-identical | Sole factual source and runtime offline input |
| E4 | `adaptation/LICENSE-OfficeBench` / upstream `LICENSE` | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`; byte-identical | Apache-2.0 permission and attribution |
| E5 | Old v0.2 Golden | `17d832895038bc3b8f09828ab0c7af792aee51f3f8bfaac2af10aba907f3b635` | Inspected only; two plain paragraphs, superseded rather than used as truth |
| E6 | Workbook cells `A1:J10` | 9 rows; exact Final=100 at rows 6 and 9; detailed truth in adaptation README | Selection, locators, calculations, aggregates, boundary, closure |
| E7 | Task-specific Golden | generated only from E3 by `solution/solve.py` | Audit/reference output; never read by Judge |
| E8 | Task-specific deterministic Judge | recomputes from E3 | Exact threshold, evidence, calculations, Word objects, structure, safety, rendering |

No host path, network fetch, temporary URL, provider secret, Golden hash, self-reported candidate value, or solver-only marker participates in scoring.
