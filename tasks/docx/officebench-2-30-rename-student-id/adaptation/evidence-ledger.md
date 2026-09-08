# Evidence ledger

| ID | Evidence | SHA256 / finding | Adaptation use |
|----|----------|------------------|----------------|
| E1 | `adaptation/source-task.json` / upstream `tasks/2-30/subtasks/0.json` | `b11ef3866e170f8a2cf3267e33dfbd4c5cbda89742f905284eac9f902c6ffcb4`; byte-identical | Original identity-derived rename Query and filename-only evaluator arguments |
| E2 | Upstream `utils/evaluate.py` | `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90` | Exact `evaluate_file_exist` semantics; audit only |
| E3 | Upstream/package/environment DOCX copies | `7df2fd4bf45616e4e78f6adf4dfb5c10ac18886bc777bd43064740be12b8b03b`; byte-identical | Sole factual source and runtime-offline input |
| E4 | `adaptation/LICENSE-OfficeBench` / upstream `LICENSE` | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`; byte-identical | Apache-2.0 permission and attribution |
| E5 | Old v0.2 Golden and solution copies | `7df2fd4bf45616e4e78f6adf4dfb5c10ac18886bc777bd43064740be12b8b03b`; byte-identical to source | Copy-and-rename baseline inspected and superseded, not used as truth |
| E6 | Source DOCX body | 13 paragraphs, 14 runs, 1,817 normalized characters, body SHA256 `9215262c621e091ae431a8a2b35e56b5d75a92f8183693ce5f5c0a13ccaabf05` | Protected block, per-paragraph ledger, closure |
| E7 | Source paragraph/run/numbering structure | paragraph 3 run 2=`A12345678`; paragraph 7 level 0; paragraphs 8–10 level 1; source `numId 1` | Identity locator, exact run boundaries, native numbering preservation |
| E8 | Task-specific Golden | generated only from E3 by `solution/solve.py` | Audit/reference output; never read by Judge |
| E9 | Task-specific deterministic Judge | independently recomputes from E3 | Identity, original content/format, ledger, controls, metadata, structure, safety, and rendering |

No host path, network fetch, temporary URL, provider secret, Golden hash, self-reported candidate value, student answer, or solver-only marker participates in scoring.
