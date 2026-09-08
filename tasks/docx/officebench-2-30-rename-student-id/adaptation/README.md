# officebench-2-30-rename-student-id v0.3 adaptation record

Status: `harbor_replayed`; difficulty remains `unassigned`; `calibration_status=uncalibrated`.

## Upstream seed, evaluator, and license

- Dataset: OfficeBench task `2-30/0`, Git commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Frozen Query/evaluator arguments: `adaptation/source-task.json`, SHA256 `b11ef3866e170f8a2cf3267e33dfbd4c5cbda89742f905284eac9f902c6ffcb4`.
- Frozen input: upstream `tasks/2-30/testbed/data/homework_withname.docx`, SHA256 `7df2fd4bf45616e4e78f6adf4dfb5c10ac18886bc777bd43064740be12b8b03b`. The upstream, `data/input_files`, and `environment/input_files` copies are byte-identical.
- Evaluator implementation consulted: upstream `utils/evaluate.py`, SHA256 `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90`. Its `evaluate_file_exist` path checked only whether `./data/A12345678.docx` existed; it did not inspect identity extraction, document content, formatting, editability, safety, or renaming exclusivity.
- License: Apache-2.0. `adaptation/LICENSE-OfficeBench` is byte-identical to upstream `LICENSE`, SHA256 `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`.

## Full context read before authoring

Every package file was read directly: instruction, TOML, rubric, Dockerfile, both input copies, manifests, source task/license, adaptation files, metadata, solver/wrapper, Judge/test wrapper/expectations, and both old DOCX outputs. The source, old solution, and old Golden were byte-identical. The DOCX was parsed down to paragraphs, runs, paragraph properties, numbering IDs/levels, OOXML parts, relationships, metadata, tables, images, hyperlinks, bookmarks, fields, comments, equations, and section geometry; it was also rendered and visually inspected.

The exact upstream `2-30/0` JSON, source DOCX, README, license, revision, and complete relevant `evaluate_file_exist` implementation were inspected. The source is a single portrait Letter page with 13 paragraphs, 14 runs, four numbered paragraphs, one section, no tables/images/hyperlinks/bookmarks/fields/comments/equations, and no hidden auxiliary evidence. The old answer merely copied the input under the requested filename, so any empty or unrelated DOCX with that filename would have passed the upstream evaluator.

## Frozen input truth

The source contains one identifier matching `^[A-Z]\d{8}$`: `A12345678`, in body paragraph 3, run 2. Paragraph 1 is `CSE 221: Homework 1`, yielding course `CSE 221` and assignment `Homework 1`; paragraph 2 is `Winter 2022`. The normalized 13-paragraph body contains 1,817 characters and has SHA256 `9215262c621e091ae431a8a2b35e56b5d75a92f8183693ce5f5c0a13ccaabf05` when the normalized paragraphs are joined by LF.

The per-paragraph normalized SHA-12 values are `0995458ecbde`, `635915f5dc59`, `426bba8e61df`, `ef135f540e64`, `7da6256e2e68`, `ce9721b09ec1`, `706969d595e0`, `0815efe9d8ce`, `719ea478708f`, `383e2d946278`, `33f71ce9b039`, `c4104a868260`, and `9cb0a612dddd`. Paragraph 7 is the level-0 question and paragraphs 8–10 are its level-1 Hydra/Multics/Pilot options, all using source `numId 1`. These values are recomputed from the input rather than trusted from candidate text.

## Adaptation boundary and leakage control

Version 0.3 supersedes the copy-and-rename behavior in v0.2. Apache-2.0 permits the derivative adaptation. OfficeBench contributes the source assignment, identity facts, original Query/evaluator record, and license. CoworkEvoBench contributes the identity-control-sheet workflow, exact protected-block preservation, paragraph evidence ledger, package/closure controls, custom Word properties, internal navigation, dedicated Oracle, rubric, and directed failures.

The task does not answer the homework, infer student performance, add grades, declare submission status, or invent instructor decisions. Its leakage group remains `officebench-2-30-rename`; its capability target is non-target-content preservation and identity-derived package control, unlike the CS161 roster/review/exception tasks and the OCR reconstruction tasks.

## Golden and scoring boundary

The only deliverable is `A12345678.docx`. The task-specific solver builds three editable Word pages: an identity/control sheet; the contiguous original 13-paragraph assignment block with source run boundaries, paragraph properties, and numbering preserved; and a 13-row preservation ledger plus independent closure. It also provides six native tables, repeating headers, non-splitting rows, a source-ID bookmark/internal link, six custom document properties, page header, and automatic page numbers.

The Judge reads neither Golden nor solver-only markers. It independently extracts the source identity, course, assignment, term, body hashes, runs, formatting signatures, paragraph properties, and numbering from the frozen input; it then validates the candidate filename, tables, original block, native OOXML objects, metadata, navigation, safe packaging, and LibreOffice rendering. XML canonicalization is used for semantically identical numbering-part comparisons so ZIP serializer byte order is not mistaken for a formatting change.

## Local validation and visual QA

`validate-task`, `validate-task-assets`, package `compileall`, JSON parsing, input-copy hashes, ownership, and `git diff --check` pass. Golden and five independently generated fresh replays are byte-identical, SHA256 `690e20d82d0a9628704b8ef6e42d5c82458567640aa9e8cb5036ff87b7102c41`. Five fresh Judge runs each score `1.0`; all result JSON files are byte-identical, SHA256 `e2e02c62eee6cde6e8cb9074444425d1e5b1b2c48e27a83527f4de7441b20135`.

Seventeen non-Golden cases all fail: alternate filename, corrupt, empty, flattened image, unexpected extra DOCX, and several pagination-breaking native-structure/content changes are `0.0`; wrong extracted identity is `0.30`; protected question mutation, removed protected paragraph, and merged protected run boundaries are `0.40`; original title formatting damage is `0.58`; a wrong ledger SHA or missing ledger row is `0.70`; the medium sample missing one redundant closure row is `0.76`; broken bookmark navigation and a missing custom property are `0.78`. The medium sample is strictly between zero and Golden.

LibreOffice 7.3.7.2 renders Golden and fresh as three portrait Letter pages. Golden/fresh, representative identity/content/format/ledger/closure/structure negatives, the flattened case, and both LibreOffice round-trips were checked page by page at original render resolution. There are no unintended blank pages, clipping, overlapping text or tables, off-page content, split ledger rows, missing fonts, or unreadable sections. The protected source page retains the exact visual assignment body; its lowest text ends above the automatic footer without collision.

LibreOffice can open, save, and export both valid documents as three readable pages while retaining all business text, six tables, six custom properties, the bookmark/link, and ordered page contract. LibreOffice 7.3 rewrites source paragraph properties and numbering definitions and changes one run-format signature, so the strict Judge scores each round-trip `0.58`; this interoperability loss is recorded rather than hidden. Normal solver/Word OOXML output retains the protected source formatting and scores `1.0`.

Official Harbor 0.21.0 Docker Oracle A/B ran against the same frozen task snapshot with `n_concurrent_trials=1`, `force_build=true`, full verifier, and no Agent network dependency. A job `docx-hardening-b8-identity-control-a-20260908` / trial `officebench-2-30-rename-student__3ZWtzzo` and B job `docx-hardening-b8-identity-control-b-20260908` / trial `officebench-2-30-rename-student__ExYy7u2` each scored `1.0` with zero exceptions. Both report task checksum `0a35816331fbfcb502ce1e73031b83131ab171ec4bc64ac0f640650d884e3f37` and digest `sha256:fb9c5e83034fc36f6722616a7a3e183076910fa5074b2609097ca2a1d3125677`. Their DOCX SHA256 is `690e20d82d0a9628704b8ef6e42d5c82458567640aa9e8cb5036ff87b7102c41`, Judge JSON SHA256 is `e2e02c62eee6cde6e8cb9074444425d1e5b1b2c48e27a83527f4de7441b20135`, and artifact manifest SHA256 is `3b2200d91f247c9edd3849a8c80e83402294fbabaf72c2f7e2523134acac9c2c` in both runs. Oracle success is not treated as real-Agent difficulty calibration; difficulty remains unassigned.
