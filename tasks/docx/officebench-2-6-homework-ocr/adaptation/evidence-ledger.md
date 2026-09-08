# Evidence ledger — editable OCR reconstruction

## Files read before redesign

| ID  | Frozen item                                                                                  | SHA256                                                            | Finding / contribution                                                                                           |
|-----|-----------------------------------------------------------------------------------------------|-------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------|
| E01 | OfficeBench `tasks/2-6/subtasks/0.json` / packaged `source-task.json`                          | `6bac6e9c1fa45ed7101d05c19a436c5b4579078e9822a47886e5e238cbaf3062` | Exact query and three-keyword evaluator; evaluator is inadequate for full reconstruction fidelity.               |
| E02 | OfficeBench `tasks/2-6/testbed/data/homework.jpg`                                              | `1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf` | One 1700 × 2200 JPEG containing the full visible homework page and pipeline matrix.                               |
| E03 | OfficeBench `utils/evaluate.py`                                                                | `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90` | Full `evaluate_contain` implementation read; confirms substring presence rather than structural OCR validation. |
| E04 | OfficeBench `README.md`                                                                        | `04104d99107c2f109ea57d06f75541051efd4e7b5a8b2e41083315b3d21d5ed2` | Dataset purpose and execution context.                                                                           |
| E05 | OfficeBench `LICENSE` / packaged `LICENSE-OfficeBench`                                         | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4` | Apache-2.0 permits this derivative benchmark adaptation with attribution and retained license.                   |
| E06 | v0.2 instruction, inputs, Golden, solution, solver, Judge, rubric, adaptation, and metadata    | recorded in Git history                                            | Revealed single-paragraph OCR, generic multi-task solver, weak keyword checks, and missing traceability.         |

Upstream repository revision: `b978b808667c32b52ce19a67ce1def1de9ae02b7`.

## Verified visible truth

- Header: `CSE 240A Homework 2`; `Pipelines and Memory`; `Ryan Wang ID:A012345`; `Due June 1, 2021`.
- Instructions: five exact bullets, including the source's contradictory claim of two questions.
- Section heading: visibly duplicated numbering, `1    1: Pipelining (45 pts)`.
- Questions: three exact prompts; the note about not filling every cycle appears twice.
- Pipeline: six instruction rows and cycles 1–15. Occupied cells include repeated decode/fetch stalls; the visible down arrows occur at `lw` W@5, `add R3` W@8, and `add R2` M@9.
- Timing: native equation `11 * 1.75 = 19.25ns > 14 * 1.21 = 16.94ns`; conclusion `So 7-stage processor is faster.`
- Conflicts: declared question count 2 versus visible prompts 1/2/3; Question 2 names a diagram below, but no second diagram is visible.

The complete machine-readable ledger is `transcription-truth.json`, SHA256 `e918036d4952c4586b44209e84020fbb012c9563a6ac051b020d7e31a078d3dc`. Its segment hashes intentionally preserve raw punctuation and spacing even where Word display collapses whitespace.

## Runtime and output integrity

| Item                                   | SHA256                                                            | Check                                                                                 |
|----------------------------------------|-------------------------------------------------------------------|---------------------------------------------------------------------------------------|
| packaged input                         | `1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf` | Byte-identical to environment copy and upstream E02.                                  |
| Golden / solution / five fresh replays | `298605033d3a28ced8225da7062fc3ec69338874a1cdc641ec9d0512153e8268` | Deterministic three-page editable DOCX; all score 1.0.                                 |
| solution/test truth snapshot           | `f9ac0f3acea88968d77563bdf4dde0aaf860d2a88d597f4581da682d6aa702cf` | Byte-identical evaluator truth; semantic match after dropping `schema_version`.       |
| final LibreOffice PDF                  | `56ad54d049b8c5aedf86b2f70a5355d0a52b3c0cbeb47d8678db836eae408f29` | Three portrait Letter pages inspected in full.                                        |
| LibreOffice round-trip DOCX / PDF      | `a771e653528179464dd89a2d1b12ef08296a720844c7ae302201676438ef52f0` / `43aedcecdca9e5d731593af02dc07c04b93f19386b5009a48ab6b8de7ee960ec` | Three readable pages; semantic Judge score 1.0. |

## Judge coverage and causal calibration

R001 validates the frozen JPEG, only-one-DOCX contract, native structure, OOXML safety, and ordered render contract. R002–R006 score exact visible segments, native equation, pipeline cells, fingerprints, audit relations, conflicts, conventions, and closure rows. R007 checks OMML, bookmarks, internal hyperlinks, core metadata, and six custom properties. R008 checks headers/footers, repeating non-splitting tables, heading order, prohibited claims, and final rendering.

Hurdles/caps are attached to causal failure boundaries: unusable/non-native output 0.00; source header/instruction failure 0.30; question/equation failure 0.42; pipeline failure 0.55; evidence-ledger failure 0.68; audit/conflict failure 0.76; native-object/navigation failure 0.80; structural/render failure 0.82. The medium-quality sample omits only the custom-property closure row and receives 0.76, strictly between empty output and Golden.

## Visual inspection ledger

- Golden pages 1–3: exact reconstruction remains on page 1; the 7 × 16 table fits without horizontal clipping; evidence ledger fits page 2; audit/conflict/closure tables fit page 3; no extra blank page.
- Round-trip pages 1–3: no crop, overlap, table split, blank page, or unreadable content. LibreOffice normalizes the equation multiplication glyph and displays the repeating header only on its even-page style; these are documented interoperability changes and do not alter task facts.
- Directed negatives inspected: wrong Student ID and M1/Ml on page 1; missing dependency arrow on page 1; wrong fingerprint on page 2; missing closure row and invented diagram policy on page 3; flattened cheat on both of its invalid pages.

## Official Harbor replay evidence

Official Harbor Docker jobs `docx-hardening-b8-ocr-reconstruction-a-20260908` and `docx-hardening-b8-ocr-reconstruction-b-20260908` each completed one Oracle trial at reward `1.0` with zero errors. Trial IDs are `officebench-2-6-homework-ocr__zmGmBAs` and `officebench-2-6-homework-ocr__3y2kPJ5`. Both lock files report digest `sha256:e9bc27302500ad6355b37e757c708092423c752bdd16c902b97aefd409c186d1`; output DOCX, artifact manifest, Judge payload (`c0b63a8…5210`), and reward payload are byte-identical. The manifest captures `/app/output` with status `ok`; the optional log artifact directory is present and empty as expected.

Acceptance remains limited to authoring and Oracle reproducibility. No paid agent baseline has been run, and difficulty is still `unassigned`.
