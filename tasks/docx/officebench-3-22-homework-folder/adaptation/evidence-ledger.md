# Evidence ledger

## Upstream files read in full

| ID | Frozen evidence | SHA256 | Finding / contribution |
|---|---|---|---|
| E01 | OfficeBench `tasks/3-22/subtasks/0.json` and packaged `source-task.json` | `da4167bc38e730b88c09cc5cc37e7d99a7b0612cd5dd9a726597967e3ce2b3aa` | Exact two-path query; evaluator checks existence only. |
| E02 | OfficeBench `tasks/3-22/testbed/data/homework.jpg` and both packaged input copies | `1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf` | Sole runtime source; 1700 x 2200 JPEG. |
| E03 | OfficeBench `LICENSE` and packaged `LICENSE-OfficeBench` | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4` | Apache-2.0 permits adaptation with notice retained. |
| E04 | OfficeBench `README.md` | `04104d99107c2f109ea57d06f75541051efd4e7b5a8b2e41083315b3d21d5ed2` | Dataset context and benchmark scope. |
| E05 | OfficeBench `utils/evaluate.py` | `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90` | Confirms file-existence evaluator semantics; not reused as semantic truth. |

Reference checkout HEAD was verified as `b978b808667c32b52ce19a67ce1def1de9ae02b7`. E01-E03 were byte-compared with their packaged copies.

## Package files read before remake

| ID | Prior artifact | Finding |
|---|---|---|
| P01 | Original `instruction.md`, `task.toml`, `rubric.json`, `judge.py`, `expectations.json`, `metadata/*`, `adaptation/*` | Contract largely inherited the upstream existence check and did not prove semantic or Word-native quality. |
| P02 | Original Golden root and nested DOCX | Byte-identical copies (`e7076…`), each a noisy single OCR paragraph without tables, comments, controls, or role separation; LibreOffice showed three poorly split pages. |
| P03 | Original solution root and nested DOCX and generic solver | Reproduced the same duplicate artifact and lacked a governed two-copy workflow. |

All original package inputs and all four prior DOCX artifacts were opened structurally; all three pages of the prior Golden were rendered and inspected before redesign.

## Authoritative truth and invariants

| ID | Evidence | SHA256 / value | Judge use |
|---|---|---|---|
| T01 | `adaptation/package-truth.json`; identical copies in `solution/` and `tests/` | `271e323474d9532977d4f392bef3012daf9b48e7de2eab42da6263d59b51113e` | Exact source-visible header, instructions, three questions, ambiguity notes, equation, and six-row 15-cycle pipeline. |
| T02 | Canonical payload digest | `83425e9072d022c0a797826c1d6ebf38884fa59ffb4a638c88f8cd9903cf9848` | Independently recomputed; must match both custom-property ledgers and source payloads. |
| T03 | `instruction.md` | `33328989bc223789b2393856d10e776f376fdc98392a68b6bf84666fa7498f3c` | Public two-role output and editability contract. |
| T04 | `rubric.json` | `76ed2b1778d0fa555e9d6f88cf87ff6e76b5e94a077b228c8713f478895f992a` | Eight task-specific criteria with entity-level partial scoring and causal caps. |
| T05 | Task-specific `solution/solve.py` | `35641bbf5237d540049cfe7541b633647e89597ae0dae6135522c37a818661f1` | Builds native role-specific Word objects; not imported by Judge. |
| T06 | Task-specific `tests/judge.py` | `40e3077f6a954e8be8d4f87e52cbb94758e8264928c87bd1130811c2169608b4` | Independently recomputes truth and validates OOXML, semantics, roles, controls, comments, and render order. |

Key ambiguity decisions are source-preserving: the absent second diagram is never manufactured; `M ↓` versus `W ↓` and the `M1`/`M_l` conflict remain disclosed; the timing equation is a native editable OMML object in each copy.

## Deterministic outputs and replay

| Check | Result |
|---|---|
| Golden master | `homework.docx`, SHA256 `b6ef6ea83066f97142d5a2dadac155bf94816291d2ce36518d39290f5a4c9df0`. |
| Golden review | `homework/homework.docx`, SHA256 `fd2006d51dba1158d6cdd5307288902e246af92fd96ea1110a2aa3ec576907b2`. |
| Fresh determinism | Five independent builds matched the two Golden hashes byte-for-byte. |
| Judge determinism | Five Golden plus five fresh evaluations scored `1.000000`; every final result JSON had SHA256 `32ea9cb6f32c2dfe2a11169bbd99966060b1a7f0c6d2927a0ee1c14e451a2173`. |
| Native package | Master has 17 safe OOXML parts and eight tables; review has 18 safe parts and seven tables. Neither contains media, embedded binaries, macros, or external relationships. |

## Negative calibration

The final suite contains 23 independently constructed artifacts. All returned non-passing exit status.

| Failure family | Cases | Observed score |
|---|---|---|
| Unavailable / invalid package | `empty_output`, `corrupt_master`, `missing_review_copy`, `flattened_pair`, `unexpected_extra_docx` | `0.00` |
| Fabricated or cross-copy-wrong source content | `fabricated_review_diagram`, `unsynchronized_review_student_id` | `0.32` |
| Pipeline or equation damage | `wrong_master_pipeline_stage`, `missing_review_native_equation` | `0.50` |
| Role/property synchronization damage | `wrong_master_role_property`, `wrong_review_copy_path_property`, `wrong_review_payload_digest` | `0.62` |
| Master governance damage | `missing_master_payload_block`, `wrong_master_exception_treatment` | `0.72` |
| Review evidence damage / medium sample | `missing_review_comment`, `wrong_comment_author`, `wrong_review_issue_disposition`, `medium_incomplete_review_queue` | `0.78` |
| Review control damage | `locked_qa_status_control`, `missing_qa_reviewer_control` | `0.82` |
| Fine structural damage | `broken_master_navigation`, `missing_review_role_header`, `missing_reviewer_boundary_note` | `0.996` |

The `0.996` structural cases intentionally demonstrate fine-grained partial scoring but still fail the `all criteria perfect` pass gate. The medium sample removes part of the source-bound review queue while keeping an otherwise strong package, so `0.78` lies strictly between empty and Golden.

## LibreOffice visual and roundtrip evidence

| Artifact | Pages | SHA256 / outcome | Inspection |
|---|---:|---|---|
| Golden master PDF | 4 | `296b46e189d763583a7a52e5ae14f276489a12c5920d0f21a39461cc082ac04c` | All pages checked at original resolution; clean tables, header/footer, readable type, no clipping or blank page. |
| Golden review PDF | 4 | `776a7e01c1d3f2cc8a1e223431fff92be979de49ebe46cbcec182c4f84f6a280` | All pages checked; comments/queue/control inventory and handoff fit without overflow. |
| Roundtrip master DOCX/PDF | 4 | DOCX `9483054e6d167ccc6e01b01ff7c1369d4a39d53932bf8485dee4c789a252561f`; PDF `e07f91fb3209455e3a20b60788a23eb12d3d75c792ca5d829dda4589c9657212` | Open/save/export retained layout and semantics. |
| Roundtrip review DOCX/PDF | 4 | DOCX `20802a210f3e9c9d3c0d9e2f68be668fd134644b0de3ab77c9c90f1ab1536e3d`; PDF `dd96f90ca6f70275b843ac5f932761d3c91bb3fa8d089704281f88f6ffb38f59` | Open/save/export retained comments and controls; LibreOffice cleared only `w:tag` values. |
| Strict roundtrip Judge | — | `0.820000`; result JSON `a2e19d63a447e783f614f322826a8f7f58a372dfc7b05b57d0a1fb52b2be48fd` | R007 exposed the tag loss; requirement was not relaxed. |

Nine causal negative packages were also rendered. The source-page, register, queue, control, role-header, handoff, and flattened pages were visually inspected; mutations remained readable and localized, while the flattened substitute collapsed into two unusable pages as intended.

## Harbor evidence

| Replay | Job / trial | Reward / errors | Frozen evidence |
|---|---|---|---|
| A | `docx-hardening-b8-controlled-ocr-package-a-20260908` / `officebench-3-22-homework-folder__zBNdcot` | `1.0` / `0` | Job result `2a056d6512c28326e8227d00cbe328290f580487a987fca9d9a3831cc03118fa`; task digest `sha256:cc5f80a2b03981aa8d667cd597349145bf2a48d046bd8edb8fad6b2c57f83355`. |
| B | `docx-hardening-b8-controlled-ocr-package-b-20260908` / `officebench-3-22-homework-folder__BjjXqvp` | `1.0` / `0` | Job result `6d82ed97476709e87a7ff291f986d84aedf76d2a94dd8f29a3c62c5a2b81615d`; same task digest. |

Both replays produced manifest SHA256 `3b2200d91f247c9edd3849a8c80e83402294fbabaf72c2f7e2523134acac9c2c`, Judge result SHA256 `32ea9cb6f32c2dfe2a11169bbd99966060b1a7f0c6d2927a0ee1c14e451a2173`, master SHA256 `b6ef6ea83066f97142d5a2dadac155bf94816291d2ce36518d39290f5a4c9df0`, and review SHA256 `fd2006d51dba1158d6cdd5307288902e246af92fd96ea1110a2aa3ec576907b2`. Each artifact set contains the two required DOCX files plus Judge evidence; the business payload is stable across A/B.
