# paper-metadata v0.3 adaptation record

Status: `local_validated_harbor_pending`; difficulty remains `unassigned`.

## Why this is a natural upgrade

- OfficeBench `1-19/0` asks only for a title and author list from `paper.pdf`; all three Sonnet 5 Harnesses produced a two-to-five-paragraph DOCX and received 1.0 under the tightened v0.2 Judge.
- The frozen PDF itself already contains a much richer, internally coherent research-intake problem: symbol-coded affiliations, an equal-contribution footnote, one listed contact, arXiv identity, an abstract with qualified claims, a multi-scenario Figure 1 caption, and blank/untagged embedded metadata.
- The v0.3 task therefore retains the same PDF and the original title/author extraction core. It does not append unrelated business facts or replace the seed. It asks the solver to resolve and audit only evidence already visible on the frozen page.

## Full context read

- Read the original instruction, one-page PDF, both v0.2 solution DOCX files, rubric, Judge, solver, expectations, task config, adaptation, metadata, OfficeBench Apache-2.0 license, and complete upstream Query/Evaluator.
- Parsed the whole PDF with `pdftotext -layout` and `pdfinfo`, and visually inspected its only page. The PDF has 1 Letter page, PDF 1.5, blank embedded Title/Author fields, and `Tagged: no`.
- Parsed both legacy DOCX files paragraph by paragraph and table by table. Each has two editable paragraphs and no tables.
- Read all three tightened rejudge JSON files and all three actual output DOCX files. Claude Code used five paragraphs; OpenCode and OpenHands each used two. All three contained only title/authors and scored 1.0.

## Source and license

- Dataset: OfficeBench; task `1-19/0`; immutable revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Input SHA256: `ab6d2a7e041b267c07e064c4acd1c65c557969b1a440e80e5b2ecbab70cbf228`.
- OfficeBench is Apache-2.0; the full license remains in `LICENSE-OfficeBench`. The public benchmark snapshot and original evaluator remain frozen in `source-task.json`.
- No external paper page, website, DOI service, venue database, or live network source is used. This prevents accidental expansion beyond the one-page evidence boundary.

## Truth design

- Exact displayed title and eight-author order are authoritative. Author-affiliation mapping follows only the four visible symbols: Shanghai AI Laboratory, East China Normal University, Princeton University, and The University of Hong Kong.
- The asterisk supports equal contribution for Zhiyong Wu and Chengcheng Han only. The single email is recorded as a listed contact, not converted into a corresponding-author claim.
- The abstract supports the OS-Copilot/FRIDAY relationship, five OS element classes, the reported GAIA 35% improvement, Excel/PowerPoint learning with minimal supervision, and the authors' stated research significance.
- Figure 1 supports exactly three scenarios and three reading rules. The one-page file does not support a named GAIA baseline, confidence interval, venue, DOI, publication status, methods detail, limitations, or full-paper conclusions.
- `evidence-ledger.md` is the author truth ledger. The task-specific Judge first parses the frozen PDF text/properties, then independently checks candidate facts, ordering, mappings, evidence matrix, comments, bookmarks, fields and section structure. It never reads Golden bytes or solver markers.

## Word-specific adaptation boundary

- The deliverable is one editable DOCX. PDF remains input evidence only; no cross-office state package is requested.
- Native objects are purposeful: six editable tables, three sections with a landscape evidence matrix, 12 bookmarks/internal links, three exact reviewer comments, a TOC field, core properties, repeating headings, non-splitting rows, and PAGE/NUMPAGES fields.
- Visual rendering is author-side usability verification only. No VLM or visual reward is added.

## Local validation (2026-09-08)

- `validate-task`, `validate-task-assets`, JSON parsing, task-specific compileall, data/environment byte comparison and `git diff --check` all pass.
- Golden and fresh replay each score 1.0 five times; their 24/24 uncompressed DOCX payload members are byte-identical. LibreOffice 7.3.7.2 round-trips of both also score 1.0 and have 22/22 identical payload members.
- Sixteen task-specific cases calibrate causal failures: missing/corrupt/extra output are 0; truncated title and wrong page count are 0.35; author order/affiliation defects are 0.55; the 35% error is 0.65; missing figure or evidence row is 0.72; fabricated DOI/corresponding author is 0.78; screenshot flattening is 0.171667. Removing comments, navigation, or landscape retains partial credit at 0.98, 0.95, and 0.994286 respectively.
- The source, Golden, fresh, both round-trips and four relevant negative renders total 23 pages; 23/23 pages were viewed. Positive outputs are stable three-page documents with a readable landscape matrix and no blank page, overlap, clipping, split row, missing glyph or unreadable content. The portrait-matrix negative visibly loses its rightmost status column, matching the structural deduction.
- Rendering evidence is author-side only and does not enter reward. Harbor Oracle A/B remains pending on the frozen task snapshot.
