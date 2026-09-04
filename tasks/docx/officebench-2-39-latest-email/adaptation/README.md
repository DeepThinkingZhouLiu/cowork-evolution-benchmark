# officebench-2-39-latest-email source adaptation record

Status: `pending_human_review`; difficulty is an author hypothesis only; `calibration_status=uncalibrated`.

## Upstream evidence

- Dataset: OfficeBench, task `2-39/0`, immutable revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Original query (copied from `adaptation/source-task.json`): `find Bob's latest email, save it in latest.docx`
- The complete upstream evaluator is retained in `adaptation/source-task.json`; its positive keywords are: （上游 evaluator 未提供关键词；以原始 Query 和可见输入为真值）.
- Its negative keywords are: （上游 evaluator 未提供排除关键词）.
- The source license is copied to `adaptation/LICENSE-OfficeBench`.

## Output decision

The retained source evaluator names only DOCX outputs, so this package remains in `docx`. Solver deliverables are: `latest.docx`. No PDF, XLSX, EML, ICS, image, or TXT deliverable is silently removed. If a source task ever requires one of those formats, this script classifies it as `route_to_cross_office` instead.

## Input evidence

| Package input | SHA256 checked by this repair | Role |
|---|---|---|
| `data/input_files/rental.eml` | `54f31902b3b4b18a1fefba0e4298ed8f1f2205d87595b99f22e5affa09acd807` | `solver_input` |
| `data/input_files/scholarship approved.eml` | `d0f0ef57930c7a9f31252f81b81453669aa8e4fd52c58a62cc4935b4fadbd614` | `solver_input` |
| `data/input_files/meeting.eml` | `1f2476e2467e6fc0b98e97a03c666e1e018469ad0490b512195688644fd14769` | `solver_input` |
| `data/input_files/gradescope.eml` | `0b61b4197e0eb9e88c01732c66ea92df320a8304743a17e8e04aee6baf47a401` | `solver_input` |

## Truth and Judge

Golden generation is kept separate from scoring. The source task JSON and visible package inputs are the evidence basis; the Judge reads the package inputs at evaluation time and applies the source-required terms plus task-specific structural checks. It does not compare a ZIP hash to the Golden. The original evaluator is preserved for audit only and is not treated as sufficient semantic validation.

## Known limitations and review gates

- OfficeBench fixtures are benchmark snapshots and are marked `synthetic=true`; this is a provenance label, not a claim that the source story was invented here.
- Several source families share files or near-identical queries; the `leakage_group_id` in metadata intentionally groups them for split control.
- Human review must inspect the source preview, Golden text, editable structure, pagination/rendering, and empty/corrupt/unchanged/core-error negative cases before acceptance.
