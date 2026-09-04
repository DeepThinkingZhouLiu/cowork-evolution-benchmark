# Evidence ledger

| Evidence | Source | Local path | Role | SHA256 | Supports |
|---|---|---|---|---|---|
| E1 | Spreadsheet-RL `6f274cd4cbc27e656d7ddae7c3b591669866dc84`, task-job-f59390cb071bb9fd/input.xlsx | `data/input_files/archer_scores_input.xlsx` | Solver-visible state | `88950f8d641c7bc680f1e3e87e090504158bdd8e072077445c51799c3c4efa8e` | All source rows and labels |
| E2 | Spreadsheet-RL task-job-f59390cb071bb9fd/instruction.json | `adaptation/source-instruction.json` | Original intent | n/a | User workflow and output semantics |
| E3 | Spreadsheet-RL task-job-f59390cb071bb9fd/target.xlsx | `solution/Golden/archer_scores_ranked.xlsx` | Audit-only reference | `247abb90b841702ca82add3095abd6650de2f54d60f1700e85798ecd5e88a0bc` | Independent cross-check of output shape |
| E4 | CC-BY-SA-4.0 text | `adaptation/LICENSE-CC-BY-SA-4.0.txt` | Rights evidence | `missing` | Attribution and ShareAlike basis |

The Judge uses E1 and stated transformation rules; E3 is never loaded during scoring.
