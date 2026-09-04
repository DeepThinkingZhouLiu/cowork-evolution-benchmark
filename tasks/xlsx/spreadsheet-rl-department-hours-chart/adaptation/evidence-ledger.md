# Evidence ledger

| Evidence | Source | Local path | Role | SHA256 | Supports |
|---|---|---|---|---|---|
| E1 | Spreadsheet-RL `6f274cd4cbc27e656d7ddae7c3b591669866dc84`, task-job-1813a28727a048cb/input.xlsx | `data/input_files/department_hours_input.xlsx` | Solver-visible state | `79b9fd32457920f7607ac721bf245013c05f07ffc97002af3a7c78b6373526ad` | All source rows and labels |
| E2 | Spreadsheet-RL task-job-1813a28727a048cb/instruction.json | `adaptation/source-instruction.json` | Original intent | n/a | User workflow and output semantics |
| E3 | Spreadsheet-RL task-job-1813a28727a048cb/target.xlsx | `solution/Golden/department_hours_analysis.xlsx` | Audit-only reference | `f4205998c41aa7bf8d04740b34db1a1576bec33aa2004e87b54d82b34244ad09` | Independent cross-check of output shape |
| E4 | CC-BY-SA-4.0 text | `adaptation/LICENSE-CC-BY-SA-4.0.txt` | Rights evidence | `missing` | Attribution and ShareAlike basis |

The Judge uses E1 and stated transformation rules; E3 is never loaded during scoring.
