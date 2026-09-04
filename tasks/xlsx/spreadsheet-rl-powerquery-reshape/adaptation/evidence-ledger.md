# Evidence ledger

| Evidence | Source | Local path | Role | SHA256 | Supports |
|---|---|---|---|---|---|
| E1 | Spreadsheet-RL `6f274cd4cbc27e656d7ddae7c3b591669866dc84`, task-job-93b8d5ae5b1f85f1/input.xlsx | `data/input_files/raw_header_input.xlsx` | Solver-visible state | `d85e231a8a07afa79789b0985b5fbed5a68590ad5a525a1e3dc1b5237b100222` | All source rows and labels |
| E2 | Spreadsheet-RL task-job-93b8d5ae5b1f85f1/instruction.json | `adaptation/source-instruction.json` | Original intent | n/a | User workflow and output semantics |
| E3 | Spreadsheet-RL task-job-93b8d5ae5b1f85f1/target.xlsx | `solution/Golden/shaped_table.xlsx` | Audit-only reference | `df5b7de274413be0b658a65e364da17a39672b1adc94450c254a099a868d5719` | Independent cross-check of output shape |
| E4 | CC-BY-SA-4.0 text | `adaptation/LICENSE-CC-BY-SA-4.0.txt` | Rights evidence | `missing` | Attribution and ShareAlike basis |

The Judge uses E1 and stated transformation rules; E3 is never loaded during scoring.
