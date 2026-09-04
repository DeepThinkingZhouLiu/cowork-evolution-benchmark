# Evidence ledger

| Evidence | Source | Local path | Role | SHA256 | Supports |
|---|---|---|---|---|---|
| E1 | Spreadsheet-RL `6f274cd4cbc27e656d7ddae7c3b591669866dc84`, task-job-2d603b0f50935350/input.xlsx | `data/input_files/account_form_input.xlsx` | Solver-visible state | `4bd7e6e036fe06ccabba76dbdb5cdc76d117f4ba4b36ca4b7b95b549baa76ea2` | All source rows and labels |
| E2 | Spreadsheet-RL task-job-2d603b0f50935350/instruction.json | `adaptation/source-instruction.json` | Original intent | n/a | User workflow and output semantics |
| E3 | Spreadsheet-RL task-job-2d603b0f50935350/target.xlsx | `solution/Golden/account_form_completed.xlsx` | Audit-only reference | `dd2726369e21e4010138d4deeade550eff93462c5c913621c4ef4d117855bcc2` | Independent cross-check of output shape |
| E4 | CC-BY-SA-4.0 text | `adaptation/LICENSE-CC-BY-SA-4.0.txt` | Rights evidence | `missing` | Attribution and ShareAlike basis |

The Judge uses E1 and stated transformation rules; E3 is never loaded during scoring.
