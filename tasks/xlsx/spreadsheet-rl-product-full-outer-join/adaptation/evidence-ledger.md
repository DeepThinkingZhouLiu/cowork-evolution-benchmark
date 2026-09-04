# Evidence ledger

| Evidence | Source | Local path | Role | SHA256 | Supports |
|---|---|---|---|---|---|
| E1 | Spreadsheet-RL `6f274cd4cbc27e656d7ddae7c3b591669866dc84`, task-job-68594670fd5aafc2/input.xlsx | `data/input_files/product_sales_input.xlsx` | Solver-visible state | `5d6fad7006decb19d9eb960c5bbea931dfb5595f0d69d2a873e293648f838fb5` | All source rows and labels |
| E2 | Spreadsheet-RL task-job-68594670fd5aafc2/instruction.json | `adaptation/source-instruction.json` | Original intent | n/a | User workflow and output semantics |
| E3 | Spreadsheet-RL task-job-68594670fd5aafc2/target.xlsx | `solution/Golden/product_sales_outer_join.xlsx` | Audit-only reference | `ffa66fd34b203968ebafdf6f86a1c2dd99b27b7da4bc88c4d25871d5b1377f63` | Independent cross-check of output shape |
| E4 | CC-BY-SA-4.0 text | `adaptation/LICENSE-CC-BY-SA-4.0.txt` | Rights evidence | `missing` | Attribution and ShareAlike basis |

The Judge uses E1 and stated transformation rules; E3 is never loaded during scoring.
