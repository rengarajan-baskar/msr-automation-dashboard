# MSR Automation (Python)

Automate Monthly Status Review (MSR) reporting from an Excel export.

## Features
- Reads any Excel tracker (XLSX) — no hardcoding of column names required.
- Configurable mappings via `config.yaml` (root-cause keywords, column aliases, output pivots).
- Summaries by Type, Priority, State, Category, Subcategory, RootCause, AssignmentGroup, Customer.
- Exports a formatted Excel report with multiple sheets.
- Optional charts.
- CLI with sensible defaults: 
  ```bash
  python msr_automator.py --input msr_sample.xlsx --outdir out --config config.yaml --charts
  ```

## Quickstart
1. `pip install -r requirements.txt`
2. `python msr_automator.py --input /path/to/your_tracker.xlsx --outdir out`
3. Open `out/MSR_Summary.xlsx`

## Notes
- Column names are matched case-insensitively using aliases in `config.yaml`.
- Root cause can be taken from a column if present or inferred from description via rules.
