# Excel → DOCX Aggregator

Extracts line items from messy Excel files, groups by **Description**, and exports a single formatted **.docx** table.
Handles header rows in row 2/3, small typos (case/spacing), and computes missing **Line Amount** or **Unit Price**.
Output is sorted by **Quantity (ascending)** and includes a **TOTAL** row at the bottom.

## Quick start (Windows)
1. Download the latest ZIP from **Releases**.
2. Unzip and double‑click `run.bat` (or drag an Excel file/folder onto it).
   - First run creates `venv` and installs from `requirements.txt` automatically.

## CLI (advanced)
```bash
# auto-detect newest Excel in current directory
python excel_to_docx_aggregate.py --dir .

# or specify a file
python excel_to_docx_aggregate.py input.xlsx -o output.docx --sheet "Sheet1"
```

## What it handles
- Header not always in the first row (auto-detects within top 10 rows)
- Small typos / case / extra spaces for: **Description**, **Quantity**, **Unit Price**, **Line Amount**
- Computes missing Line Amount (qty × price) or Unit Price (amount ÷ qty)
- Groups by Description, sums **Quantity** and **Line Amount**
- Sorts by **Quantity (ascending)**, appends a **TOTAL** row

## Requirements
- Python 3.9+
- See `requirements.txt`

## License
MIT — see `LICENSE`.
