# Split Excel by Manager — Python Script

This script reads an Excel file and writes one output workbook per unique manager into a dedicated folder (default: `manager_reports/`). It supports `.xlsx`, `.xls`, and `.xlsm` input files and handles common edge cases such as locked files, missing columns, blank manager names, and file-system-unsafe characters.

---

## Requirements

```bash
pip install pandas openpyxl
```

> `pathlib` and `re` are part of the Python standard library — no separate install needed.

---

## Usage

```bash
python split_excel_by_manager.py <file> [column] [output_dir]
```

### Arguments

| Argument | Required | Default | Description |
|---|---|---|---|
| `file` | Yes | — | Path to the input Excel file |
| `column` | No | `Manager` | Header name of the manager column |
| `output_dir` | No | `manager_reports` | Folder where reports are written |

### Examples

```bash
# Default 'Manager' column, output to manager_reports/
python split_excel_by_manager.py employees.xlsx

# Custom column name
python split_excel_by_manager.py staff.xlsx Supervisor

# Custom column and output folder
python split_excel_by_manager.py staff.xlsx Supervisor ./output

# Path with spaces
python split_excel_by_manager.py "C:/Data Files/employee list.xlsx" "Team Lead"
```

---

## Output Structure

```
project_directory/
├── split_excel_by_manager.py
├── employees.xlsx
└── manager_reports/
    ├── John_Smith_report.xlsx
    ├── Sarah_Johnson_report.xlsx
    └── ...
```

- File name pattern: `{sanitized_manager_name}_report.xlsx`
- Special characters (`/ \ : * ? < > |`) replaced with `_`
- Windows reserved names (CON, PRN, AUX, NUL, COM1–9, LPT1–9) replaced with `Unknown_Manager`
- Names truncated to 200 characters

---

## Column Width Handling

Each output file is post-processed to auto-fit column widths:

| Setting | Value |
|---|---|
| Minimum width | 8 characters |
| Maximum width | 50 characters |
| Method | Content-length scan via `openpyxl` |

---

## Error Handling

| Scenario | Behaviour |
|---|---|
| Input file not found | Prints error, exits |
| Unsupported file extension | Prints error, exits |
| File locked in Excel | Prints error, exits |
| Manager column missing | Prints error with available column names, exits |
| All manager values null | Prints error, exits |
| Output folder creation fails | Prints error, exits |
| Single manager write fails | Skips that manager, continues others |
| Column auto-fit fails | Warns, still counts file as written |
| Blank/null manager row | Skipped silently |

---

## Sample Console Output

```
Split Excel by Manager
========================================
Input file    : employees.xlsx
Manager column: Manager
Output folder : manager_reports
----------------------------------------
Reading: employees.xlsx
Found 3 unique manager(s).
  Writing: 'John Smith' -> manager_reports/John_Smith_report.xlsx
  Writing: 'Sarah Johnson' -> manager_reports/Sarah_Johnson_report.xlsx
  Writing: 'Michael Brown' -> manager_reports/Michael_Brown_report.xlsx

Done. 3/3 report(s) saved to: /home/user/project/manager_reports

✅ Completed successfully.
```

---

## Troubleshooting

**"Column not found"** — Check exact spelling and capitalisation. Column names are case-sensitive. The error message lists all available columns.

**"File is locked"** — Close the file in Excel before running the script.

**"Permission denied" on output folder** — Run from a directory where you have write access, or specify a writable `output_dir`.

**No output files created** — Verify the manager column contains non-blank values. Run without arguments to see the usage prompt.

---

## Version History

### v3.0
- `output_dir` added as an optional third CLI argument (default: `manager_reports`)
- `autofit_columns` extracted as a standalone reusable helper
- Full Windows reserved device name list in `sanitize_filename` (COM1–9, LPT1–9)
- `Path`-based file handling throughout (replaces `os.path`)
- Simplified `groupby` loop — blank managers skipped inline
- Redundant nested `try/except` blocks removed

### v2.0
- Comprehensive error handling added
- Filename sanitization
- Organised output into `manager_reports/` directory
- Real-time progress reporting

### v1.0
- Basic Excel split by manager column
- Column width auto-fit
