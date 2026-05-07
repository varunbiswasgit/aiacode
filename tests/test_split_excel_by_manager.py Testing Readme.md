# test_split_excel_by_manager.py — Testing Readme

Automated pytest test suite for `scripts/split_excel_by_manager.py` (Python equivalent of `SplitExcelByManager.bas`).

## How to Run

From the repository root:

```bash
pytest tests/test_split_excel_by_manager.py -v
```

All 15 tests must pass against `scripts/split_excel_by_manager.py`.

## Requirements

```bash
pip install pandas openpyxl pytest
```

## Test Case Summary

| ID | Area | What it verifies |
|----|------|------------------|
| TC-PY-01 | `sanitize_filename` | Characters `/ : " \ \| ? *` replaced with `_` |
| TC-PY-02 | `sanitize_filename` | All 10 Windows reserved names (`CON`, `PRN`, `AUX`, `NUL`, `COM1`–`COM9`, `LPT1`–`LPT9`) mapped to `Unknown_Manager` |
| TC-PY-03 | `sanitize_filename` | Output length capped at 200 characters |
| TC-PY-04 | `sanitize_filename` | Normal names (no special chars) returned unchanged |
| TC-PY-05 | Input validation | Missing file returns `False` |
| TC-PY-06 | Input validation | Unsupported file type (`.csv`) returns `False` |
| TC-PY-07 | Column validation | Wrong column name returns `False` |
| TC-PY-08 | Data validation | Empty DataFrame returns `False` |
| TC-PY-09 | Split correctness | One output file per unique manager |
| TC-PY-10 | Split correctness | Correct row counts per manager file |
| TC-PY-11 | Split correctness | Blank / null manager rows skipped; only non-null managers produce files |
| TC-PY-12 | Configuration | Configurable manager column name works correctly |
| TC-PY-13 | Sanitization | Slashes in manager name sanitized in output filename |
| TC-PY-14 | File system | Nested output directory auto-created when absent |
| TC-PY-15 | Column widths | Column widths clamped between 8 and 50 in output files |

## Detailed Test Cases

### TC-PY-01 · Invalid characters replaced

| Field | Detail |
|-------|--------|
| Input | `sanitize_filename('Alice/Bob:Report')` |
| Expected | `'Alice_Bob_Report'` |
| Pass criteria | No `/` or `:` in output. |

### TC-PY-02 · Windows reserved names

| Field | Detail |
|-------|--------|
| Input | Each of `CON`, `PRN`, `AUX`, `NUL`, `COM1`–`COM9`, `LPT1`–`LPT9` (upper and lower case) |
| Expected | All return `'Unknown_Manager'` |
| Pass criteria | No reserved name passes through as a filename. |

### TC-PY-03 · Length cap

| Field | Detail |
|-------|--------|
| Input | String of 300 `A` characters |
| Expected | `len(result) <= 200` |

### TC-PY-04 · Normal name unchanged

| Field | Detail |
|-------|--------|
| Input | `'John Smith'` |
| Expected | `'John Smith'` |

### TC-PY-05 · Missing input file

| Field | Detail |
|-------|--------|
| Input | Path to a non-existent file |
| Expected | Returns `False`. No exception raised. |

### TC-PY-06 · Unsupported file type

| Field | Detail |
|-------|--------|
| Input | A `.csv` file |
| Expected | Returns `False` before attempting to read. |

### TC-PY-07 · Manager column not found

| Field | Detail |
|-------|--------|
| Setup | Excel file with headers `Supervisor`, `Score` |
| Input | `manager_column='Manager'` |
| Expected | Returns `False`. |

### TC-PY-08 · Empty DataFrame

| Field | Detail |
|-------|--------|
| Setup | Excel file with headers only, no data rows |
| Expected | Returns `False`. |

### TC-PY-09 · Standard split

| Field | Detail |
|-------|--------|
| Setup | 3 rows: Alice/90, Bob/85, Alice/92 |
| Expected | `Alice_report.xlsx` and `Bob_report.xlsx` in output dir |
| Pass criteria | Returns `True`. Both files exist. |

### TC-PY-10 · Correct row counts

| Field | Detail |
|-------|--------|
| Setup | 4 rows: Alice/90, Bob/85, Alice/92, Alice/88 |
| Expected | `Alice_report.xlsx` has 3 data rows; `Bob_report.xlsx` has 1 |

### TC-PY-11 · Blank manager rows skipped

| Field | Detail |
|-------|--------|
| Setup | Rows: Alice/90, `None`/85, `""`/88 |
| Expected | Only `Alice_report.xlsx` created |
| Pass criteria | `manager_reports` contains exactly 1 file. |

### TC-PY-12 · Configurable column name

| Field | Detail |
|-------|--------|
| Setup | Headers: `Supervisor`, `Score` |
| Input | `manager_column='Supervisor'` |
| Expected | Returns `True`. Files created per unique supervisor. |

### TC-PY-13 · Sanitized manager name in filename

| Field | Detail |
|-------|--------|
| Setup | Manager named `Alice/Bob` |
| Expected | Output file: `Alice_Bob_report.xlsx` |
| Pass criteria | File with `/` in name does not exist. |

### TC-PY-14 · Output directory auto-created

| Field | Detail |
|-------|--------|
| Setup | Output path does not exist (`new/nested/reports`) |
| Expected | Directory created; file saved inside |
| Pass criteria | `out_dir.exists() == True` after call. |

### TC-PY-15 · Column width clamped

| Field | Detail |
|-------|--------|
| Setup | Manager `Alice`. Column B contains a 200-character string. |
| Expected | All column widths in output: `8 <= width <= 50` |
| Pass criteria | openpyxl `column_dimensions` check passes for every column. |

## Cross-Reference

For the equivalent VBA manual test cases covering the same logic boundaries, see `tests/SplitExcelByManager.bas Testing Readme.md`.
