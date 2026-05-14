# Split Excel by Manager — Testing Guide

This document covers test procedures for both `SplitExcelByManager.bas` (VBA) and `split_excel_by_manager.py` (Python).

---

## Python Automated Tests

Run from the repository root:

```bash
pytest tests/test_split_excel_by_manager.py -v
```

All 15 tests must pass.

### Requirements

```bash
pip install pandas openpyxl pytest
```

### Test Case Summary

| ID | Area | What it verifies |
|----|------|------------------|
| TC-PY-01 | `sanitize_filename` | Characters `/ : " \ \| ? *` replaced with `_` |
| TC-PY-02 | `sanitize_filename` | All 10 Windows reserved names mapped to `Unknown_Manager` |
| TC-PY-03 | `sanitize_filename` | Output length capped at 200 characters |
| TC-PY-04 | `sanitize_filename` | Normal names returned unchanged |
| TC-PY-05 | Input validation | Missing file returns `False` |
| TC-PY-06 | Input validation | Unsupported file type (`.csv`) returns `False` |
| TC-PY-07 | Column validation | Wrong column name returns `False` |
| TC-PY-08 | Data validation | Empty DataFrame returns `False` |
| TC-PY-09 | Split correctness | One output file per unique manager |
| TC-PY-10 | Split correctness | Correct row counts per manager file |
| TC-PY-11 | Split correctness | Blank / null manager rows skipped |
| TC-PY-12 | Configuration | Configurable manager column name works correctly |
| TC-PY-13 | Sanitization | Slashes in manager name sanitized in output filename |
| TC-PY-14 | File system | Nested output directory auto-created when absent |
| TC-PY-15 | Column widths | Column widths clamped between 8 and 50 |

---

## VBA Manual Tests

Environment setup:

1. Open a new Excel workbook with sample data and **save it to disk**.
2. Press **Alt + F11**, import `SplitExcelByManager.bas`.
3. Run `SplitExcelByManager` from **Developer → Macros**.

### TC-VBA-01 · Standard split — one file per unique manager

| Field | Detail |
|-------|--------|
| Setup | Headers: `Manager`, `Score`. Rows: Alice/90, Bob/85, Alice/92. |
| Expected | `Alice_report.xlsx` and `Bob_report.xlsx` in `manager_reports`. MsgBox: **"Done! 2/2 report(s) saved"** |
| Pass criteria | Both files exist. Alice file has 2 data rows; Bob file has 1. |

### TC-VBA-02 · Correct row distribution per manager

| Field | Detail |
|-------|--------|
| Setup | 4 rows: Alice/90, Bob/85, Alice/92, Alice/88 |
| Expected | `Alice_report.xlsx` has 3 data rows; `Bob_report.xlsx` has 1 |
| Pass criteria | Open each output file and count rows below header. |

### TC-VBA-03 · Workbook not saved — guard fires

| Field | Detail |
|-------|--------|
| Setup | New unsaved workbook (`ThisWorkbook.Path = ""`) |
| Expected | MsgBox: **"Please save the workbook first…"** Macro exits. |
| Pass criteria | No `manager_reports` folder created. |

### TC-VBA-04 · Manager column not found

| Field | Detail |
|-------|--------|
| Setup | Headers: `Supervisor`, `Score` — no `Manager` column |
| Expected | MsgBox: **"Error: Column 'Manager' not found in row 1."** |
| Pass criteria | No output files created. |

### TC-VBA-05 · Empty data below header

| Field | Detail |
|-------|--------|
| Setup | Row 1: `Manager`, `Score`. No data rows. |
| Expected | MsgBox: **"Error: No data rows found…"** |
| Pass criteria | No output files created. |

### TC-VBA-06 · Blank manager cells skipped

| Field | Detail |
|-------|--------|
| Setup | Rows: Alice/90, *(blank)*/85, Alice/88 |
| Expected | Only `Alice_report.xlsx` created |
| Pass criteria | `manager_reports` contains exactly one file. |

### TC-VBA-07 · Invalid characters in manager name sanitized

| Field | Detail |
|-------|--------|
| Setup | Manager named `Alice/Bob` |
| Expected | Output file: `Alice_Bob_report.xlsx` |
| Pass criteria | File with slash in name does not exist. |

### TC-VBA-08 · Windows reserved name replaced

| Field | Detail |
|-------|--------|
| Setup | Manager named `CON` |
| Expected | Output file: `Unknown_Manager_report.xlsx` |
| Pass criteria | No file named `CON_report.xlsx`. |

### TC-VBA-09 · Configurable manager column name

| Field | Detail |
|-------|--------|
| Setup | Change constant `MANAGER_COL_NAME = "Supervisor"`. Headers: `Supervisor`, `Score`. |
| Expected | Two report files created, one per supervisor |
| Pass criteria | Files present. No column-not-found error. |

### TC-VBA-10 · Column width clamped between 8 and 50

| Field | Detail |
|-------|--------|
| Setup | Manager `Alice`. Column B contains a 200-character string. |
| Expected | All column widths ≥ 8 and ≤ 50 in output file |
| Pass criteria | Check Format → Column → Width for each column. |

### TC-VBA-11 · Output directory created when absent

| Field | Detail |
|-------|--------|
| Setup | Delete `manager_reports` folder. Sheet with 1 manager row. |
| Expected | Folder created automatically; report saved inside |
| Pass criteria | Folder and file both exist after macro. |

### TC-VBA-12 · Existing AutoFilter reset before and after

| Field | Detail |
|-------|--------|
| Setup | Manually apply an AutoFilter before running |
| Expected | AutoFilter cleared at start and on exit |
| Pass criteria | `ws.AutoFilterMode = False` after macro. All rows visible. |

### TC-VBA-13 · File naming matches Python v3.0 output

| Field | Detail |
|-------|--------|
| Setup | Run both `.bas` and `.py` on the same dataset with default settings |
| Expected | Output file names are identical between both tools |
| Pass criteria | File name sets match exactly. |
