# SplitExcelByManager.bas — Manual Test Cases

All tests target `SplitExcelByManager` in `scripts/SplitExcelByManager.bas`.  
Run each test on a **fresh copy** of the described workbook. Pass/fail criteria are explicit.

---

## Environment Setup

1. Open a new Excel workbook containing sample data and **save it to disk** first.
2. Press **Alt + F11**, import `SplitExcelByManager.bas` from the `scripts/` folder.
3. Close the VBA editor.
4. Run `SplitExcelByManager` from **Developer → Macros**.

---

## TC-VBA-01 · Standard split — one file per unique manager

| Field | Detail |
|-------|--------|
| Setup | Sheet with headers in row 1: `Manager`, `Score`. Three data rows: Alice/90, Bob/85, Alice/92. |
| Input | Run macro. |
| Expected | Two files created: `Alice_report.xlsx` and `Bob_report.xlsx` in `manager_reports` subfolder. MsgBox reads **"Done! 2/2 report(s) saved to: …"** |
| Pass criteria | Both files exist. `Alice_report.xlsx` contains 2 data rows + header. `Bob_report.xlsx` contains 1 data row + header. |

---

## TC-VBA-02 · Correct row distribution per manager

| Field | Detail |
|-------|--------|
| Setup | 4 data rows: Alice/90, Bob/85, Alice/92, Alice/88. |
| Input | Run macro. |
| Expected | `Alice_report.xlsx` has 3 data rows. `Bob_report.xlsx` has 1 data row. |
| Pass criteria | Open each output file and count rows below the header. |

---

## TC-VBA-03 · Workbook not saved — guard fires

| Field | Detail |
|-------|--------|
| Setup | New unsaved workbook (`ThisWorkbook.Path = ""`). |
| Input | Run macro. |
| Expected | MsgBox: **"Please save the workbook first before running this macro."** Macro exits. |
| Pass criteria | No `manager_reports` folder created. |

---

## TC-VBA-04 · Manager column not found

| Field | Detail |
|-------|--------|
| Setup | Sheet with headers `Supervisor`, `Score` — no column named `Manager`. |
| Input | Run macro (default `MANAGER_COL_NAME = "Manager"`). |
| Expected | MsgBox: **"Error: Column 'Manager' not found in row 1."** Macro exits. |
| Pass criteria | No output files created. |

---

## TC-VBA-05 · Empty data below header

| Field | Detail |
|-------|--------|
| Setup | Row 1: `Manager`, `Score`. No data rows. |
| Input | Run macro. |
| Expected | MsgBox: **"Error: No data rows found…"** Macro exits. |
| Pass criteria | No output files created. |

---

## TC-VBA-06 · Blank manager cells skipped

| Field | Detail |
|-------|--------|
| Setup | Three rows: Alice/90, *(blank manager)*/85, Alice/88. |
| Input | Run macro. |
| Expected | Only `Alice_report.xlsx` created. Blank-manager row does not produce a file. |
| Pass criteria | `manager_reports` contains exactly one file. |

---

## TC-VBA-07 · Invalid characters in manager name sanitized

| Field | Detail |
|-------|--------|
| Setup | One manager named `Alice/Bob`. |
| Input | Run macro. |
| Expected | Output file is `Alice_Bob_report.xlsx` (slash replaced by underscore). |
| Pass criteria | File `Alice_Bob_report.xlsx` exists. No file with a slash in its name. |

---

## TC-VBA-08 · Windows reserved name replaced

| Field | Detail |
|-------|--------|
| Setup | One manager named `CON`. |
| Input | Run macro. |
| Expected | Output file is `Unknown_Manager_report.xlsx`. |
| Pass criteria | `Unknown_Manager_report.xlsx` exists. No file named `CON_report.xlsx`. |

---

## TC-VBA-09 · Configurable manager column name

| Field | Detail |
|-------|--------|
| Setup | Change `MANAGER_COL_NAME` constant to `"Supervisor"`. Sheet has headers `Supervisor`, `Score` with two managers. |
| Input | Run macro. |
| Expected | Two report files created, one per unique supervisor. |
| Pass criteria | Files present. No "column not found" error. |

---

## TC-VBA-10 · Column width clamped between 8 and 50

| Field | Detail |
|-------|--------|
| Setup | One manager `Alice`. Column B contains a 200-character string. |
| Input | Run macro. |
| Expected | In `Alice_report.xlsx`, all column widths are ≥ 8 and ≤ 50. |
| Pass criteria | Open output file. Check `Format → Column → Width` for each column. |

---

## TC-VBA-11 · Output directory created when absent

| Field | Detail |
|-------|--------|
| Setup | Delete the `manager_reports` folder if it exists. Sheet with 1 manager row. |
| Input | Run macro. |
| Expected | `manager_reports` folder created automatically. Report file saved inside. |
| Pass criteria | Folder and file both exist after macro. |

---

## TC-VBA-12 · Existing AutoFilter reset before and after

| Field | Detail |
|-------|--------|
| Setup | Manually apply an AutoFilter on the sheet before running. |
| Input | Run macro. |
| Expected | Macro resets AutoFilter at start and clears it on exit. No stale filter visible after completion. |
| Pass criteria | `ws.AutoFilterMode = False` after macro. All data rows visible. |

---

## TC-VBA-13 · File naming matches Python v3.0 output

| Field | Detail |
|-------|--------|
| Setup | Run `SplitExcelByManager.bas` and `split_excel_by_manager.py` on the same dataset. |
| Input | Both tools with default settings. |
| Expected | File names produced by both tools are identical (e.g. both produce `Alice_report.xlsx`). |
| Pass criteria | File name sets match exactly between `manager_reports/` folders from both runs. |

---

## Notes

- All tests are manual because `SplitExcelByManager.bas` is UI-driven (MsgBox output).
- For automated coverage of the same functional boundaries, use `test_split_excel_by_manager.py` (pytest).
