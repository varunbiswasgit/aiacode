# Split Excel by Manager — VBA Macro (v3.0)

This VBA macro splits the active worksheet into one workbook per unique manager. Reports are saved to a `manager_reports` subfolder alongside the source workbook. **Feature-synchronized with `split_excel_by_manager.py` v3.0.**

---

## Requirements

- Microsoft Excel (any version supporting `.xlsx` / `xlOpenXMLWorkbook`)
- Source workbook **must be saved to disk** before running the macro

---

## Configuration

Two constants at the top of the macro control its behaviour — no code changes needed anywhere else:

```vba
Const MANAGER_COL_NAME  As String = "Manager"          ' Header of the manager column
Const OUTPUT_SUBFOLDER  As String = "manager_reports"  ' Output folder under workbook path
```

Change `MANAGER_COL_NAME` to match your actual column header (e.g. `"Supervisor"`, `"Team Lead"`).

---

## Usage

1. Open the workbook containing the data.
2. Press `Alt + F11` to open the VBA Editor.
3. Go to **File → Import File** and select `SplitExcelByManager.bas`.
4. Press `F5` or click **Run** to execute `SplitExcelByManager`.
5. A confirmation box shows the count and full path of saved reports.

---

## Input Requirements

| Requirement | Detail |
|---|---|
| Manager column | Located by header name (configurable via `MANAGER_COL_NAME`) |
| File must be saved | Macro uses `ThisWorkbook.Path` for output location |
| Header row | Must be in row 1; data from row 2 downward |

---

## Output

```
source_folder/
├── YourWorkbook.xlsx
└── manager_reports/
    ├── John_Smith_report.xlsx
    ├── Sarah_Johnson_report.xlsx
    └── ...
```

- File name pattern: `{sanitized_name}_report.xlsx` (lowercase suffix, matches Python)
- Invalid characters (`/ \ : * ? < > |`) replaced with `_`
- Windows reserved names (CON, PRN, AUX, NUL, COM1–9, LPT1–9) replaced with `Unknown_Manager`
- Names truncated to 200 characters
- Column widths clamped between 8 and 50 characters

---

## Feature Parity with Python v3.0

| Feature | Python | VBA |
|---|---|---|
| Configurable manager column | CLI arg 2 | `MANAGER_COL_NAME` constant |
| Configurable output folder | CLI arg 3 | `OUTPUT_SUBFOLDER` constant |
| Column located by header name | Yes | Yes |
| Skip blank/null managers | Yes | Yes |
| Sanitize invalid filename chars | `_` replacement | `_` replacement |
| Windows reserved name check | CON–NUL, COM1–9, LPT1–9 | CON–NUL, COM1–9, LPT1–9 |
| Name length cap | 200 chars | 200 chars |
| Column width cap | 8–50 chars | 8–50 chars |
| Per-manager error handling | try/except, continues | On Error GoTo, continues |
| Output file suffix | `_report.xlsx` | `_report.xlsx` |
| Missing column error | Yes | Yes (MsgBox) |
| Empty data guard | Yes | Yes (MsgBox) |
| Completion summary | Print to console | MsgBox with count + path |

---

## Troubleshooting

**"Please save the workbook first"** — Save the file (`Ctrl + S`) before running.

**"Column not found"** — Update `MANAGER_COL_NAME` to match the exact header in your data (case-sensitive).

**Reports not appearing** — Check the `manager_reports` subfolder in the same directory as your workbook.

**Special characters in names** — Characters `/ \ : * ? < > |` are automatically replaced with `_`.

---

## Version History

### v3.0
- Manager column located by configurable header name (not hardcoded column A)
- Configurable output subfolder via `OUTPUT_SUBFOLDER` constant
- Full Windows reserved name list added (COM1–9, LPT1–9)
- 200-character name length cap
- Explicit column width cap (min 8, max 50) — matches Python
- Per-manager `On Error GoTo` error handling (failed manager skipped, loop continues)
- Sanitization changed from `-` to `_` to match Python
- Output suffix standardised to lowercase `_report.xlsx`
- Column-missing and empty-data guards with `MsgBox`

### v2.0
- Fixed output path to use `ThisWorkbook.Path`
- Added unsaved workbook guard
- Added AutoFilter state reset
- Added manager name sanitization
- Added blank manager name skip
- Added completion MsgBox

### v1.0
- Basic split-by-manager functionality
- AutoFilter + SpecialCells copy approach
- Column auto-fit on output workbooks
