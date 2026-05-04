# Split Excel by Manager — VBA Macro

This VBA macro splits the active worksheet into separate workbooks, one per unique manager. It reads manager names from **column A** (row 1 = header, data from row 2), filters rows for each manager, and saves each output workbook to the **same folder as the source file**.

Each output workbook has auto-fitted column widths for readability.

---

## Requirements

- Microsoft Excel (any version supporting `.xlsx` / `xlOpenXMLWorkbook`)
- The source workbook **must be saved to disk** before running the macro

---

## Usage

1. Open the workbook containing the data.
2. Press `Alt + F11` to open the VBA Editor.
3. Go to **File → Import File** and select `SplitExcelByManager.bas`.
4. Press `F5` or click **Run** to execute `SplitExcelByManager`.
5. Output files are saved in the same folder as the source workbook, named `<ManagerName>_Report.xlsx`.

---

## Input Requirements

| Requirement | Detail |
|---|---|
| Manager column | Column A (row 1 = header, data starts row 2) |
| File must be saved | Macro uses `ThisWorkbook.Path` to determine output location |
| No pre-existing AutoFilter conflict | Macro clears any active AutoFilter before running |

---

## Output

```
source_folder/
├── YourWorkbook.xlsx          ← source file
├── JohnSmith_Report.xlsx
├── SarahJohnson_Report.xlsx
└── ...
```

---

## Key Fixes Applied (v2.0)

| Issue | Fix |
|---|---|
| Output saved to Excel default directory | Now uses `ThisWorkbook.Path` |
| No guard if workbook was unsaved | Added check; shows message and exits |
| AutoFilter conflict if already active | Cleared before applying new filter |
| Manager names with `/\:*?<>|` crashed SaveAs | Sanitization loop replaces invalid chars with `-` |
| `On Error Resume Next` scope too broad | Narrowed to Collection deduplication only |
| Blank/null manager names not filtered | Skipped before processing |

---

## Troubleshooting

**"Please save the workbook first"**
Save the file to a folder on disk (`Ctrl + S`) before running the macro.

**Output file not found**
Check the folder where the source workbook is saved — all reports are written there.

**Special characters in manager names**
Characters invalid in Windows file names (`/ \ : * ? < > |`) are automatically replaced with `-`.

---

## Version History

### v2.0
- Fixed output path to use `ThisWorkbook.Path`
- Added unsaved workbook guard
- Added AutoFilter state reset before execution
- Added manager name sanitization for file-system safety
- Added null/blank manager name skip
- Added completion message box with file count and path

### v1.0
- Basic split-by-manager functionality
- AutoFilter + `SpecialCells` copy approach
- Column auto-fit on output workbooks
