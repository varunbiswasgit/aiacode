# SplitExcelByManager.bas

An Excel VBA macro that splits the active sheet into one `.xlsx` workbook per unique manager value, saving each file to a `manager_reports` subfolder beside the open workbook.

## Features

- Locates the manager column by header name — no hardcoded column index
- Sanitizes filenames: replaces invalid characters, guards against Windows reserved device names (`CON`, `PRN`, `NUL`, `COM1`–`COM9`, `LPT1`–`LPT9`), caps length at 200 characters
- Auto-fits column widths, clamped between 8 and 50 characters
- Skips blank manager entries
- Per-manager error handling — one failure does not abort the entire run
- Displays a completion summary: reports saved vs. total unique managers

## Configuration

Edit the two constants at the top of the macro before running:

```vb
Const MANAGER_COL_NAME  As String = "Manager"          ' Column header
Const OUTPUT_SUBFOLDER  As String = "manager_reports"  ' Output subfolder
```

## Requirements

- Microsoft Excel (any version supporting VBA)
- Macro execution must be enabled
- The workbook must be saved before running (output path is derived from `ThisWorkbook.Path`)

## Installation

1. Press **Alt + F11** in Excel.
2. Right-click your workbook in Project Explorer → **Import File**.
3. Select `SplitExcelByManager.bas`.
4. Run `SplitExcelByManager` from **Developer → Macros**.

## License

See [LICENSE](../LICENSE) in the repository root.
