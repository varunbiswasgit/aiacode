# ExcelFormatting.bas

A combined, hardened Excel VBA macro module for generic workbook formatting and data cleanup. Contains no personally identifiable information, file system paths, credentials, or organisation-specific strings.

## Features

| Option | Scope | What it does |
|--------|-------|---------------|
| 1 – Simple formatting | Active sheet | Cleans text, runs `TextToColumns` (no delimiters) for automatic data-type detection, caps column width at 55, wraps text, auto-fits rows and columns |
| 2 – Advanced formatting | Active sheet | Identical to Option 1; intended as an explicit advanced invocation |
| 3 – Keyword crop and format | Active sheet | User supplies a keyword matching the first header cell; the macro crops leading rows and columns, fills any unlabelled spill columns with synthetic names, removes blank rows and columns, then applies the same formatting as Options 1 and 2 |

## How Column Normalisation Works

**Options 1 and 2** run `TextToColumns` on every column with no delimiters specified. This instructs Excel to re-evaluate each cell's data type without splitting content.

**Option 3** locates the anchor cell by exact whole-cell match (case-insensitive), then scans rightward and downward to determine the true table boundary.

## Requirements

- Microsoft Excel (any version supporting VBA)
- Macro execution must be enabled
- Do not run on sheets containing PivotTables

## Installation

1. Open Excel and press **Alt + F11**.
2. Right-click your workbook in Project Explorer and choose **Import File**.
3. Select `ExcelFormatting.bas` and click **Open**.
4. Run `RunUnifiedDataFormatter_v3` from **Developer → Macros**.

## Entry Point

```vb
Public Sub RunUnifiedDataFormatter_v3()
```

## Option 3 – Keyword Crop and Format

Option 3 requires a single input: the exact text of the first header cell. The match is case-insensitive, whole-cell.

Once the anchor is located, the macro:

1. Scans rightward and downward to determine the true table boundary across all populated columns.
2. Fills blank header cells within the boundary with synthetic names (`Column1`, `Column2`, …).
3. Clears data outside the table boundary, then deletes leading rows and columns.
4. Removes blank rows and blank columns.
5. Applies autofit, column-width cap at 55, text wrapping, and row autofit.

Cancel at the keyword prompt aborts the operation cleanly.

## Safety Notes

- `Application.ScreenUpdating`, `Calculation`, and `EnableEvents` are always restored on error.
- All `InputBox` cancel handling uses type-checked wrapper functions.
- No personally identifiable data, file paths, or credentials are embedded.

## License

See [LICENSE](../LICENSE) in the repository root.
