# ExcelFormatting.bas

A combined, hardened Excel VBA macro module for generic workbook formatting and data cleanup. Contains no personally identifiable information, file system paths, credentials, or organisation-specific strings.

## Features

| Option | Scope | What it does |
|--------|-------|--------------|
| 1 – Simple formatting | Active sheet | Cleans text, runs `TextToColumns` (no delimiters) for automatic data-type detection, caps column width at 55, wraps text, auto-fits rows and columns |
| 2 – Advanced formatting | Active sheet | Identical to Option 1; intended as an explicit advanced invocation |
| 3 – Keyword crop and format | Active sheet | User supplies a keyword matching the first header cell; the macro crops leading rows and columns, fills any unlabelled spill columns with synthetic names, splits wide columns on comma / semicolon / pipe, removes blank rows and columns, then applies the same formatting as Options 1 and 2 |

## How Column Normalisation Works

**Options 1 and 2** run `TextToColumns` on every column with no delimiters specified. This instructs Excel to re-evaluate each cell's data type without splitting content — dates stored as text become date values, numbers stored as text become numeric, and genuine text strings are preserved as-is.

**Option 3** runs `TextToColumns` with comma, semicolon, and pipe (`|`) delimiters enabled. This splits any cell whose content was joined by those characters during an earlier export or copy-paste operation. The delimiter set is fixed — no user selection is required.

## Requirements

- Microsoft Excel (any version supporting VBA)
- Macro execution must be enabled
- Do not run on sheets containing PivotTables — `TextToColumns` will raise an error if the data range overlaps a PivotTable. Switch to a non-PivotTable sheet before running.

## Installation

1. Open Excel and press **Alt + F11** to open the VBA editor.
2. In the Project Explorer, right-click your workbook and choose **Import File**.
3. Select `ExcelFormatting.bas` and click **Open**.
4. Close the VBA editor.
5. Run `RunUnifiedDataFormatter_v3` from **Developer → Macros**.

## Entry Point

```vb
Public Sub RunUnifiedDataFormatter_v3()
```

Call this macro to launch the interactive menu.

## Option 3 – Keyword Crop and Format

Option 3 requires a single user input: the exact text of the first header cell of the target table. The match is case-insensitive but must equal the full cell value (no partial matches).

Once the anchor cell is located, the macro:

1. Scans rightward and downward from the anchor to determine the true table boundary across all populated columns, ignoring the wider sheet extent.
2. Fills any blank header cells within that boundary with synthetic names (`Column1`, `Column2`, etc.) to handle columns created by delimiter-split exports.
3. Clears data outside the identified table boundary, then deletes leading rows above and columns to the left of the anchor.
4. Splits cell values on comma, semicolon, and pipe — no user input required for delimiter selection.
5. Removes blank rows and blank columns, then applies autofit, column-width capping at 55, text wrapping, and row autofit — identical to Options 1 and 2.

No additional prompts are shown. Cancel at the keyword input aborts the operation.

## Safety Notes

- `Application.ScreenUpdating`, `Application.Calculation`, and `Application.EnableEvents` are always restored, even on error.
- All `Application.InputBox` cancel handling uses type-checked wrapper functions — no fragile direct `False` comparisons.
- `TextToColumns` is applied only to explicit named ranges — no `Selection`-based calls.
- No personally identifiable data, file paths, credentials, or organisation-specific strings are embedded.

## License

See [LICENSE](LICENSE) in the repository root.
