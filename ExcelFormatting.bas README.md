# ExcelFormatting.bas

A combined, hardened Excel VBA macro module for generic workbook formatting and data cleanup. Contains no personally identifiable information, file system paths, credentials, or organisation-specific strings.

## Features

| Option | Scope | What it does |
|--------|-------|------|
| 1 – Simple formatting | Active sheet | Trims and cleans text, runs `TextToColumns` on all columns for automatic data-type detection, caps column width at 55, wraps text, auto-fits rows and columns |
| 2 – Advanced formatting | Active sheet | Identical to Option 1; intended for explicit advanced use |
| 3 – SAP output processing | Active sheet | Marker-based row/column cropping, deduplication, blank row/column removal, optional Excel table conversion |

## How Column Normalisation Works

All options run `TextToColumns` on every column with **no delimiters specified**. This instructs Excel to re-evaluate each cell's data type without splitting content — dates stored as text become date values, numbers stored as text become numeric, and genuine text strings are preserved as-is. No user input is required.

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

## SAP Mode Details

SAP mode (Option 3) supports two built-in marker strings used to crop leading metadata rows and columns that SAP exports typically include:

- `Selection No.`
- `Date`

You can choose which marker to use at runtime, or skip marker trimming entirely. Additional prompts control deduplication, blank column handling (with or without header-row awareness), and optional conversion to an Excel table.

## Safety Notes

- `Application.ScreenUpdating`, `Application.Calculation`, and `Application.EnableEvents` are always restored, even on error.
- All `Application.InputBox` cancel handling uses type-checked wrapper functions — no fragile direct `False` comparisons.
- `TextToColumns` is applied only to explicit named ranges — no `Selection`-based calls.
- Header-row awareness is user-controlled for both duplicate removal and table creation.
- No personally identifiable data, file paths, credentials, or organisation-specific strings are embedded.

## License

See [LICENSE](LICENSE) in the repository root.
