# ExcelFormatting.bas

A combined, hardened Excel VBA macro module for generic workbook formatting and data cleanup. Contains no personally identifiable information, file system paths, credentials, or organisation-specific strings.

## Features

| Option | Scope | What it does |
|--------|-------|-------------- |
| 1 – Simple formatting | All worksheets | Trims/cleans text, caps column width at 55, wraps text, auto-fits rows and columns |
| 2 – Advanced formatting | Active sheet | Same as above plus `TextToColumns` datatype coercion on every column |
| 3 – Advanced + optional split | Active sheet | Option 2 plus an interactive column-split prompt with delimiter choice |
| 4 – SAP output processing | Active sheet | Marker-based row/column cropping, deduplication, blank row/column removal, optional table conversion |

## Requirements

- Microsoft Excel (any version supporting VBA)
- Macro execution must be enabled

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

SAP mode (option 4) supports two built-in marker strings used to crop leading metadata rows and columns that SAP exports typically include:

- `Selection No.`
- `Date`

You can choose which marker to use at runtime, or skip marker trimming entirely.

## Safety Notes

- `Application.ScreenUpdating`, `Application.Calculation`, and `Application.EnableEvents` are always restored, even on error.
- All `Application.InputBox` cancel handling uses type-checked wrapper functions — no fragile direct `False` comparisons.
- `TextToColumns` is applied only to explicit named ranges — no `Selection`-based calls.
- Header-row awareness is user-controlled for both duplicate removal and table creation.
- No personally identifiable data, file paths, credentials, or organisation-specific strings are embedded.

## License

See [LICENSE](LICENSE) in the repository root.
