# WordNormalizeTable.bas

A Word VBA module that normalises all tables in the active document to match the standard manual **Table Properties** settings in Microsoft Word.

## Macros

### `NormalizeTables_Light`

Entry point. Iterates every table in the active document body and calls `FormatTable` on each one. Displays a completion message with the table count when done.

### `FormatTable(tbl As Table)` *(Private)*

Applies all normalisation to a single table. Logic mirrors exactly what Word does when you right-click a table and set properties manually:

| Table Properties Tab | Setting | Code |
|----------------------|---------|------|
| **Table** | AutoFit first, then width = 100% | `AutoFitBehavior wdAutoFitContent` → `PreferredWidthType = wdPreferredWidthPercent`, `PreferredWidth = 100` |
| **Table** | Alignment = Left | `Rows.Alignment = wdAlignRowLeft` |
| **Row** | Uncheck *Specify Height* | `HeightRule = wdRowHeightAuto` |
| **Cell** | Uncheck *Preferred Width* | `PreferredWidthType = wdPreferredWidthAuto` |
| **Cell** | Vertical alignment = Centre | `VerticalAlignment = wdCellAlignVerticalCenter` |

**Design decisions:**
- `AutoFitBehavior` is called before setting 100% width — this mirrors the silent normalisation Word performs when Table Properties is opened.
- The `Columns` collection is skipped entirely. Iterating it on tables with merged or irregular cells causes a *Value out of range* runtime error.
- `On Error Resume Next` guards the `AutoFitBehavior` call and the cell loop against edge cases in merged cells.
- Font and border settings are intentionally excluded — this macro only resets structural layout properties.

## Requirements

- Microsoft Word (any version supporting VBA)
- Macro execution enabled in Trust Center

## Installation

1. Press **Alt + F11** in Word to open the VBA IDE.
2. Right-click your document in Project Explorer → **Import File**.
3. Select `WordNormalizeTable.bas` from this folder.
4. Run `NormalizeTables_Light` from **Developer → Macros**.

## Version History

| Version | Summary |
|---------|---------|
| v1 | Initial release — `NormalizeTables_Light` with table, row, cell, and font normalisation. |
| v2 | Rewrote as `FormatTable` private sub. Removed font changes and `AllowBreakAcrossPages`. Replaced numeric literals with named Word constants. Skipped `Columns` collection to prevent runtime errors on merged cells. Added `AutoFitBehavior` pre-pass. |

## License

See [LICENSE](../LICENSE) in the repository root.
