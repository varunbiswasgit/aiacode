# WordNormalizeTable.bas

A Word VBA module providing a fast, lightweight table normalisation macro suitable for everyday use.

## Macro

### `NormalizeTables_Light`

Processes all tables in the active document body.

- Sets every table to 100% preferred width with `AllowAutoFit = True`
- Clears row height constraints and enables `AllowBreakAcrossPages`
- Resets cell width locks to auto
- Sets all cell text to **Arial 10 pt**
- Displays a completion message with the count of tables processed

## Requirements

- Microsoft Word (any version supporting VBA)
- Macro execution must be enabled

## Installation

1. Press **Alt + F11** in Word.
2. Right-click your document in Project Explorer → **Import File**.
3. Select `WordNormalizeTable.bas`.
4. Run `NormalizeTables_Light` from **Developer → Macros**.

## License

See [LICENSE](../LICENSE) in the repository root.
