# NormalizeTable.bas

A Word VBA module providing two table normalisation macros: a fast daily-driver and a thorough two-pass version for documents containing images or nested content.

## Macros

### `NormalizeTables_Light`

Fast, safe, suitable for most documents.

- Sets every table to 100% preferred width, `AllowAutoFit = True`
- Clears row height constraints, enables `AllowBreakAcrossPages`
- Resets cell width locks to auto
- Sets all cell text to **Arial 10 pt**

### `StandardizeTables_TwoPass_AllStories`

Heavier version for documents with images, text boxes, or nested tables.

- Iterates all story ranges (body, headers, footers, text boxes, footnotes)
- Two-pass approach: unlock dimensions first, then apply font — prevents dimension locks from blocking font changes
- Same output as `NormalizeTables_Light` but more thorough

## When to Use Which

| Scenario | Recommended macro |
|----------|-------------------|
| Standard tables, no images | `NormalizeTables_Light` |
| Tables with embedded images | `StandardizeTables_TwoPass_AllStories` |
| Headers / footers contain tables | `StandardizeTables_TwoPass_AllStories` |
| Text boxes contain tables | `StandardizeTables_TwoPass_AllStories` |

## Requirements

- Microsoft Word (any version supporting VBA)
- Macro execution must be enabled

## Installation

1. Press **Alt + F11** in Word.
2. Right-click your document in Project Explorer → **Import File**.
3. Select `NormalizeTable.bas`.
4. Run `NormalizeTables_Light` or `StandardizeTables_TwoPass_AllStories` from **View → Macros**.

## License

See [LICENSE](../LICENSE) in the repository root.
