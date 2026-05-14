# ExcelFormatting.bas

A combined, hardened Excel VBA macro module for generic workbook formatting, data cleanup, deduplication, and table conversion. Contains no personally identifiable information, file system paths, credentials, or organisation-specific strings.

## Options at a Glance

| Option | Scope | Pipeline steps |
|--------|-------|----------------|
| **1 – Simple formatting** | Active sheet | Clean text → delete blank rows → cap columns at 55 + autofit columns → remove duplicate rows → autofit row heights |
| **2 – Advanced formatting** | Active sheet | Option 1 steps + text-to-columns (auto type, no delimiters) → delete blank columns → convert to table |
| **3 – Crop and format** | Active sheet | Crop data range by keyword anchor or cell selection, then run all Option 2 steps |

## Pipeline Details

### Shared across all options

1. **Clean text** — strips non-breaking spaces (`Chr(160)`) and tabs (`Chr(9)`), trims whitespace from every string cell.
2. **Delete blank rows** — iterates bottom-up; removes any row where `CountA = 0`.
3. **Cap columns + autofit** — sets `WrapText = True`, `VerticalAlignment = xlVAlignCenter`, autofits each column, then re-caps any column exceeding `MAX_COL_WIDTH` (55).
4. **Remove duplicate rows** — in-memory dictionary dedup, case-insensitive, preserves row order, keeps first occurrence.
5. **Autofit row heights** — recalculates row heights after dedup on the final row count.

### Option 2 and 3 additionally

6. **Text-to-columns** (before step 3 above) — re-evaluates each column’s data type with no delimiter specified, causing Excel to auto-convert text numbers, dates, and booleans.
7. **Delete blank columns** — removes any column where `CountA = 0` across the full used range.
8. **Convert to table** — wraps the used range in an Excel ListObject using style `TableStyleMedium2`. Skipped if the range is already inside a table.

### Option 3 additionally

9. **Crop** (runs first, before all other steps) — two crop modes:
   - **Keyword anchor**: user types the exact text of the first header cell (case-insensitive whole-cell match). The macro locates the anchor, scans the true table boundary, fills blank header cells with synthetic names (`Column1`, `Column2`, …), clears data outside the boundary, and deletes leading rows and columns.
   - **Cell selection**: user clicks the top-left header cell when prompted via an `InputBox(Type:=8)` cell-picker. The same crop logic runs from that cell as the anchor.

## Requirements

- Microsoft Excel (any desktop version supporting VBA)
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

## Constants

| Constant | Default | Purpose |
|----------|---------|--------|
| `MAX_COL_WIDTH` | `55` | Maximum column width (characters) applied after autofit |
| `DEFAULT_TABLE_STYLE` | `"TableStyleMedium2"` | Table style applied by Option 2 and 3 |

## TRunOptions Type

| Field | Type | Purpose |
|-------|------|---------|
| `OptionLevel` | `Long` | Carries the user’s menu choice (1, 2, or 3) to gate pipeline steps |
| `UseKeywordMode` | `Boolean` | `True` = keyword anchor; `False` = cell selection (Option 3 only) |
| `MarkerText` | `String` | Keyword text supplied by the user (keyword mode only) |

## Safety Notes

- `Application.ScreenUpdating`, `Calculation`, and `EnableEvents` are always restored on error via `BeginAppState` / `EndAppState`.
- All `InputBox` cancel paths use type-checked wrapper functions (`TryGetLongInput`, `TryGetMarkerKeyword`).
- Cell selection uses `On Error Resume Next` around the `InputBox(Type:=8)` call; a `Nothing` result raises a clean error with a descriptive message.
- No personally identifiable data, file paths, or credentials are embedded.

## License

See [LICENSE](../LICENSE) in the repository root.
