# NormalizeTable.bas

## Overview

`NormalizeTable.bas` is a lightweight, fast Word VBA macro that standardizes all tables in the active document to a consistent layout, font, and width. It is designed as the default daily-driver for most documents — significantly faster than the full two-pass `StandardizeTables_TwoPass_AllStories` macro.

For documents with stubborn tables containing embedded images, nested tables, or text boxes, use the heavier `StandardizeTables_TwoPass_AllStories` (Version 3) macro as a repair tool.

---

## Macro: `NormalizeTables_Light`

### What It Does

- Loops through all story ranges (main body, headers, footers, footnotes, etc.) in the active document.
- For each table found, applies the following normalization:
  - Removes text wrap around the table.
  - Left-aligns the table with zero left indent.
  - Clears all manual row heights (sets to Auto).
  - Clears column preferred widths (sets to Auto).
  - Applies `AutoFitBehavior wdAutoFitWindow` to fit the table to the page width.
  - Sets the table preferred width to **100%** and locks it (`AllowAutoFit = False`).
  - Sets all table text to **Arial 10pt**, left-aligned.
- Suppresses screen updates and alerts during the run for performance.
- Reports total elapsed time on completion.

### What It Does NOT Do

- Does **not** process tables inside floating shapes or text boxes.
- Does **not** clear cell-level preferred widths.
- Does **not** handle nested tables separately.
- Does **not** resize or reposition images inside table cells.

These are intentional omissions to keep the macro fast. If cell/column widths remain locked after running (visible as checked **Preferred width** in Table Properties > Column or Cell), the cause is likely embedded images forcing a minimum width. In that case, use the full repair macro.

---

## Configuration

Two constants at the top of the module control formatting output:

| Constant | Default | Description |
|---|---|---|
| `TARGET_FONT_NAME` | `Arial` | Font applied to all table text |
| `TARGET_FONT_SIZE` | `10` | Font size applied to all table text |

---

## Installation

1. Open the Word document or your Personal Macro Workbook (`Normal.dotm`).
2. Press `Alt + F11` to open the VBA Editor.
3. In the Project Explorer, right-click the target project > **Import File**.
4. Select `NormalizeTable.bas`.
5. Close the VBA Editor.
6. Run `NormalizeTables_Light` via **Developer > Macros** or assign to a Quick Access Toolbar button.

---

## Usage

1. Open the Word document to normalize.
2. Run the macro `NormalizeTables_Light`.
3. A completion dialog shows elapsed time.
4. Check tables visually. If column or cell widths are still fixed (due to images), run `StandardizeTables_TwoPass_AllStories` as a follow-up repair.

---

## When to Use Each Macro

| Scenario | Macro to Use |
|---|---|
| Standard document with normal tables | `NormalizeTables_Light` (this file) |
| Document with images inside table cells | `StandardizeTables_TwoPass_AllStories` |
| Document with nested tables | `StandardizeTables_TwoPass_AllStories` |
| Document with tables inside text boxes | `StandardizeTables_TwoPass_AllStories` |
| Performance-sensitive / large documents | `NormalizeTables_Light` (this file) |

---

## Known Limitations

- Embedded images in table cells can prevent Word from fully releasing column and cell preferred widths, regardless of macro. This is a Word layout engine constraint, not a macro bug.
- Tables inside floating shapes are not processed. This is by design for performance.
- Nested tables are processed at the story-range level only (Word's `sr.Tables` collection includes nested tables via story iteration).

---

## Version History

| Version | Date | Notes |
|---|---|---|
| 1.0 | 2026-05-05 | Initial release — lightweight single-pass normalization |

---

## Related Files

- `WordResizeBorderImagesCleanlines.bas` — image resizing and border cleanup for Word documents.
- `ExcelFormatting.bas` — Excel table and cell formatting standardization.
