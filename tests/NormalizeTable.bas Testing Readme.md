# NormalizeTable.bas — Testing Readme

This document describes how testing is conducted for `scripts/NormalizeTable.bas`, a Word macro module containing two table-normalization subroutines.

## Automated Tests

No automated VBA test harness exists for this script. Both subroutines are UI-driven (InputBox / MsgBox output) and operate on the Word object model. All testing is manual.

## Environment Setup

1. Open a Word document containing one or more tables.
2. Press **Alt + F11**, import `scripts/NormalizeTable.bas`.
3. Run the target subroutine from **Developer → Macros**.

---

## Subroutine: `NormalizeTables_Light`

Light-weight daily-driver. Processes tables in the document body only.

### TC-NT-01 · Table width set to 100%

| Field | Detail |
|-------|--------|
| Setup | Document with a table whose preferred width is set to a fixed pixel/point value |
| Expected | Table `PreferredWidthType = wdPreferPercent`, `PreferredWidth = 100`, `AllowAutoFit = True` |
| Pass criteria | Table stretches to full page width. |

### TC-NT-02 · Row height constraints cleared

| Field | Detail |
|-------|--------|
| Setup | Table with rows locked to an exact height |
| Expected | `HeightRule = wdRowHeightAuto`, `Height = 0`, `AllowBreakAcrossPages = True` on every row |
| Pass criteria | Rows auto-size to content. |

### TC-NT-03 · Cell width constraints cleared

| Field | Detail |
|-------|--------|
| Setup | Table with fixed column widths |
| Expected | `Width = 0`, `PreferredWidthType = wdPreferAuto` on every cell |
| Pass criteria | Columns auto-distribute across the table width. |

### TC-NT-04 · Font standardized to Arial 10

| Field | Detail |
|-------|--------|
| Setup | Table cells with mixed fonts and sizes |
| Expected | Every paragraph in every cell: `Font.Name = "Arial"`, `Font.Size = 10` |
| Pass criteria | Select all text in the table; Font toolbar shows Arial, 10pt. |

### TC-NT-05 · Completion message

| Field | Detail |
|-------|--------|
| Setup | Document with 3 tables |
| Expected | MsgBox: **"NormalizeTables_Light complete. 3 table(s) processed."** |
| Pass criteria | Count in message matches actual table count. |

### TC-NT-06 · Document with no tables

| Field | Detail |
|-------|--------|
| Setup | Document with no tables |
| Expected | MsgBox: **"NormalizeTables_Light complete. 0 table(s) processed."** No error raised. |

---

## Subroutine: `StandardizeTables_TwoPass_AllStories`

Heavier two-pass macro. Processes tables in all document stories — body, headers, footers, and text boxes.

### TC-NT-07 · Tables in headers and footers processed

| Field | Detail |
|-------|--------|
| Setup | Document with a table in the page header |
| Expected | Header table normalized: 100% width, Arial 10, auto row/cell sizing |
| Pass criteria | Open header/footer view; confirm formatting applied. |

### TC-NT-08 · Tables in text boxes processed

| Field | Detail |
|-------|--------|
| Setup | Document with a table inside a floating text box |
| Expected | Text box table normalized as above |
| Pass criteria | Click inside text box; confirm table spans full width with Arial 10. |

### TC-NT-09 · Two-pass correctness — unlock then font

| Field | Detail |
|-------|--------|
| Setup | Table with both fixed dimensions and non-standard fonts |
| Expected | Pass 1 unlocks all dimensions; Pass 2 applies font. Both changes persist. |
| Pass criteria | After macro: no fixed widths/heights remain AND font is Arial 10. |

### TC-NT-10 · Completion message with total count across all stories

| Field | Detail |
|-------|--------|
| Setup | Document with 2 body tables + 1 header table |
| Expected | MsgBox: **"StandardizeTables_TwoPass_AllStories complete. 3 table(s) processed."** |
| Pass criteria | Count reflects all stories, not body only. |

### TC-NT-11 · Document with no tables — no error

| Field | Detail |
|-------|--------|
| Setup | Document with no tables |
| Expected | MsgBox: **"StandardizeTables_TwoPass_AllStories complete. 0 table(s) processed."** |
| Pass criteria | No runtime error. |
