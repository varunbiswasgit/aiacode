# WordNormalizeTable.bas — Testing Readme

This document describes how testing is conducted for `scripts/WordNormalizeTable.bas`, a Word macro module containing the `NormalizeTables_Light` subroutine.

## Automated Tests

No automated VBA test harness exists for this script. The subroutine is UI-driven (MsgBox output) and operates on the Word object model. All testing is manual.

## Environment Setup

1. Open a Word document containing one or more tables.
2. Press **Alt + F11**, import `scripts/WordNormalizeTable.bas`.
3. Run `NormalizeTables_Light` from **Developer → Macros**.

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
| Pass criteria | Macro exits cleanly with correct zero count. |
