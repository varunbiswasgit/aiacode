# ExcelFormatting.bas — Testing Readme

Test coverage for `scripts/ExcelFormatting.bas`.

## Automated Tests

No automated VBA test harness exists because `ExcelFormatting.bas` uses `Application.InputBox` and `MsgBox`, which cannot be driven without UI mocking.  
Future enhancement: introduce a `TRunOptions`-driven headless call path to enable automated testing.

## Manual Test Cases

Environment setup:

1. Open a new Excel workbook.
2. Press **Alt + F11**, import `scripts/ExcelFormatting.bas`.
3. Run `RunUnifiedDataFormatter_v3` from **Developer → Macros**.

---

## Option 1 — Simple Formatting

### TC-01 · Clean sheet with plain data

| Step | Action |
|------|--------|
| Setup | 3 columns, 5 rows. Column A: values stored as text (`'100`, `'200`). Column B: non-breaking spaces (Chr 160) in two cells. |
| Input | Enter `1` |
| Expected | Non-breaking spaces replaced. Text numbers coerced to numeric. Width ≤ 55. Text wrapped. Completion message: **"Simple formatting completed."** |
| Pass criteria | No data deleted. Cell types corrected. |

### TC-02 · Completely empty sheet

| Step | Action |
|------|--------|
| Setup | Blank sheet |
| Input | Enter `1` |
| Expected | Macro exits silently. No error. Sheet unchanged. |

### TC-03 · Columns wider than 55 characters

| Step | Action |
|------|--------|
| Setup | One column containing a 200-character string |
| Input | Enter `1` |
| Expected | Column width capped at 55. Text wrapped. Row auto-fitted. |
| Pass criteria | `ws.Columns(1).ColumnWidth <= 55` |

### TC-04 · Cancel at option prompt

| Step | Action |
|------|--------|
| Setup | Any sheet with data |
| Input | Click **Cancel** |
| Expected | Macro exits. Sheet unchanged. No error. |

---

## Option 2 — Advanced Formatting

### TC-05 · Blank rows and columns present

| Step | Action |
|------|--------|
| Setup | 4-column, 8-row dataset. Row 4 entirely blank. Column C entirely blank. |
| Input | Enter `2` |
| Expected | Blank row and column deleted. Completion message: **"Advanced formatting completed."** |

### TC-06 · Tab characters in cells

| Step | Action |
|------|--------|
| Setup | Three cells with embedded tab characters (Chr 9) |
| Input | Enter `2` |
| Expected | Tabs replaced with spaces. Values otherwise preserved. |
| Pass criteria | No cell contains Chr(9) after macro. |

### TC-07 · Data types coerced correctly

| Step | Action |
|------|--------|
| Setup | Dates as text strings in column B. Numbers as text in column C. |
| Input | Enter `2` |
| Expected | `TextToColumns` re-evaluates types. Dates become date serials. Numbers become numeric. |

---

## Option 3 — Keyword Crop and Format

### TC-08 · Standard crop with leading rows and columns

| Step | Action |
|------|--------|
| Setup | Row 1: report title. Rows 2–3: blank. Row 4 headers at column C. `C4 = "Material Number"`. Data in C4:F12. |
| Input | Enter `3`. Keyword: `Material Number` |
| Expected | Rows 1–3 deleted. Columns A–B deleted. `ws.Cells(1,1) = "Material Number"`. Completion: **"Keyword crop and format completed."** |
| Pass criteria | Row count = 9 (1 header + 8 data). |

### TC-09 · Blank header cells filled with synthetic names

| Step | Action |
|------|--------|
| Setup | Headers in A, C, E. Columns B and D blank in header row but have data below. |
| Input | Enter `3`. Keyword = A1 value. |
| Expected | B1 = `"Column1"`, D1 = `"Column2"` |
| Pass criteria | No blank cells in header row within table boundary. |

### TC-10 · Sparse data — true boundary detection

| Step | Action |
|------|--------|
| Setup | 5-column table. Column A data to row 10. Column E has one value at row 20. Rows 11–19 of A blank. |
| Input | Enter `3`. Keyword = A1 header. |
| Expected | Table boundary extends to row 20 (column E drives last row). |
| Pass criteria | `GetLastUsedRow` ≥ 20 before blank-column cleanup. |

### TC-11 · Keyword not found — continue

| Step | Action |
|------|--------|
| Input | Enter `3`. Keyword = `"ZZNOTEXIST"`. Choose **Yes** at not-found dialog. |
| Expected | No cropping. Formatting runs on full sheet. No error. |

### TC-12 · Keyword not found — abort

| Step | Action |
|------|--------|
| Input | Enter `3`. Keyword = `"ZZNOTEXIST"`. Choose **No**. |
| Expected | Macro stops. Error dialog shown. Sheet unchanged. |

### TC-13 · Cancel at keyword prompt

| Step | Action |
|------|--------|
| Input | Enter `3`. Click **Cancel** at keyword InputBox. |
| Expected | Macro exits immediately. Sheet unchanged. |

### TC-14 · Keyword match is case-insensitive

| Step | Action |
|------|--------|
| Setup | Header cell A5 = `"Order Type"` |
| Input | Enter `3`. Keyword = `"order type"` (lowercase) |
| Expected | Anchor found. Rows 1–4 deleted. Formatting applied. |
| Pass criteria | `ws.Cells(1,1).Value = "Order Type"` (original casing preserved) |

### TC-15 · Formatting output identical to Option 2 on same data

| Step | Action |
|------|--------|
| Setup | Two identical sheets with table at A1. |
| Input | Option 2 on Sheet1. Option 3 (keyword = A1 header) on Sheet2. |
| Expected | Both sheets have identical widths, heights, wrap settings, and values. |
| Pass criteria | Cell-by-cell comparison shows no differences. |

---

## Regression — Application State

### TC-16 · Application state restored after error

| Step | Action |
|------|--------|
| Setup | Sheet where `TextToColumns` will fail (data overlaps a PivotTable) |
| Input | Any option |
| Expected | Error dialog shown. After dismissal: `ScreenUpdating = True`, `Calculation = xlCalculationAutomatic`, `EnableEvents = True` |
| Pass criteria | Check each property in the Immediate window. |
