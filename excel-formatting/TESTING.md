# Excel Formatting — Testing Guide

No automated test harness exists because `ExcelFormatting.bas` uses `Application.InputBox` and `MsgBox`, which cannot be driven without UI mocking.

> **Future enhancement:** Introduce a `TRunOptions`-driven headless call path to enable automated testing.

---

## Environment Setup

1. Open a new Excel workbook.
2. Press **Alt + F11**, import `ExcelFormatting.bas`.
3. Run `RunUnifiedDataFormatter_v3` from **Developer → Macros**.

---

## Option 1 — Simple Formatting

### TC-01 · Clean sheet with plain data

| Step | Action |
|------|--------|
| Setup | 3 columns, 10 rows. Column A: values stored as text (`'100`, `'200`). Column B: non-breaking spaces (`Chr(160)`) in two cells. |
| Input | Enter `1` |
| Expected | Non-breaking spaces replaced. Column widths ≤ 55. Text wrapped. Row heights autofitted. Completion message: **"Simple formatting completed."** |
| Pass criteria | No data deleted. Cell values preserved. |

### TC-02 · Duplicate rows removed

| Step | Action |
|------|--------|
| Setup | 5-row dataset where rows 3 and 5 are identical to row 2. |
| Input | Enter `1` |
| Expected | Rows 3 and 5 removed. Sheet has 3 unique data rows (header + 2 unique). |
| Pass criteria | `GetLastUsedRow` = 3 after macro. |

### TC-03 · Columns wider than 55 characters

| Step | Action |
|------|--------|
| Setup | One column containing a 200-character string. |
| Input | Enter `1` |
| Expected | Column autofitted, then capped at 55. Text wrapped. Row height autofitted. |
| Pass criteria | `ws.Columns(1).ColumnWidth <= 55` |

### TC-04 · Completely empty sheet

| Step | Action |
|------|--------|
| Setup | Blank sheet. |
| Input | Enter `1` |
| Expected | Macro exits silently. No error. Sheet unchanged. |

### TC-05 · Cancel at option prompt

| Step | Action |
|------|--------|
| Setup | Any sheet with data. |
| Input | Click **Cancel** at the option prompt. |
| Expected | Macro exits. Sheet unchanged. No error. |

### TC-06 · Text-to-columns NOT triggered for Option 1

| Step | Action |
|------|--------|
| Setup | Column A contains date strings (`"01/01/2024"`) stored as text. |
| Input | Enter `1` |
| Expected | Dates remain as text strings. No type conversion occurs. |
| Pass criteria | `VarType(ws.Cells(2,1).Value) = vbString` after macro. |

---

## Option 2 — Advanced Formatting

### TC-07 · Blank rows and columns deleted

| Step | Action |
|------|--------|
| Setup | 4-column, 8-row dataset. Row 4 entirely blank. Column C entirely blank. |
| Input | Enter `2` |
| Expected | Blank row and blank column deleted. Completion message: **"Advanced formatting completed."** |

### TC-08 · Text-to-columns auto type conversion

| Step | Action |
|------|--------|
| Setup | Column B: date strings. Column C: numbers stored as text. |
| Input | Enter `2` |
| Expected | `TextToColumns` re-evaluates types. Dates become date serials. Numbers become numeric. |
| Pass criteria | `VarType(ws.Cells(2,2).Value) = vbDate`. `IsNumeric(ws.Cells(2,3).Value) = True`. |

### TC-09 · Duplicate rows removed and table applied

| Step | Action |
|------|--------|
| Setup | 6-row dataset (header + 5 data rows); rows 4 and 6 are duplicates of row 2. |
| Input | Enter `2` |
| Expected | Duplicates removed. Range converted to a ListObject with style `TableStyleMedium2`. |
| Pass criteria | `ws.ListObjects.Count = 1`. Unique row count = 4. |

### TC-10 · Tab characters in cells

| Step | Action |
|------|--------|
| Setup | Three cells with embedded tab characters (`Chr(9)`). |
| Input | Enter `2` |
| Expected | Tabs replaced with spaces. Values otherwise preserved. |
| Pass criteria | No cell contains `Chr(9)` after macro. |

### TC-11 · Sheet already inside a table

| Step | Action |
|------|--------|
| Setup | Used range already converted to a ListObject before running. |
| Input | Enter `2` |
| Expected | `ConvertUsedRangeToTable` skips. No error. |
| Pass criteria | `ws.ListObjects.Count` unchanged. |

---

## Option 3 — Crop and Format

### TC-12 · Keyword anchor — standard crop with leading rows and columns

| Step | Action |
|------|--------|
| Setup | Row 1: report title. Rows 2–3: blank. Row 4 headers starting at column C. `C4 = "Material Number"`. Data in C4:F12. |
| Input | Enter `3`. Crop method: `1`. Keyword: `Material Number` |
| Expected | Rows 1–3 deleted. Columns A–B deleted. `ws.Cells(1,1) = "Material Number"`. Completion: **"Crop and format completed."** |
| Pass criteria | Row count = 9 (1 header + 8 data). Column count = 4. |

### TC-13 · Cell selection anchor

| Step | Action |
|------|--------|
| Setup | Same as TC-12. |
| Input | Enter `3`. Crop method: `2`. Select cell C4 when prompted. |
| Expected | Same result as TC-12. Rows 1–3 and columns A–B deleted. |
| Pass criteria | `ws.Cells(1,1).Value = "Material Number"`. |

### TC-14 · Blank header cells filled with synthetic names

| Step | Action |
|------|--------|
| Setup | Headers in columns A, C, E. Columns B and D blank in header row but contain data below. |
| Input | Enter `3`. Crop method `1`. Keyword = A1 value. |
| Expected | B1 = `"Column1"`, D1 = `"Column2"`. |
| Pass criteria | No blank cells in header row within table boundary. |

### TC-15 · Keyword not found — continue

| Step | Action |
|------|--------|
| Input | Enter `3`. Crop method `1`. Keyword = `"ZZNOTEXIST"`. Choose **Yes** at not-found dialog. |
| Expected | No cropping. Full Option 2 pipeline runs on the full sheet. No error. |

### TC-16 · Keyword not found — abort

| Step | Action |
|------|--------|
| Input | Enter `3`. Crop method `1`. Keyword = `"ZZNOTEXIST"`. Choose **No**. |
| Expected | Macro stops. Error dialog shown. Sheet unchanged. |

### TC-17 · Cancel at crop method prompt

| Step | Action |
|------|--------|
| Input | Enter `3`. Click **Cancel** at the crop method prompt. |
| Expected | Macro exits immediately. Sheet unchanged. No error. |

### TC-18 · Cancel at keyword prompt

| Step | Action |
|------|--------|
| Input | Enter `3`. Crop method `1`. Click **Cancel** at the keyword `InputBox`. |
| Expected | Macro exits immediately. Sheet unchanged. No error. |

### TC-19 · Cancel at cell selection prompt

| Step | Action |
|------|--------|
| Input | Enter `3`. Crop method `2`. Click **Cancel** at the cell-picker `InputBox`. |
| Expected | Macro stops. Error dialog shown with message "No anchor cell selected." Sheet unchanged. |

### TC-20 · Keyword match is case-insensitive

| Step | Action |
|------|--------|
| Setup | Header cell A5 = `"Order Type"`. |
| Input | Enter `3`. Crop method `1`. Keyword = `"order type"` (all lowercase). |
| Expected | Anchor found. Rows 1–4 deleted. Full pipeline applied. |
| Pass criteria | `ws.Cells(1,1).Value = "Order Type"` (original casing preserved). |

### TC-21 · Option 3 output matches Option 2 on same clean data

| Step | Action |
|------|--------|
| Setup | Two identical sheets with the table already at A1. |
| Input | Option 2 on Sheet1. Option 3 (crop method `1`, keyword = A1 header) on Sheet2. |
| Expected | Both sheets have identical widths, heights, wrap settings, deduped row count, and table formatting. |
| Pass criteria | Cell-by-cell comparison shows no differences. |

---

## Regression — Application State

### TC-22 · Application state restored after error

| Step | Action |
|------|--------|
| Setup | Sheet where `TextToColumns` will fail (data overlaps a PivotTable). |
| Input | Enter `2`. |
| Expected | Error dialog shown. After dismissal: `ScreenUpdating = True`, `Calculation = xlCalculationAutomatic`, `EnableEvents = True`. |
| Pass criteria | Verify each property in the Immediate window after dismissing the error. |
