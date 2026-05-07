# ExcelFormatting.bas — Integration and Regression Test Cases

All tests target `RunUnifiedDataFormatter_v3` in `ExcelFormatting.bas`.  
Run each test on a **fresh copy** of the described worksheet. Pass/fail criteria are explicit.

---

## Environment Setup

1. Open a new Excel workbook.
2. Press **Alt + F11**, import `ExcelFormatting.bas`.
3. Close the VBA editor.
4. Run `RunUnifiedDataFormatter_v3` from **Developer → Macros**.

---

## Option 1 — Simple Formatting

### TC-01 · Clean sheet with plain data

| Step | Action |
|------|--------|
| Setup | Sheet with 3 columns, 5 data rows. Column A has values stored as text (`'100`, `'200`). Column B has non-breaking spaces (Chr 160) in two cells. |
| Input | Enter `1` at the option prompt. |
| Expected | Non-breaking spaces replaced with regular spaces. Text numbers coerced to numeric. Column widths ≤ 55. Text wrapped. Rows auto-fitted. Completion message reads **"Simple formatting completed."** |
| Pass criteria | No data deleted. Cell types corrected. No prompt after the option selection. |

### TC-02 · Completely empty sheet

| Step | Action |
|------|--------|
| Setup | Activate a blank sheet with no data. |
| Input | Enter `1`. |
| Expected | Macro exits silently. No error. No data added or removed. |
| Pass criteria | Sheet remains empty. No error dialog. |

### TC-03 · Columns wider than 55 characters

| Step | Action |
|------|--------|
| Setup | One column containing a 200-character string. |
| Input | Enter `1`. |
| Expected | Column width capped at 55. Text wrapped. Row height auto-fitted. |
| Pass criteria | `ws.Columns(1).ColumnWidth <= 55`. |

### TC-04 · Cancel at option prompt

| Step | Action |
|------|--------|
| Setup | Any sheet with data. |
| Input | Click **Cancel** at the option number prompt. |
| Expected | Macro exits immediately. Sheet unchanged. No error. |
| Pass criteria | Sheet data identical to before macro was run. |

---

## Option 2 — Advanced Formatting

### TC-05 · Blank rows and columns present

| Step | Action |
|------|--------|
| Setup | 4-column, 8-row dataset. Row 4 is entirely blank. Column C is entirely blank. |
| Input | Enter `2`. |
| Expected | Blank row 4 deleted. Blank column C deleted. Remaining data formatted. Completion message reads **"Advanced formatting completed."** |
| Pass criteria | `GetLastUsedRow` decreases by 1. `GetLastUsedColumn` decreases by 1. |

### TC-06 · Sheet with tab characters in cells

| Step | Action |
|------|--------|
| Setup | Three cells containing embedded tab characters (Chr 9). |
| Input | Enter `2`. |
| Expected | Tab characters replaced with a space. Cell values otherwise preserved. |
| Pass criteria | No cell contains Chr(9) after macro completes. |

### TC-07 · Data types coerced correctly

| Step | Action |
|------|--------|
| Setup | Dates stored as text strings in column B (e.g. `"01/15/2024"`). Numbers stored as text in column C. |
| Input | Enter `2`. |
| Expected | `TextToColumns` (no delimiters) re-evaluates data types. Dates become date serials. Numbers become numeric. |
| Pass criteria | `IsNumeric(ws.Cells(2,3).Value)` = True. `IsDate` or date serial check passes for column B cells. |

---

## Option 3 — Keyword Crop and Format

### TC-08 · Standard crop with leading rows and columns

| Step | Action |
|------|--------|
| Setup | Row 1: report title. Rows 2–3: blank. Row 4: headers starting at column C. `C4 = "Material Number"`. Data in C4:F12. Columns A–B empty above row 4. |
| Input | Enter `3`. Keyword: `Material Number`. |
| Expected | Rows 1–3 deleted. Columns A–B deleted. Sheet now starts at A1 = `"Material Number"`. Data rows follow. Standard formatting applied. Completion message reads **"Keyword crop and format completed."** |
| Pass criteria | `ws.Cells(1,1).Value = "Material Number"`. Row count = 9 (1 header + 8 data). |

### TC-09 · Blank header cells filled with synthetic names

| Step | Action |
|------|--------|
| Setup | Header row has values in columns A, C, E. Columns B and D are blank in the header row but contain data below. |
| Input | Enter `3`. Keyword = value of A1. |
| Expected | B1 = `"Column1"`, D1 = `"Column2"`. All data preserved. |
| Pass criteria | No blank cells remain in the header row within the table boundary. |

### TC-10 · Data sparse across columns — true boundary detection

| Step | Action |
|------|--------|
| Setup | 5-column table. Column A has data to row 10. Column E has one value at row 20. Rows 11–19 of column A are blank. |
| Input | Enter `3`. Keyword = A1 header. |
| Expected | Table boundary extends to row 20 (column E drives the last row). No data in row 20, column E is deleted after blank-column cleanup only if that column is entirely blank. |
| Pass criteria | `GetLastUsedRow` ≥ 20 before blank-column cleanup. |

### TC-11 · Keyword not found — user chooses to continue

| Step | Action |
|------|--------|
| Setup | Any sheet. |
| Input | Enter `3`. Keyword = `"ZZNOTEXIST"`. At the not-found dialog choose **Yes** (continue without crop). |
| Expected | No cropping. Standard formatting pipeline runs on full sheet. |
| Pass criteria | Sheet data unchanged in structure. Formatting applied. No error. |

### TC-12 · Keyword not found — user chooses to abort

| Step | Action |
|------|--------|
| Setup | Any sheet. |
| Input | Enter `3`. Keyword = `"ZZNOTEXIST"`. At the not-found dialog choose **No**. |
| Expected | Macro stops. Error dialog shown. Sheet unchanged. |
| Pass criteria | Error dialog appears. Sheet data identical to pre-run state. |

### TC-13 · Cancel at keyword prompt

| Step | Action |
|------|--------|
| Setup | Any sheet with data. |
| Input | Enter `3`. Click **Cancel** at the keyword InputBox. |
| Expected | Macro exits immediately. Sheet unchanged. No error. |
| Pass criteria | Sheet data identical to before macro was run. |

### TC-14 · Keyword match is case-insensitive

| Step | Action |
|------|--------|
| Setup | Header cell A5 = `"Order Type"`. |
| Input | Enter `3`. Keyword = `"order type"` (lowercase). |
| Expected | Anchor found. Rows 1–4 deleted. Formatting applied. |
| Pass criteria | `ws.Cells(1,1).Value` = `"Order Type"` (original casing preserved). |

### TC-15 · Formatting output identical to Option 2 on same data

| Step | Action |
|------|--------|
| Setup | Two identical sheets: Sheet1 and Sheet2 with table starting at A1 (no leading rows). |
| Input | Run Option 2 on Sheet1. Run Option 3 (keyword = A1 header) on Sheet2. |
| Expected | Both sheets have identical column widths, row heights, wrap settings, and data values after macro. |
| Pass criteria | Cell-by-cell comparison of Sheet1 and Sheet2 shows no differences. |

---

## Regression — Application State

### TC-16 · Application state always restored after error

| Step | Action |
|------|--------|
| Setup | Sheet where `TextToColumns` will fail (e.g. data range overlaps a PivotTable). |
| Input | Enter any option (1, 2, or 3). |
| Expected | Error dialog shown. After dismissal: `Application.ScreenUpdating = True`, `Application.Calculation = xlCalculationAutomatic`, `Application.EnableEvents = True`. |
| Pass criteria | Check each property in the Immediate window after the error. All three = True / xlCalculationAutomatic. |

---

## Notes

- All tests are manual because `ExcelFormatting.bas` uses `Application.InputBox` and `MsgBox`, which cannot be driven by automated VBA test harnesses without mocking.
- Tests TC-01 through TC-16 cover all three option paths, shared pipeline functions (`CleanTextInRange`, `NormalizeColumns`, `DeleteBlankRows`, `DeleteBlankColumns`, `ApplyStandardFormatting`), and application-state recovery.
- Future enhancement: introduce a `TRunOptions`-driven headless call path (no UI prompts) to enable automated testing via a separate test runner module.
