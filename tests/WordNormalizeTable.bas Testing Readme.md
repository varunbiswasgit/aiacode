# WordNormalizeTable.bas — Testing Readme

This document describes how testing is conducted for `scripts/WordNormalizeTable.bas`. The module exposes one public entry point (`NormalizeTables_Light`) and one private helper (`FormatTable`). All testing is manual — no automated VBA test harness exists.

## Environment Setup

1. Open a Word document containing one or more tables.
2. Press **Alt + F11**, import `scripts/WordNormalizeTable.bas`.
3. Run `NormalizeTables_Light` from **Developer → Macros**.
4. Verify results using **right-click → Table Properties** on each table.

See [tests/README.md](README.md) for full VBA IDE setup instructions.

---

## Subroutine: `NormalizeTables_Light`

### TC-NT-01 · Table width set to 100%

| Field | Detail |
|-------|--------|
| Setup | Document with a table whose preferred width is a fixed point value |
| Expected | Table → Size tab: **100%**, **Percent** selected |
| Pass criteria | Table stretches to full text-area width |

### TC-NT-02 · Table alignment set to Left

| Field | Detail |
|-------|--------|
| Setup | Table with centre or right alignment |
| Expected | Table → Table tab: **Left** alignment selected |
| Pass criteria | Table aligns to the left margin |

### TC-NT-03 · Row height constraints cleared

| Field | Detail |
|-------|--------|
| Setup | Table with rows locked to an exact height |
| Expected | Row → Row tab: **Specify height** checkbox is unchecked |
| Pass criteria | Rows auto-size to their content |

### TC-NT-04 · Cell preferred width unchecked

| Field | Detail |
|-------|--------|
| Setup | Table with fixed column widths |
| Expected | Cell → Cell tab: **Preferred width** checkbox is unchecked |
| Pass criteria | Columns distribute automatically across the table width |

### TC-NT-05 · Cell vertical alignment set to Centre

| Field | Detail |
|-------|--------|
| Setup | Table with cells set to Top or Bottom vertical alignment |
| Expected | Cell → Cell tab: **Center** selected under Vertical alignment |
| Pass criteria | Cell content is vertically centred |

### TC-NT-06 · Merged cell table — no runtime error

| Field | Detail |
|-------|--------|
| Setup | Table containing horizontally or vertically merged cells |
| Expected | Macro completes without a *Value out of range* or any other runtime error |
| Pass criteria | Completion MsgBox appears; all reachable cells normalised |

### TC-NT-07 · Completion message

| Field | Detail |
|-------|--------|
| Setup | Document with 3 tables |
| Expected | MsgBox: **"Done. 3 table(s) normalized."** |
| Pass criteria | Count in message matches actual table count |

### TC-NT-08 · Document with no tables

| Field | Detail |
|-------|--------|
| Setup | Document with no tables |
| Expected | MsgBox: **"Done. 0 table(s) normalized."** — no error raised |
| Pass criteria | Macro exits cleanly with zero count |

---

## Pass Criteria Summary

| TC | Description | Expected outcome |
|----|-------------|------------------|
| TC-NT-01 | Table width | 100% Percent in Table Properties |
| TC-NT-02 | Table alignment | Left |
| TC-NT-03 | Row height | Specify height unchecked |
| TC-NT-04 | Cell width | Preferred width unchecked |
| TC-NT-05 | Cell vertical align | Center |
| TC-NT-06 | Merged cells | No runtime error |
| TC-NT-07 | Completion message | Correct count shown |
| TC-NT-08 | No tables | Clean exit, zero count |
