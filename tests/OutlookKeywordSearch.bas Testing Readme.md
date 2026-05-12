# OutlookKeywordSearch.bas — Testing Readme

## Test Environment

| Item | Requirement |
|------|-------------|
| Application | Microsoft Outlook (2016 / 2019 / M365) |
| Macro enabled | Yes — Trust Center: Enable macros |
| Excel required | Yes for batch mode only |
| Test mailbox | Any mailbox with known emails in known folders |

---

## One-Time Setup

Follow the VBA IDE setup steps in [`tests/README.md`](README.md) (Steps 1–6) before running any test.

Import `scripts/OutlookKeywordSearch.bas` into an Outlook VBA module. No library references are needed.

---

## Test Cases

### TC-01 — Single mode: keyword found

| Field | Value |
|-------|-------|
| Mode input | `S` |
| Keyword | A word known to exist in the body of at least one email |
| Expected MsgBox | Shows received date, subject, sender, folder path of the oldest matching email |
| Expected Immediate Window | Same content printed between separator lines |
| Pass criteria | Result matches the oldest email with that keyword in body; no error shown |

---

### TC-02 — Single mode: keyword not found

| Field | Value |
|-------|-------|
| Mode input | `S` |
| Keyword | A string guaranteed not to exist in any email body |
| Expected MsgBox | `No email found for keyword: <keyword>` |
| Pass criteria | MsgBox shown; no crash; Immediate Window prints the same message |

---

### TC-03 — Single mode: empty keyword

| Field | Value |
|-------|-------|
| Mode input | `S` |
| Keyword | (leave blank, click OK) |
| Expected | MsgBox: `No keyword entered.` |
| Pass criteria | Macro exits gracefully without searching |

---

### TC-04 — Batch mode: normal run

| Field | Value |
|-------|-------|
| Mode input | `B` |
| Excel file | Workbook with keywords in column A (row 1 = header, rows 2+ = keywords) |
| Keyword column | `A` |
| Expected | Three new columns appended: Match Email, Sender, Status |
| Pass criteria | Found keywords show date/subject/folder and sender; unfound rows show `Not Found`; file saved |

---

### TC-05 — Batch mode: blank keyword rows

| Field | Value |
|-------|-------|
| Setup | Excel file with some blank cells in keyword column |
| Expected | Blank rows show `Blank Keyword` in Status; all other rows processed normally |
| Pass criteria | No crash; blank rows handled silently |

---

### TC-06 — Batch mode: file not found

| Field | Value |
|-------|-------|
| File path | A path that does not exist |
| Expected | MsgBox: `File not found: <path>` |
| Pass criteria | Macro exits without opening Excel; no crash |

---

### TC-07 — Batch mode: invalid column letter

| Field | Value |
|-------|-------|
| Keyword column | `3` or `!` (not a valid letter) |
| Expected | MsgBox: `Invalid Excel column reference: <input>` |
| Pass criteria | Macro exits without writing to Excel; no crash |

---

### TC-08 — Batch mode: output columns appended automatically

| Field | Value |
|-------|-------|
| Setup | Excel file with data already in columns A–D |
| Expected | Match Email written to column E, Sender to F, Status to G |
| Pass criteria | No existing data overwritten; new columns placed at E, F, G |

---

### TC-09 — Invalid mode entry

| Field | Value |
|-------|-------|
| Mode input | Any value other than `S` or `B` (e.g. `X`) |
| Expected | MsgBox: `Invalid mode. Please enter S or B.` |
| Pass criteria | Macro exits gracefully; no search performed |

---

### TC-10 — Oldest email returned

| Field | Value |
|-------|-------|
| Setup | At least two emails in different folders with the same keyword in body; note the received dates |
| Expected | Macro returns the one with the earlier received date |
| Pass criteria | Returned email ReceivedTime is the earliest among all matches |

---

## Pass Criteria Summary

| TC | Description | Pass criteria |
|----|-------------|---------------|
| TC-01 | Keyword found — single mode | Oldest match shown in MsgBox and Immediate Window |
| TC-02 | Keyword not found — single mode | `No email found` message; no crash |
| TC-03 | Empty keyword — single mode | Graceful exit with `No keyword entered` |
| TC-04 | Normal batch run | Columns appended; statuses correct; file saved |
| TC-05 | Blank keyword rows | `Blank Keyword` status; no crash |
| TC-06 | File not found | `File not found` message; no crash |
| TC-07 | Invalid column | `Invalid column` message; no crash |
| TC-08 | Auto-append columns | New columns placed after last used column |
| TC-09 | Invalid mode | `Invalid mode` message; no crash |
| TC-10 | Oldest match rule | Earliest ReceivedTime returned across all folders |
