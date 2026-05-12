# OutlookKeywordSearch_Standalone.bas — Testing Readme

## Test Environment

| Item | Requirement |
|------|-------------|
| Application | Microsoft Outlook (2016 / 2019 / M365) |
| Macro enabled | Yes — Trust Center: Enable macros |
| Excel required | Yes for batch mode only |
| PowerShell | Not required |
| Test mailbox | Any mailbox with known emails in known folders |

---

## TC-01 — Single mode: keyword found

| Field | Value |
|-------|-------|
| Mode | `S` |
| Keyword | A word known to exist in the body of at least one email |
| Expected MsgBox | Shows received date, subject, sender, folder path of oldest matching email |
| Expected Immediate Window | Same content printed between separator lines |
| Pass criteria | Result matches oldest email with that keyword in body; no error shown |

---

## TC-02 — Single mode: keyword not found

| Field | Value |
|-------|-------|
| Mode | `S` |
| Keyword | A string guaranteed not to exist in any email body |
| Expected | MsgBox: `No email found for keyword: <keyword>` |
| Pass criteria | MsgBox shown; no crash |

---

## TC-03 — Single mode: empty keyword

| Field | Value |
|-------|-------|
| Mode | `S` |
| Keyword | (leave blank, click OK) |
| Expected | MsgBox: `No keyword entered.` |
| Pass criteria | Macro exits gracefully |

---

## TC-04 — Batch mode: normal run

| Field | Value |
|-------|-------|
| Mode | `B` |
| Excel file | Workbook with keywords in column A (row 1 = header, rows 2+ = keywords) |
| Keyword column | `A` |
| Expected | Three new columns appended: Match Email, Sender, Status |
| Pass criteria | Found keywords show date/subject/folder and sender; unfound rows show `Not Found`; file saved |

---

## TC-05 — Batch mode: blank keyword rows

| Field | Value |
|-------|-------|
| Setup | Excel file with some blank cells in keyword column |
| Expected | Blank rows show `Blank Keyword` in Status |
| Pass criteria | No crash; all other rows processed normally |

---

## TC-06 — Batch mode: file not found

| Field | Value |
|-------|-------|
| File path | A path that does not exist |
| Expected | MsgBox: `File not found: <path>` |
| Pass criteria | Macro exits; no crash |

---

## TC-07 — Batch mode: invalid column letter

| Field | Value |
|-------|-------|
| Keyword column | `3` or `!` |
| Expected | MsgBox: `Invalid Excel column reference: <input>` |
| Pass criteria | Macro exits; Excel not written; no crash |

---

## TC-08 — Batch mode: auto-append output columns

| Field | Value |
|-------|-------|
| Setup | Excel file with data already in columns A–D |
| Expected | Match Email to column E, Sender to F, Status to G |
| Pass criteria | No existing data overwritten |

---

## TC-09 — Invalid mode entry

| Field | Value |
|-------|-------|
| Mode | Any value other than `S` or `B` (e.g. `X`) |
| Expected | MsgBox: `Invalid mode. Please enter S or B.` |
| Pass criteria | Macro exits gracefully |

---

## TC-10 — Oldest email returned

| Field | Value |
|-------|-------|
| Setup | At least two emails in different folders with same keyword in body |
| Expected | Macro returns the one with the earlier received date |
| Pass criteria | Returned email ReceivedTime is the earliest across all folders |

---

## TC-11 — Non-mail folder skipped

| Field | Value |
|-------|-------|
| Setup | Mailbox with content in Calendar and Contacts |
| Expected | Calendar and Contacts entries not returned as matches |
| Pass criteria | No match returned from non-mail folder names |

---

## Pass Criteria Summary

| TC | Description | Pass criteria |
|----|-------------|---------------|
| TC-01 | Single — found | Oldest match in MsgBox and Immediate Window |
| TC-02 | Single — not found | `No email found` message; no crash |
| TC-03 | Single — empty keyword | Graceful exit |
| TC-04 | Batch — normal run | Columns appended; statuses correct; file saved |
| TC-05 | Batch — blank rows | `Blank Keyword` status; no crash |
| TC-06 | Batch — file not found | MsgBox; macro exits |
| TC-07 | Batch — invalid column | MsgBox; Excel not written |
| TC-08 | Batch — auto-append | New columns after last used column |
| TC-09 | Invalid mode | MsgBox; macro exits |
| TC-10 | Oldest match | Earliest ReceivedTime returned |
| TC-11 | Non-mail folder skip | No Calendar/Contacts matches |
