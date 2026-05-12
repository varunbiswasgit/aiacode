# OutlookKeywordSearch — Testing Readme (v2)

## Architecture Change (v2)

From v2, the macro is a **two-file solution**:
- `OutlookKeywordSearch.bas` — VBA launcher only (no search logic)
- `OutlookKeywordSearch.ps1` — PowerShell engine (all search logic)

Most test cases now validate PS script behaviour. The VBA launcher tests validate input collection and PS invocation only.

---

## Test Environment

| Item | Requirement |
|------|-------------|
| Application | Microsoft Outlook (2016 / 2019 / M365), running |
| OS | Windows 10 or Windows 11 |
| PowerShell | 5.1 or 7+ |
| Execution policy | `RemoteSigned` for CurrentUser |
| Macro enabled | Yes — Trust Center: Enable macros |
| Excel required | Yes for batch mode only |
| Log file | `%USERPROFILE%\Documents\OutlookKeywordSearch.log` |

---

## TC-01 — VBA launcher: PS script path not found

| Field | Value |
|-------|-------|
| Setup | Temporarily change `PS_SCRIPT_PATH` in VBA to a non-existent path |
| Expected | MsgBox: `PowerShell script not found: <path>` |
| Pass criteria | Macro exits before launching PS; no crash |

---

## TC-02 — VBA launcher: invalid mode

| Field | Value |
|-------|-------|
| Mode input | `X` |
| Expected | MsgBox: `Invalid mode. Please enter S or B.` |
| Pass criteria | Macro exits; PS script not launched |

---

## TC-03 — Single mode: keyword found

| Field | Value |
|-------|-------|
| Mode | `S` |
| Keyword | A word known to exist in the body of at least one email |
| Expected | Windows toast showing received date, subject, sender, folder path |
| Expected log | `FOUND:` entry in `OutlookKeywordSearch.log` |
| Pass criteria | Toast shows oldest matching email; log entry written |

---

## TC-04 — Single mode: keyword not found

| Field | Value |
|-------|-------|
| Mode | `S` |
| Keyword | A string guaranteed not to exist in any email body |
| Expected | Toast: `No email found for keyword: <keyword>` |
| Pass criteria | Toast shown; log entry written; no crash |

---

## TC-05 — Single mode: empty keyword

| Field | Value |
|-------|-------|
| Mode | `S` |
| Keyword | (leave blank, click OK) |
| Expected | MsgBox: `No keyword entered.` |
| Pass criteria | Macro exits; PS script not launched |

---

## TC-06 — Batch mode: normal run

| Field | Value |
|-------|-------|
| Mode | `B` |
| Excel file | Workbook with keywords in column A (row 1 = header, rows 2+ = keywords) |
| Keyword column | `A` |
| Expected | Three new columns appended: Match Email, Sender, Status |
| Expected toast | `Batch complete. Processed: N \| Found: N` |
| Pass criteria | Found keywords show date/subject/folder and sender; unfound rows show `Not Found`; file saved; toast shown |

---

## TC-07 — Batch mode: blank keyword rows

| Field | Value |
|-------|-------|
| Setup | Excel file with some blank cells in keyword column |
| Expected | Blank rows show `Blank Keyword` in Status column |
| Pass criteria | No crash; all other rows processed normally |

---

## TC-08 — Batch mode: file not found (VBA pre-check)

| Field | Value |
|-------|-------|
| File path | A path that does not exist |
| Expected | MsgBox: `File not found: <path>` from VBA launcher |
| Pass criteria | Macro exits before launching PS; no crash |

---

## TC-09 — Batch mode: invalid column letter

| Field | Value |
|-------|-------|
| Keyword column | `3` or `!` |
| Expected | Toast + log: `ERROR: Invalid column: <input>` |
| Pass criteria | PS exits without writing to Excel; no crash |

---

## TC-10 — Batch mode: auto-append output columns

| Field | Value |
|-------|-------|
| Setup | Excel file with data already in columns A–D |
| Expected | Match Email written to column E, Sender to F, Status to G |
| Pass criteria | No existing data overwritten; new columns placed correctly |

---

## TC-11 — Optimization 3: non-mail folder skipped

| Field | Value |
|-------|-------|
| Setup | Note the time taken searching a mailbox with large Calendar/Contacts |
| Run 1 | v1 (no folder skip) |
| Run 2 | v2 (with folder skip) |
| Expected | v2 completes faster; Calendar and Contacts entries not returned as matches |
| Pass criteria | Log shows no entries for skipped folder names; elapsed time lower |

---

## TC-12 — Optimization 4: early exit per folder

| Field | Value |
|-------|-------|
| Setup | Mailbox with a known keyword in an old email; many newer emails exist after it |
| Expected | Once oldest match in a folder is found, remaining newer items in that folder are not iterated |
| Pass criteria | Result returned is oldest match; no newer email overrides it |

---

## TC-13 — Oldest match rule (cross-folder)

| Field | Value |
|-------|-------|
| Setup | Same keyword in two emails in different folders; note received dates |
| Expected | PS returns the one with the earlier received date |
| Pass criteria | `ReceivedTime` of returned email is the earliest across all folders |

---

## TC-14 — Log file written

| Field | Value |
|-------|-------|
| After any run | Open `%USERPROFILE%\Documents\OutlookKeywordSearch.log` |
| Expected | Timestamped entries for search start, each keyword, found results, search end |
| Pass criteria | Log file exists and contains correct entries |

---

## Pass Criteria Summary

| TC | Description | Pass criteria |
|----|-------------|---------------|
| TC-01 | PS script not found | MsgBox shown; PS not launched |
| TC-02 | Invalid mode | MsgBox shown; PS not launched |
| TC-03 | Single — found | Toast shows oldest match; log entry written |
| TC-04 | Single — not found | Toast shows not-found message |
| TC-05 | Single — empty keyword | MsgBox; macro exits |
| TC-06 | Batch — normal run | Columns appended; toast on completion |
| TC-07 | Batch — blank rows | `Blank Keyword` status; no crash |
| TC-08 | Batch — file not found | MsgBox from VBA; PS not launched |
| TC-09 | Batch — invalid column | Toast/log error; Excel not written |
| TC-10 | Batch — auto-append | New columns placed after last used column |
| TC-11 | Opt 3: folder skip | Non-mail folders not iterated; faster run |
| TC-12 | Opt 4: early exit | Remaining newer items in folder skipped |
| TC-13 | Oldest match cross-folder | Earliest ReceivedTime returned |
| TC-14 | Log file | Timestamped log written after every run |
