# TESTING — email-campaign-tracker

Manual test cases for `ManagerResponseTracker.bas`.

---

## Environment

- Outlook desktop (Microsoft 365 or Outlook 2016+)
- Excel desktop (same version as Outlook)
- VBA editor accessible via `Alt+F11`

---

## Test Cases

### TC-01 — Normal run, all managers replied

| Field | Detail |
|-------|-------|
| **Setup** | Manager list Excel file with 3 manager email addresses in column B. All 3 have sent at least one reply in the selected Outlook folder. Output file does not exist. |
| **Action** | Run `BuildManagerResponseTracker`. Select folder, enter excluded sender, select manager list file, enter `B`. |
| **Expected** | `Manager_Response_Tracke.xlsx` created in the same folder as the manager list. All 3 rows show `Response Received = Yes`. |
| **Pass** | File exists, 3 data rows, no VBA error. |

---

### TC-02 — Some managers did not reply

| Field | Detail |
|-------|-------|
| **Setup** | Manager list with 4 managers. Only 2 have replies in the folder. |
| **Action** | Run macro as above. |
| **Expected** | 4 rows. 2 show `Yes`, 2 show `No` with blank time/subject. |
| **Pass** | Correct Yes/No split, no error. |

---

### TC-03 — Excluded sender filtered out

| Field | Detail |
|-------|-------|
| **Setup** | One of the "manager" replies is actually from your own email address. Enter your address as the excluded sender. |
| **Action** | Run macro. |
| **Expected** | That manager shows `Response Received = No` (your reply is ignored). |
| **Pass** | Self-reply excluded correctly. |

---

### TC-04 — Output file already open

| Field | Detail |
|-------|-------|
| **Setup** | `Manager_Response_Tracke.xlsx` is open in Excel from a previous run. |
| **Action** | Run macro. |
| **Expected** | Message box: "Please close Manager_Response_Tracke.xlsx before running the macro." Macro exits without error. File on disk is not modified. |
| **Pass** | Warning shown, no runtime error, file unchanged. |

---

### TC-05 — Overwrite previous output

| Field | Detail |
|-------|-------|
| **Setup** | `Manager_Response_Tracke.xlsx` exists from a previous run but is **closed**. |
| **Action** | Run macro. |
| **Expected** | File is overwritten silently. New run data appears. No prompt. |
| **Pass** | File updated, no prompt, no error. |

---

### TC-06 — Manager with Excel attachment

| Field | Detail |
|-------|-------|
| **Setup** | One manager reply includes an `.xlsx` attachment. |
| **Action** | Run macro. |
| **Expected** | `Last Email has Attachment? = Yes`, `Excel Seen In Thread = Yes`, attachment filename populated, `Clarification Required = No`. |
| **Pass** | All four attachment columns correct. |

---

### TC-07 — Manager with no attachment

| Field | Detail |
|-------|-------|
| **Setup** | One manager replied but with no `.xlsx` attachment. |
| **Action** | Run macro. |
| **Expected** | `Last Email has Attachment? = No`, `Excel Seen In Thread = No`, `Clarification Required = Yes`. |
| **Pass** | Correct values in all attachment columns. |

---

### TC-08 — User cancels folder picker

| Field | Detail |
|-------|-------|
| **Setup** | Any valid setup. |
| **Action** | Run macro. When the folder picker opens, click Cancel. |
| **Expected** | Macro exits silently. No output file created. No error. |
| **Pass** | Silent exit. |
