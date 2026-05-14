# OutlookKeywordSearch — Testing

## PS-Assisted (`OutlookKeywordSearch_PS.bas` + `OutlookKeywordSearch_PS.ps1`)

### Test Environment

| Item | Requirement |
|------|-------------|
| Application | Microsoft Outlook (2016 / 2019 / M365), running |
| OS | Windows 10 or Windows 11 |
| PowerShell | 5.1 or 7+ |
| Execution policy | `RemoteSigned` for CurrentUser |
| Macro enabled | Yes — Trust Center: Enable macros |
| Excel required | Yes for batch mode only |
| Log file | `%USERPROFILE%\Documents\OutlookKeywordSearch.log` |

| TC | Description | Pass criteria |
|----|-------------|---------------|
| TC-01 | PS script not found | MsgBox; PS not launched |
| TC-02 | Invalid mode | MsgBox; PS not launched |
| TC-03 | Single — found | Toast + log entry with oldest match |
| TC-04 | Single — not found | Toast not-found message |
| TC-05 | Single — empty keyword | MsgBox; macro exits |
| TC-06 | Batch — normal run | Columns appended; toast shown |
| TC-07 | Batch — blank rows | `Blank Keyword` status; no crash |
| TC-08 | Batch — file not found | MsgBox from VBA; PS not launched |
| TC-09 | Batch — invalid column | Toast/log error; Excel not written |
| TC-10 | Batch — auto-append | Columns placed after last used column |
| TC-11 | Non-mail folder skip | No Calendar/Contacts matches |
| TC-12 | Early exit per folder | Newest items in folder skipped |
| TC-13 | Oldest match cross-folder | Earliest ReceivedTime returned |
| TC-14 | Log file | Timestamped log written after every run |

---

## Standalone (`OutlookKeywordSearch_Standalone.bas`)

### Test Environment

| Item | Requirement |
|------|-------------|
| Application | Microsoft Outlook (2016 / 2019 / M365) |
| Macro enabled | Yes — Trust Center: Enable macros |
| Excel required | Yes for batch mode only |
| PowerShell | Not required |

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
