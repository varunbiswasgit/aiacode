# OutlookKeywordSearch_Standalone.bas README

## Purpose

Pure VBA Outlook macro. No external dependencies — no PowerShell, no COM servers beyond Outlook and Excel. Runs entirely inside the Outlook VBA IDE.

Single mode: result shown in MsgBox and Immediate Window (`Ctrl+G` in VBE).
Batch mode: reads keywords from an Excel file, appends Match Email, Sender, and Status columns to the same workbook.

---

## When to use this script

Use `OutlookKeywordSearch_Standalone.bas` when:
- PowerShell execution is restricted by group policy
- You prefer a single-file solution with no external dependencies
- You do not need background execution (Outlook is acceptable to pause briefly during search)

Use `OutlookKeywordSearch_PS.bas` + `OutlookKeywordSearch_PS.ps1` when:
- You want Outlook to remain fully responsive during search
- You want Windows toast notifications on completion
- You want a persistent log file of all searches

---

## Installation

1. Open Outlook → press **Alt + F11**
2. Go to **Insert → Module**
3. Paste the full contents of `scripts/OutlookKeywordSearch_Standalone.bas`
4. Close the VBA editor
5. Run via **Tools → Macros → RunKeywordSearch**

No library references required. Excel automation uses late binding (`CreateObject`).

---

## Modes

### S — Single Keyword

| Prompt | Input |
|--------|-------|
| Mode | `S` |
| Keyword | Any word or phrase |

**Output:** MsgBox + Immediate Window (`Ctrl+G` in VBE)

Output includes: keyword searched, received date/time, subject, sender, folder path.

### B — Batch Mode

| Prompt | Input |
|--------|-------|
| Mode | `B` |
| Excel file path | Full path, e.g. `C:\Users\Varun\keywords.xlsx` |
| Keyword column | Column letter, e.g. `A` |

**Output:** Three columns appended automatically at the end of the sheet.

| Column | Content |
|--------|---------|
| Match Email | `yyyy-mm-dd hh:nn:ss \| Subject \| FolderPath` |
| Sender | `Display Name <email@address>` |
| Status | `Found`, `Not Found`, or `Blank Keyword` |

Excel file saved automatically after all rows are processed.

---

## Search Logic

- **Scope:** All Outlook stores, folders, and subfolders (recursive), excluding non-mail folders.
- **Field:** Body only (plain text, case-insensitive via `vbTextCompare`).
- **Match rule:** Oldest email — folders sorted ascending by `ReceivedTime`; earliest match across all folders is kept.
- **Early exit:** Once oldest match in a folder is found, remaining newer items in that folder are skipped.
- **Non-mail folders skipped:** Calendar, Contacts, Tasks, Junk Email, Deleted Items, Drafts, and others.

---

## Output Fields

| Field | Source | Notes |
|-------|--------|-------|
| Received date | `MailItem.ReceivedTime` | Formatted `yyyy-mm-dd hh:nn:ss` |
| Subject | `MailItem.Subject` | Blank if missing |
| Sender | `MailItem.SenderName` + `SenderEmailAddress` | Combined as `Name <email>` |
| Folder path | `Folder.FolderPath` | Full path including store name |

---

## Configuration

No configuration file. All inputs collected at runtime via `InputBox`.

To change the default worksheet targeted in batch mode, modify:
```vb
Set ws = wb.Worksheets(1)   ' Change 1 to sheet name or index
```

---

## Known Limitations

- Body search uses plain text (`MailItem.Body`). Does not search HTML body or attachments.
- Outlook UI may pause briefly during large mailbox searches (no background execution).
- `SenderEmailAddress` may return an Exchange alias in some on-premises Exchange environments.

---

## Version History

| Version | Summary |
|---------|---------|
| v1 | Initial standalone release. Extracted from combined script. Pure VBA, no PS dependency. Single and batch modes. Body-only search. Oldest match by ReceivedTime. Non-mail folder skip. Early exit per folder. DoEvents between batch rows. |
