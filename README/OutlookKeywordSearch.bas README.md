# OutlookKeywordSearch.bas README

## Purpose

Two-file solution for searching Outlook email bodies for keywords.

- **`OutlookKeywordSearch.bas`** — thin VBA launcher (Outlook macro). Collects user inputs and fires the PowerShell script as a background process. Outlook UI remains fully responsive.
- **`OutlookKeywordSearch.ps1`** — PowerShell engine. Does all the searching, Excel writing, and result notification out-of-process.

Supports two modes: **Single** (one keyword → Windows toast + log) and **Batch** (Excel-driven → appends results columns to same workbook).

---

## Installation

### Step 1 — Copy the PowerShell script

Copy `scripts/OutlookKeywordSearch.ps1` to a local folder, e.g.:
```
C:\Users\Varun\scripts\OutlookKeywordSearch.ps1
```

### Step 2 — Import the VBA launcher into Outlook

1. Open Outlook → press **Alt + F11**
2. Go to **Insert → Module**
3. Paste the full contents of `scripts/OutlookKeywordSearch.bas`
4. Update the `PS_SCRIPT_PATH` constant at the top of the module to match where you saved the `.ps1` file
5. Close the VBA editor
6. Run via **Tools → Macros → RunKeywordSearch**

### Step 3 — Allow PowerShell execution

Run once in an elevated PowerShell window:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

---

## Modes

### S — Single Keyword

| Prompt | Input |
|--------|-------|
| Mode | `S` |
| Keyword | Any word or phrase |

**Output:** Windows toast notification + log entry in `Documents\OutlookKeywordSearch.log`

### B — Batch Mode

| Prompt | Input |
|--------|-------|
| Mode | `B` |
| Excel file path | Full path, e.g. `C:\Users\Varun\keywords.xlsx` |
| Keyword column | Column letter, e.g. `A` |

**Output:** Three columns appended automatically at the end of the sheet.

| Column | Content |
|--------|---------|
| Match Email | `yyyy-MM-dd HH:mm:ss \| Subject \| FolderPath` |
| Sender | `Display Name <email@address>` |
| Status | `Found`, `Not Found`, or `Blank Keyword` |

Windows toast notification shown on completion. Excel file saved automatically.

---

## Performance Optimizations (v2)

| # | Optimization | Where implemented |
|---|---|---|
| 3 | Skip non-mail folders (Calendar, Contacts, Tasks, Junk, Deleted Items, Drafts, etc.) by name and `DefaultItemType` | `Test-SkipFolder` in PS script |
| 4 | Early exit per folder once current item ReceivedTime exceeds best match — folders sorted ascending so remaining items are guaranteed newer | `Search-FolderRecursive` in PS script |
| 5 | OS yield between batch keywords via `Thread.Sleep(0)` | Batch loop in PS script |
| 6 | Entire search runs in a separate PowerShell process — Outlook UI never blocked | VBA `Shell` launcher |

---

## Search Logic

- **Scope:** All Outlook stores, folders, and subfolders (recursive), excluding skipped folders.
- **Field:** Body only (plain text, case-insensitive via `StringComparison.OrdinalIgnoreCase`).
- **Match rule:** Oldest email — folders sorted ascending by `ReceivedTime`; earliest match across all folders is kept.
- **Early exit:** Once a match is found in a folder, remaining items in that folder (which are newer due to sort order) are skipped.

---

## Log File

All runs are logged to:
```
%USERPROFILE%\Documents\OutlookKeywordSearch.log
```
The log captures start/end timestamps, each keyword searched, found results, and any folder-level errors.

---

## Output Fields

| Field | Source | Notes |
|-------|--------|-------|
| Received date | `MailItem.ReceivedTime` | Formatted `yyyy-MM-dd HH:mm:ss` |
| Subject | `MailItem.Subject` | Blank if missing |
| Sender | `MailItem.SenderName` + `SenderEmailAddress` | Combined as `Name <email>` |
| Folder path | `Folder.FolderPath` | Full path including store name |

---

## Configuration

| Setting | Where | Default |
|---------|-------|---------|
| PS script path | `PS_SCRIPT_PATH` constant in VBA module | `C:\Users\Varun\scripts\OutlookKeywordSearch.ps1` |
| Target worksheet | `$wb.Worksheets.Item(1)` in PS script | Sheet 1 |
| Log file location | `$LogPath` in PS script | `Documents\OutlookKeywordSearch.log` |

---

## Known Limitations

- Body search uses plain text (`MailItem.Body`). Does not search HTML body or attachments.
- `SenderEmailAddress` may return an Exchange alias in some on-premises Exchange environments.
- Outlook must be running when the PS script executes (COM interop requires a live Outlook instance).
- Toast notifications require Windows 10/11. On older OS versions the result is logged only.

---

## Version History

| Version | Summary |
|---------|---------|
| v1 | Initial release. Single and batch modes. Body-only search. Oldest match by ReceivedTime. All folders recursive. Auto-append Excel output columns. Late binding for Excel. |
| v2 | Architecture change: VBA is now a thin launcher only. All search logic moved to OutlookKeywordSearch.ps1 running as a separate process. Added optimizations 3 (skip non-mail folders), 4 (early exit per folder), 5 (OS yield between batch keywords), 6 (out-of-process execution). Single mode result delivered via Windows toast + log. |
