# OutlookKeywordSearch

Two implementations for searching Outlook email bodies by keyword — choose based on your environment.

- **`OutlookKeywordSearch_PS.bas`** — thin VBA launcher (Outlook macro). Collects user inputs and fires the PowerShell script as a background process. Outlook UI remains fully responsive.
- **`OutlookKeywordSearch_PS.ps1`** — PowerShell engine. Does all searching, Excel writing, and result notification out-of-process.
- **`OutlookKeywordSearch_Standalone.bas`** — pure VBA macro. No external dependencies. Runs entirely inside Outlook.

Single mode: result delivered via Windows toast notification (PS) or MsgBox (Standalone).
Batch mode: appends Match Email, Sender, and Status columns to the Excel workbook.

---

## When to use which script

Use `OutlookKeywordSearch_PS.bas` + `OutlookKeywordSearch_PS.ps1` when:
- You want Outlook to remain fully responsive during search
- You want Windows toast notifications on completion
- You want a persistent log file of all searches

Use `OutlookKeywordSearch_Standalone.bas` when:
- PowerShell execution is restricted by group policy
- You prefer a single-file, no-dependency solution

---

## Installation — PS-Assisted

### Step 1 — Copy the PowerShell script

Copy `OutlookKeywordSearch_PS.ps1` to a local folder, e.g.:
```
C:\Users\Varun\scripts\OutlookKeywordSearch_PS.ps1
```

### Step 2 — Import the VBA launcher into Outlook

1. Open Outlook → press **Alt + F11**
2. Go to **Insert → Module**
3. Paste the full contents of `OutlookKeywordSearch_PS.bas`
4. Update the `PS_SCRIPT_PATH` constant at the top to match where you saved the `.ps1` file
5. Close the VBA editor
6. Run via **Tools → Macros → RunKeywordSearch**

### Step 3 — Allow PowerShell execution

Run once in an elevated PowerShell window:
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

---

## Installation — Standalone

1. Open Outlook → press **Alt + F11**
2. Go to **Insert → Module**
3. Paste the full contents of `OutlookKeywordSearch_Standalone.bas`
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

**PS output:** Windows toast notification + log entry in `Documents\OutlookKeywordSearch.log`

**Standalone output:** MsgBox + Immediate Window (`Ctrl+G` in VBE)

### B — Batch Mode

| Prompt | Input |
|--------|-------|
| Mode | `B` |
| Excel file path | Full path, e.g. `C:\Users\Varun\keywords.xlsx` |
| Keyword column | Column letter, e.g. `A` |

**Output:** Three columns appended automatically.

| Column | Content |
|--------|---------|
| Match Email | `yyyy-MM-dd HH:mm:ss \| Subject \| FolderPath` |
| Sender | `Display Name <email@address>` |
| Status | `Found`, `Not Found`, or `Blank Keyword` |

---

## Search Logic

- **Scope:** All Outlook stores, folders, and subfolders (recursive), excluding non-mail folders.
- **Field:** Body only (plain text, case-insensitive).
- **Match rule:** Oldest email — earliest `ReceivedTime` across all folders.
- **Non-mail folders skipped:** Calendar, Contacts, Tasks, Junk Email, Deleted Items, Drafts, and others.
- **Early exit:** Once oldest match in a folder is found, remaining newer items are skipped.

---

## Log File (PS version only)

```
%USERPROFILE%\Documents\OutlookKeywordSearch.log
```

---

## Configuration

| Setting | Version | Where | Default |
|---------|---------|-------|---------|
| PS script path | PS | `PS_SCRIPT_DEFAULT` constant in VBA module | `C:\Users\Varun\scripts\OutlookKeywordSearch_PS.ps1` |
| Target worksheet | PS | `$wb.Worksheets.Item(1)` in PS script | Sheet 1 |
| Log file path | PS | `$LogPath` in PS script | `Documents\OutlookKeywordSearch.log` |
| Target worksheet | Standalone | `wb.Worksheets(1)` via late-bound Excel object | Sheet 1 |
| Keyword column | Standalone | Entered at runtime via `InputBox` | None (required) |
| Excel file path | Standalone | Entered at runtime via `InputBox` | None (required) |

The Standalone version has no hardcoded paths or constants. All inputs are collected at run time.

---

## Batch Mode — Data Detection (Standalone)

The Standalone script auto-detects where keyword data begins rather than assuming row 2:

- **`firstDataRow`** — scans from row 1 downward in the keyword column; the first non-empty, non-numeric cell is treated as the data start row. This handles files where the header is in row 1 or where there are blank rows above the data.
- **`lastCol`** — determined by scanning only the keyword column data range (rows `firstDataRow` to `lastRow`), not the entire used range. This prevents stray content in unrelated columns from pushing the output columns to the wrong position.

The three output columns (Match Email, Sender, Status) are always appended immediately after the last used column in the keyword data range.

---

## Known Limitations

| Limitation | Applies to |
|---|---|
| Body search uses plain text only — does not search HTML body or attachments | Both |
| Outlook must be running during the search | Both |
| `SenderEmailAddress` may return an Exchange alias in some on-premises environments | Both |
| Toast notifications require Windows 10/11 | PS only |
| Results shown in MsgBox only — no toast notification, no log file | Standalone only |
| PowerShell execution policy must allow `RemoteSigned` for CurrentUser | PS only |

---

## Version History

| Version | Summary |
|---------|---------|
| v1 (PS) | Initial PS-assisted release. VBA launcher + PS engine. Background execution, toast notifications, log file. Non-mail folder skip, early exit per folder, OS yield between batch keywords. |
| v1 (Standalone) | Initial standalone release. Pure VBA, no PS dependency. Single and batch modes. Body-only search. Oldest match by ReceivedTime. Non-mail folder skip. Early exit per folder. DoEvents between batch rows. Auto-detects `firstDataRow`; scopes `lastCol` to keyword column data range only. |
