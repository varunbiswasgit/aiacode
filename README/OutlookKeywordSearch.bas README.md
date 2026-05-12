# OutlookKeywordSearch.bas README

## Purpose

Outlook VBA macro that searches all Outlook folders and subfolders for a keyword or phrase in the **email body**. Returns the **oldest matching email** by received date. Supports two modes: single keyword (interactive) and batch (Excel-driven).

---

## Installation

1. Open Outlook and press **Alt + F11** to open the VBA editor.
2. Go to **Insert → Module**.
3. Open `scripts/OutlookKeywordSearch.bas` and paste the full contents into the new module.
4. Close the VBA editor.
5. Run the macro via **Tools → Macros → RunKeywordSearch**.

No external library references are required. Excel automation uses late binding (`CreateObject`).

---

## Modes

### S — Single Keyword

| Prompt | Input |
|--------|-------|
| Mode | `S` |
| Keyword | Any word or phrase |

**Output:** MsgBox + Immediate Window (`Ctrl + G` in VBE to view)

Output includes: keyword searched, received date/time, subject, sender (display name + email), folder path.

### B — Batch Mode

| Prompt | Input |
|--------|-------|
| Mode | `B` |
| Excel file path | Full path, e.g. `C:\Users\Varun\keywords.xlsx` |
| Keyword column | Column letter, e.g. `A` |

**Output:** Three columns appended automatically at the end of the sheet (no manual output column needed).

| Column | Content |
|--------|---------|
| Match Email | `yyyy-mm-dd hh:nn:ss \| Subject \| FolderPath` |
| Sender | `Display Name <email@address>` |
| Status | `Found`, `Not Found`, or `Blank Keyword` |

The Excel file is saved automatically after all rows are processed.

---

## Search Logic

- Scope: **all Outlook stores, folders, and subfolders** (recursive).
- Field: **Body only** (plain text, case-insensitive via `vbTextCompare`).
- Match rule: **oldest email** — folders are sorted ascending by `ReceivedTime`; the match with the earliest `ReceivedTime` across all folders is kept.
- Items sorted with `Items.Sort "[ReceivedTime]", False` (ascending) before iteration.

---

## Output Fields

| Field | Source property | Notes |
|-------|----------------|-------|
| Received date | `MailItem.ReceivedTime` | Formatted `yyyy-mm-dd hh:nn:ss` |
| Subject | `MailItem.Subject` | Blank if missing |
| Sender | `MailItem.SenderName` + `MailItem.SenderEmailAddress` | Combined as `Name <email>` |
| Folder path | `Folder.FolderPath` | Full path including store name |

---

## Configuration

No configuration file. All inputs are collected at runtime via `InputBox`.

To change the default worksheet targeted in batch mode, modify this line in `RunBatchKeywordSearch`:

```vb
Set ws = wb.Worksheets(1)   ' Change 1 to sheet name or index as needed
```

---

## Known Limitations

- Uses `InStr` on `MailItem.Body` (plain text). Does not search HTML body or attachments.
- Iterates all items in all folders; performance degrades on very large mailboxes. Consider adding a date range filter for large environments.
- `SenderEmailAddress` may return an Exchange alias rather than an SMTP address in some on-premises Exchange configurations.

---

## Version History

| Version | Summary |
|---------|---------|
| v1 | Initial release. Single and batch modes. Body-only search. Oldest match by ReceivedTime. All folders recursive. Auto-append Excel output columns. Late binding for Excel. |
