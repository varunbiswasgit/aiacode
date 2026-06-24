# email-campaign-tracker

An Outlook VBA macro that scans a selected Outlook folder for replies from a list of managers and produces a fixed-name overwrite tracker report — `Manager_Response_Tracke.xlsx` — saved to the same folder as the manager list Excel file.

---

## Purpose

Automates the process of checking which managers on a campaign list have responded to an email. Each run overwrites the same output file rather than creating dated copies.

---

## Files

| File | Purpose |
|------|--------|
| `ManagerResponseTracker.bas` | VBA module — import into Outlook VBA editor (`Alt+F11`) |
| `README.md` | This file |
| `TESTING.md` | Manual test cases |

---

## Configuration

All inputs are collected interactively at runtime via dialog boxes:

| Prompt | What to enter |
|--------|-------------|
| Pick Folder | Outlook folder to scan for replies |
| Exclude Sender | Email address to ignore (e.g. your own address) |
| Select Excel File | Manager list workbook (picker dialog) |
| Column Letter | Column in the manager list containing email addresses (e.g. `B`) |

---

## Output

The macro writes `Manager_Response_Tracke.xlsx` to the same folder as the manager list file. If the file is already open, it shows a warning and exits without saving.

### Output columns

| Column | Description |
|--------|------------|
| Manager Email | Email address from the manager list |
| Response Received | Yes / No |
| Latest Email Time | Timestamp of the most recent reply |
| Latest Email Subject | Subject line of the most recent reply |
| Last Email has Attachment? | Yes / No |
| Excel Seen In Thread | Yes if any email in the thread had an `.xlsx` attachment |
| Newest Attachment Name | Filename of the most recent `.xlsx` attachment |
| Clarification Required | Yes if no attachment found on the latest email |

---

## Logic Flow

1. Prompt user to pick an Outlook folder.
2. Prompt for excluded sender email address.
3. Open Excel via late binding; prompt for manager list file and email column.
4. Build a dictionary of manager email addresses from the list.
5. Scan Outlook folder items sorted newest first.
6. For each mail item not from the excluded sender, record the latest email and track Excel attachments.
7. Check if output file is open — if so, warn and exit.
8. Write report to a new workbook and save as `Manager_Response_Tracke.xlsx`, overwriting any previous version.
9. Close Excel and clean up.

---

## How to Install

1. Open Outlook.
2. Press `Alt+F11` to open the VBA editor.
3. Go to **File → Import File** and select `ManagerResponseTracker.bas`.
4. Press `Alt+F8`, select `BuildManagerResponseTracker`, and click **Run**.

---

## Version History

| Version | Change |
|---------|-------|
| v1 | Initial release — manager response tracker with fixed output filename and open-file warning |
