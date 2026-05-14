# Bold List Prefixes — Testing Guide

---

## Automated VBA Tests

A VBA test harness (`Test_BoldListPrefixes.bas`) covers 7 automated test cases.

**Setup:**
1. Open Microsoft Word.
2. Press `Alt+F11` to open the VBA editor.
3. Import both `BoldListPrefixesOutlook.bas` and `Test_BoldListPrefixes.bas` into the same project.
4. Run `RunAllTests` via `Alt+F8`.
5. Results appear in the Immediate window (`Ctrl+G`) and a summary MsgBox.

---

## Test Environment

| Item | Requirement |
|---|---|
| Host | Microsoft Word 2016 or later **or** Microsoft Outlook 2016 or later |
| VBA enabled | Yes — macros must be enabled |
| Module location | Any standard `.bas` module in the host VBA project |

---

## Automated Test Cases

| TC | Scenario | Pass condition |
|---|---|---|
| TC-01 | Colon delimiter | Bold ends at `:` |
| TC-02 | Dash delimiter | Bold ends at `-` |
| TC-03 | Colon before dash | Colon wins |
| TC-04 | Dash before colon | Dash wins |
| TC-05 | No delimiter | Paragraph skipped |
| TC-10 | Delimiter at position 1 | Paragraph skipped (endPos > 1 guard) |
| TC-06 | Non-list paragraph | Not bolded |

---

## Manual Test Cases

### TC-07 — Outlook compose window

| Field | Detail |
|-------|--------|
| Setup | Open a new Outlook compose message; insert a bulleted list with `Action: send report`. |
| Action | Run `BoldListPrefixesOutlook` from the Outlook VBA editor. |
| Expected | `Action:` is bold. |
| Pass criteria | Macro resolves `ActiveInspector.WordEditor` successfully; same output as TC-01. |

### TC-08 — No open document

| Field | Detail |
|-------|--------|
| Setup | Close all Word documents and all Outlook inspector windows. |
| Action | Run `BoldListPrefixesOutlook`. |
| Expected | A `MsgBox` appears: "No editable document found." |
| Pass criteria | Macro exits cleanly with no unhandled runtime error. |

### TC-09 — Mixed list types in one document

| Field | Detail |
|-------|--------|
| Setup | Document contains one bulleted list item (`Item: detail`) and one numbered list item (`Step 1 - do this`). |
| Action | Run `BoldListPrefixesOutlook`. |
| Expected | Both prefixes bolded correctly and independently. |
| Pass criteria | All `ListType <> 0` paragraphs processed; no cross-contamination. |
