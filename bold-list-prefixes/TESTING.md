# BoldListPrefixesOutlook.bas — Testing Readme

Test coverage for `bold-list-prefixes/BoldListPrefixesOutlook.bas`.

## Automated Tests

No automated test harness exists. The macro relies on live Word/Outlook document objects that cannot be driven without UI mocking.

## Manual Test Cases

Environment setup:

1. Open a Word document or an Outlook compose window.
2. Press **Alt+F11**, import `bold-list-prefixes/BoldListPrefixesOutlook.bas`.
3. Run `BoldListPrefixesOutlook` from **Alt+F8 → Run**.

---

## Logic Flow

1. Resolve the active document — try `Application.ActiveDocument` first; fall back to `Application.ActiveInspector.WordEditor` for Outlook.
2. If no editable document is found, show a `MsgBox` and exit cleanly.
3. Loop over every `Paragraph` in the document.
4. Skip paragraphs where `ListFormat.ListType = 0` (not a list item).
5. For each list paragraph, locate the first `:` and first `-` in the paragraph text.
6. Select the earlier delimiter; if neither is present, skip the paragraph.
7. Duplicate the paragraph range, trim its end to `Start + endPos - 1`, and apply `Font.Bold = True`.

---

## Test Cases

### TC-01 · Colon delimiter

| Step | Action |
|------|--------|
| Setup | Bulleted list item: `"Status: In Progress"` |
| Expected | `"Status:"` is bolded; `" In Progress"` is not |

### TC-02 · Dash delimiter

| Step | Action |
|------|--------|
| Setup | Bulleted list item: `"Owner - Varun"` |
| Expected | `"Owner -"` is bolded; `" Varun"` is not |

### TC-03 · Colon before dash

| Step | Action |
|------|--------|
| Setup | `"Type: high-priority"` |
| Expected | `"Type:"` is bolded (colon wins as first delimiter) |

### TC-04 · Dash before colon

| Step | Action |
|------|--------|
| Setup | `"Note - see: below"` |
| Expected | `"Note -"` is bolded (dash appears first) |

### TC-05 · No delimiter

| Step | Action |
|------|--------|
| Setup | List item with no `:` or `-` |
| Expected | Paragraph skipped; no formatting change |

### TC-06 · Delimiter is first character

| Step | Action |
|------|--------|
| Setup | List item starting with `:text` |
| Expected | Paragraph skipped (`endPos = 1` guard) |

### TC-07 · Non-list paragraph ignored

| Step | Action |
|------|--------|
| Setup | Normal paragraph (not in a list): `"Summary: done"` |
| Expected | Paragraph skipped; no formatting change |

### TC-08 · No active document

| Step | Action |
|------|--------|
| Setup | Run macro with no Word document or Outlook inspector open |
| Expected | `MsgBox` shown: `"No editable document found."` Macro exits cleanly |

### TC-09 · Outlook inspector

| Step | Action |
|------|--------|
| Setup | Open Outlook compose window with a bulleted list containing `"Action: review"` |
| Expected | `"Action:"` is bolded; rest of the item unchanged |

### TC-10 · Mixed list and non-list paragraphs

| Step | Action |
|------|--------|
| Setup | Document with 2 list items and 3 normal paragraphs, all containing `:` |
| Expected | Only the 2 list items are bolded up to their first delimiter |

---

## Version History

| Version | Summary |
|---|---|
| v1 | Initial release |
| v1.1 | Renamed from `BoldListPrefixes` to `BoldListPrefixesOutlook` |
