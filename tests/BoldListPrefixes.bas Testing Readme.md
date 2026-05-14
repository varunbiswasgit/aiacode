# BoldListPrefixes.bas — Testing README

## Test Environment

| Item | Requirement |
|---|---|
| Host | Microsoft Word 2016 or later **or** Microsoft Outlook 2016 or later |
| VBA enabled | Yes — macros must be enabled |
| Module location | Any standard `.bas` module in the host VBA project |
| Test document | A `.docx` or Outlook compose window with the list items described below |

---

## Manual Test Cases

### TC-01 — Colon delimiter, bulleted list

**Setup:** Bulleted list containing: `Scope: defines the boundary of work`

**Action:** Run `BoldListPrefixes`.

**Expected:** `Scope:` is bold; ` defines the boundary of work` remains normal weight.

**Pass criteria:** Bold ends at and includes the colon; nothing after the colon is bold.

---

### TC-02 — Dash delimiter, numbered list

**Setup:** Numbered list containing: `1. Owner - accountable for delivery`

**Action:** Run `BoldListPrefixes`.

**Expected:** `Owner -` is bold; ` accountable for delivery` is normal weight.

**Pass criteria:** Bold ends at and includes the dash.

---

### TC-03 — Colon appears before dash

**Setup:** Bulleted list containing: `Priority: high - urgent`

**Action:** Run `BoldListPrefixes`.

**Expected:** `Priority:` is bold (colon at position 9, dash at position 16 → colon wins).

**Pass criteria:** Bold stops at the colon, not the dash.

---

### TC-04 — Dash appears before colon

**Setup:** Bulleted list containing: `Step-by-step: follow the guide`

**Action:** Run `BoldListPrefixes`.

**Expected:** `Step-` is bold (first dash at position 5, colon at position 13 → dash wins).

**Pass criteria:** Bold stops at the first dash.

---

### TC-05 — No delimiter present

**Setup:** Bulleted list containing: `Plain list item with no delimiter`

**Action:** Run `BoldListPrefixes`.

**Expected:** No formatting change.

**Pass criteria:** Paragraph is skipped without error or side effect.

---

### TC-06 — Non-list paragraph ignored

**Setup:** A normal (non-list) paragraph: `Introduction: this section covers basics`

**Action:** Run `BoldListPrefixes`.

**Expected:** No formatting change to this paragraph.

**Pass criteria:** `ListFormat.ListType = 0` guard prevents any bolding.

---

### TC-07 — Outlook compose window

**Setup:** Open a new Outlook compose message; insert a bulleted list with `Action: send report`.

**Action:** Run `BoldListPrefixes` from the Outlook VBA editor.

**Expected:** `Action:` is bold.

**Pass criteria:** Macro resolves `ActiveInspector.WordEditor` successfully; same output as TC-01.

---

### TC-08 — No open document

**Setup:** Close all Word documents and all Outlook inspector windows.

**Action:** Run `BoldListPrefixes`.

**Expected:** A `MsgBox` appears: "No editable document found."

**Pass criteria:** Macro exits cleanly with no unhandled runtime error.

---

### TC-09 — Mixed list types in one document

**Setup:** Document contains one bulleted list item (`Item: detail`) and one numbered list item (`Step 1 - do this`).

**Action:** Run `BoldListPrefixes`.

**Expected:** Both prefixes bolded correctly and independently.

**Pass criteria:** All `ListType <> 0` paragraphs processed; no cross-contamination.

---

### TC-10 — Delimiter at position 1 (edge case guard)

**Setup:** Bulleted list containing: `: orphan colon`

**Action:** Run `BoldListPrefixes`.

**Expected:** No formatting change.

**Pass criteria:** `endPos > 1` guard prevents zero-length bold range.

---

## Pass Criteria Summary

| TC | Scenario | Pass condition |
|---|---|---|
| TC-01 | Colon delimiter | Bold ends at `:` |
| TC-02 | Dash delimiter | Bold ends at `-` |
| TC-03 | Colon before dash | Colon wins |
| TC-04 | Dash before colon | Dash wins |
| TC-05 | No delimiter | Paragraph skipped |
| TC-06 | Non-list paragraph | Paragraph skipped |
| TC-07 | Outlook inspector | Same result as Word |
| TC-08 | No open document | MsgBox shown, clean exit |
| TC-09 | Mixed list types | Both processed correctly |
| TC-10 | Delimiter at position 1 | Paragraph skipped |
