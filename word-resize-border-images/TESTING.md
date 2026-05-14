# WordResizeBorderImagesCleanlines.bas — Testing

All testing is manual — no automated VBA test harness exists. The macro is fully UI-driven (multiple InputBox prompts).

## Environment Setup

1. Open a Word document containing inline images and/or blank paragraphs.
2. Press **Alt + F11**, import `WordResizeBorderImagesCleanlines.bas` from this folder.
3. Run `ResizeImagesAndCleanDocument` from **Developer → Macros**.
4. Respond to each InputBox prompt as specified in the test case.

---

## Input Validation — Minimum Width

### TC-WR-01 · Valid minimum width accepted

| Field | Detail |
|-------|--------|
| Input | `3` (inches) |
| Expected | Accepted; macro proceeds to next prompt |

### TC-WR-02 · Non-numeric minimum width rejected

| Field | Detail |
|-------|--------|
| Input | `abc` |
| Expected | MsgBox: **"Please enter a valid positive number for Minimum Width."** Loop re-prompts. |

### TC-WR-03 · Zero / negative minimum width rejected

| Field | Detail |
|-------|--------|
| Input | `0` or `-1` |
| Expected | Same re-prompt as TC-WR-02. |

### TC-WR-04 · Cancel exits macro cleanly

| Field | Detail |
|-------|--------|
| Input | Press **Cancel** on any InputBox |
| Expected | Macro exits immediately. No changes to document. |

---

## Input Validation — Maximum Width

### TC-WR-05 · Maximum width must exceed minimum width

| Field | Detail |
|-------|--------|
| Input | Min = `3`, Max = `2` |
| Expected | MsgBox: **"Maximum Width must be a valid number greater than Minimum Width."** Loop re-prompts. |

### TC-WR-06 · Valid maximum width accepted

| Field | Detail |
|-------|--------|
| Input | Min = `3`, Max = `6.3` |
| Expected | Accepted; macro proceeds to border width prompt. |

---

## Input Validation — Border Width

### TC-WR-07 · Non-numeric border width rejected

| Field | Detail |
|-------|--------|
| Input | `xyz` |
| Expected | MsgBox: **"Please enter a valid positive number for Border Width."** Re-prompts. |

### TC-WR-08 · Valid border width accepted

| Field | Detail |
|-------|--------|
| Input | `1.2` |
| Expected | Accepted; macro proceeds to color mode prompt. |

---

## Color Input — RGB Mode

### TC-WR-09 · Valid RGB values accepted

| Field | Detail |
|-------|--------|
| Input | Color mode = `1`; R = `68`, G = `114`, B = `198` |
| Expected | Border color set to RGB(68, 114, 198). Macro proceeds to image processing. |

### TC-WR-10 · Out-of-range RGB component rejected

| Field | Detail |
|-------|--------|
| Input | R = `300` |
| Expected | MsgBox: **"Enter 0-255 for Red."** Re-prompts for red component only. |

---

## Color Input — Hex Mode

### TC-WR-11 · Valid hex code accepted (with `#` prefix)

| Field | Detail |
|-------|--------|
| Input | Color mode = `2`; hex = `#4472C6` |
| Expected | Accepted; border color applied. |

### TC-WR-12 · Valid hex code accepted (without `#` prefix)

| Field | Detail |
|-------|--------|
| Input | `4472C6` |
| Expected | Accepted; border color applied. |

### TC-WR-13 · Invalid hex code rejected

| Field | Detail |
|-------|--------|
| Input | `ZZZZZZ` or `44C6` (wrong length) |
| Expected | MsgBox: **"Enter a valid 6-character hex code."** Re-prompts. |

---

## Feature 1 — Image Resize and Border

### TC-WR-14 · Image wider than minimum width resized

| Field | Detail |
|-------|--------|
| Setup | Document with one inline image wider than the set minimum width |
| Input | Min = `3`, Max = `6.3`, border = `1.2`, color = RGB(68,114,198) |
| Expected | Image width = `6.3` inches; aspect ratio maintained; border applied with specified weight and color |
| Pass criteria | Right-click image → Format Picture → Size and Border confirm values. |

### TC-WR-15 · Image narrower than minimum width not resized

| Field | Detail |
|-------|--------|
| Setup | Small inline image (e.g., 1 inch wide) |
| Expected | Image remains unchanged. No border applied. |

### TC-WR-16 · Multiple images — each processed independently

| Field | Detail |
|-------|--------|
| Setup | Document with 3 images: 1 narrow (untouched), 2 wide (resized) |
| Expected | Only the 2 wide images are resized and bordered. |

---

## Feature 2 — Blank Paragraph Cleanup

### TC-WR-17 · Ghost bullet blank paragraphs removed (Step 0)

| Field | Detail |
|-------|--------|
| Setup | Document with blank paragraphs that have list formatting (bullets/numbering) applied |
| Expected | All such paragraphs deleted before the consecutive-blank collapse step |
| Pass criteria | No blank bulleted paragraphs remain. |

### TC-WR-18 · Consecutive blank paragraphs collapsed to one (Step 1)

| Field | Detail |
|-------|--------|
| Setup | Three consecutive blank paragraphs between content blocks |
| Expected | Reduced to exactly one blank paragraph |
| Pass criteria | Show formatting marks (¶); count blank lines between content = 1. |

### TC-WR-19 · Single blank paragraph preserved

| Field | Detail |
|-------|--------|
| Setup | One blank paragraph between two content blocks |
| Expected | Blank paragraph retained unchanged |
| Pass criteria | Paragraph count unchanged after macro. |

### TC-WR-20 · Unicode invisible characters treated as blank

| Field | Detail |
|-------|--------|
| Setup | Paragraph containing only a non-breaking space (Chr(160)) or zero-width space (U+200B) |
| Expected | Treated as blank; subject to consecutive-blank collapse rule |
| Pass criteria | After macro, paragraph is either removed or reduced to a single blank. |

### TC-WR-21 · Completion message shown

| Field | Detail |
|-------|--------|
| Expected | MsgBox: **"Images resized and document cleaned."** at end of macro |
| Pass criteria | Message appears with `vbInformation` icon. |

---

## Helper Functions

### TC-WR-22 · `HexToRGB` — correct conversion

| Field | Detail |
|-------|--------|
| Input | `#4472C6` |
| Expected | Returns `RGB(68, 114, 198)` = `13,023,430` |

### TC-WR-23 · `HexToRGB` — invalid input returns -1

| Field | Detail |
|-------|--------|
| Input | `#GGG` or 4-character hex |
| Expected | Returns `-1` |

### TC-WR-24 · `IsEffectivelyBlank` — all Unicode whitespace variants detected

| Field | Detail |
|-------|--------|
| Input | Strings containing Chr(160), U+200B, U+200C, U+200D, U+200E, U+200F, U+202F, U+205F, Chr(173), U+FEFF individually |
| Expected | All return `True` |
