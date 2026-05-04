# Resize Images and Clean Document — VBA Macro

This VBA macro for Microsoft Word performs two tasks on the active document:

1. **Resize and Border Images**: Adjusts inline images wider than a user-specified minimum width to a user-specified maximum width, maintaining aspect ratio, and applies a solid border with a user-specified color entered as either RGB values or a Hex code.
2. **Clean Empty Lines**: Removes ghost bullet and numbered list items with no content, then collapses consecutive blank lines down to a maximum of one — including blank lines that contain invisible characters such as `Chr(160)` (non-breaking space) and SAP/Unicode zero-width characters that standard cleanup methods do not catch.

---

## Features

### Image Resizing
- Prompts for minimum and maximum width (in inches) with cancel-safe, validated input.
- Resizes only inline images that exceed the minimum width.
- Locks aspect ratio during resizing to prevent distortion.
- Processes `wdInlineShapePicture` types only; floating/anchored shapes are not affected.

### Border Application
- Prompts for border width (in points).
- Prompts for border color input method: **RGB** (three separate 0–255 values) or **Hex code** (e.g. `#4472C6` or `4472C6`).
- Applies a solid single-line border to each resized image using `.Line.Weight`, `.Line.Style`, and `.Line.ForeColor.RGB`.
- Border is applied **only to images that meet the minimum width threshold** — images below the minimum are not resized or bordered.
- Default color `#4472C6` (RGB 68, 114, 198) renders as a medium corporate blue.

### Empty Line Cleanup — Step 0: Ghost Bullet and Number Removal
- Scans every paragraph backwards through the document.
- Identifies paragraphs that are blank (including all invisible character types — see `IsEffectivelyBlank` below) **and** still carry bullet or numbered list formatting.
- Deletes these ghost list items outright as they carry no content.
- Runs before the blank line collapse so the collapse step operates on a clean document.

### Empty Line Cleanup — Step 1: Blank Line Collapse
- Replaces the previous `^13{3,}` wildcard find-and-replace approach, which could not detect impostor blank lines and caused ghost bullet creation on style-inheriting `^p^p` replacements.
- Scans every paragraph backwards and counts consecutive blank runs.
- Keeps a maximum of **one** blank paragraph per run; deletes all additional consecutive blanks.
- On the single kept blank, strips any list formatting (`.ListFormat.RemoveNumbers`) and resets the style to `Normal` — preventing ghost bullet lines from being created during the collapse.
- Handles all invisible blank character types via `IsEffectivelyBlank`.

### HexToRGB — Hex Color Conversion
A private helper function that converts a 6-character hex color string to a VBA `RGB()` Long value. Accepts input with or without a leading `#`. Returns `-1` on invalid input, which triggers a re-prompt.

| Input Example | Result |
|---|---|
| `#4472C6` | `RGB(68, 114, 198)` |
| `4472C6` | `RGB(68, 114, 198)` |
| `4472c6` | `RGB(68, 114, 198)` — case-insensitive |
| `FF0000` | `RGB(255, 0, 0)` |
| `#ZZZ000` | Invalid — re-prompt |
| `4472` | Invalid — re-prompt (wrong length) |

### IsEffectivelyBlank — Invisible Character Detection
A private helper function that strips all whitespace-equivalent characters before testing whether a paragraph is blank. Covers characters that `Trim()` and Word's `^13` wildcard engine do not detect:

| Character | Value | Description |
|---|---|---|
| Tab | `Chr(9)` | Field separator |
| Line Feed | `Chr(10)` | Unix newline |
| Vertical Tab | `Chr(11)` | Word line break (Shift+Enter) |
| Form Feed | `Chr(12)` | Page break character |
| Carriage Return | `Chr(13)` | Windows paragraph mark |
| Space | `Chr(32)` | Standard space |
| Non-Breaking Space | `Chr(160)` | Common in SAP ALV/GUI exports — **not caught by `Trim()`** |
| Zero-Width Space | `ChrW(8203)` | Embedded in SAP text fields |
| Zero-Width Non-Joiner | `ChrW(8204)` | SAP Unicode text handling |
| Zero-Width Joiner | `ChrW(8205)` | SAP Unicode text handling |
| Left-to-Right Mark | `ChrW(8206)` | SAP Arabic/Hebrew locale output |
| Right-to-Left Mark | `ChrW(8207)` | SAP Arabic/Hebrew locale output |
| Narrow No-Break Space | `ChrW(8239)` | SAP EU locale number formatting |
| Medium Mathematical Space | `ChrW(8287)` | SAP BW/BI formula fields |
| Soft Hyphen | `ChrW(173)` | Renders blank, breaks string comparisons |
| BOM / Zero-Width No-Break Space | `ChrW(65279)` | SAP CSV/Excel export file marker |

---

## How to Use

### Prerequisites
- Microsoft Word with VBA macros enabled.

### Steps

1. Open the Word document you want to modify.
2. Press `Alt + F11` to open the VBA Editor.
3. Click `Insert > Module` to create a new module.
4. Paste the macro code into the module.
5. Close the VBA Editor.
6. Press `Alt + F8`, select `ResizeImagesAndCleanDocument`, and click **Run**.

### Input Prompts

You will be prompted for the following values. Press **Cancel** on any prompt to exit the macro safely.

| Prompt | Unit | Default |
|---|---|---|
| Minimum image width | Inches | 3 |
| Maximum image width | Inches | 6.3 |
| Border width | Points | 1.2 |
| Border color mode | 1 = RGB / 2 = Hex | 1 |
| **If RGB:** Red component | 0–255 | 68 |
| **If RGB:** Green component | 0–255 | 114 |
| **If RGB:** Blue component | 0–255 | 198 |
| **If Hex:** Hex color code | `#RRGGBB` or `RRGGBB` | `#4472C6` |

All inputs are validated. Non-numeric, out-of-range, or malformed hex values will prompt a retry message.

---

## Execution Order

| Order | Step | What It Does |
|---|---|---|
| 1 | Image resize and border | Processes all inline images wider than minimum width |
| 2 | Step 0 — Ghost list removal | Deletes blank bullet and numbered paragraphs with no content |
| 3 | Step 1 — Blank line collapse | Collapses consecutive blanks to max 1, strips list style on keeper |

---

## Known Limitations

- **Inline shapes only**: The macro does not process floating images or shapes wrapped with text.
- **Cancel behavior**: Pressing Cancel on any input box exits the macro immediately without making changes.
- **Border scope**: Only images wider than the minimum width threshold receive a border. Images at or below the minimum are not resized or bordered.
- **Style reset on blank keeper**: The single blank line kept per run is always reset to `Normal` style. If an intentional blank paragraph with a custom style exists in the document, that style will be cleared.

---

## Customization

- To apply borders to **all** images regardless of size, move the `.Line` border code outside the `If .Width > minWidth Then` block.
- To skip border prompts entirely, remove the border input variables and the `.Line` assignment lines from the code.
- To change the default RGB border color, update the default values in the `InputBox` calls for Red, Green, and Blue.
- To change the default Hex border color, update the default value `"#4472C6"` in the Hex `InputBox` call.
- To allow more than one blank line between paragraphs, change `If blankCount > 1 Then` in Step 1 to `If blankCount > 2 Then` (for a maximum of two blanks).

---

## Notes

- Always save a backup of your document before running any macro.
- Ensure **Enable All Macros** or a trusted publisher setting is active in Word's Trust Center.
- This macro is particularly effective on documents exported from SAP or pasted from web sources, where non-breaking spaces and Unicode invisible characters commonly produce blank lines that resist standard cleanup.

---

## License

Licensed under the [GNU General Public License v3.0](LICENSE).
