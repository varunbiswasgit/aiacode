# WordResizeBorderImagesCleanlines.bas

A Word VBA macro that resizes inline images and cleans blank paragraphs in a single pass. Designed for documents exported from SAP or other systems that produce oversized images and excessive blank lines.

## Features

### Feature 1 — Image Resize and Border

- Prompts for minimum width, maximum width (inches), and border width (points)
- Accepts border colour as RGB components or a hex code (e.g. `#4472C6`)
- Resizes all inline images wider than the minimum to the specified maximum, preserving aspect ratio
- Applies a solid border of the specified width and colour

### Feature 2 — Blank Paragraph Cleanup

**Step 0 — Remove ghost bullets/numbered items:** Paragraphs that are blank (including Unicode invisible characters) and carry list formatting are deleted outright.

**Step 1 — Collapse consecutive blank lines to one:** Runs of two or more blank paragraphs are reduced to a single blank. Handles SAP/Unicode invisible characters that `Trim()` and the wildcard engine miss:

| Character | Code point | Name |
|-----------|-----------|------|
| Non-breaking space | `Chr(160)` | Common in SAP ALV exports |
| Zero-width space | `ChrW(8203)` | |
| Zero-width non-joiner | `ChrW(8204)` | |
| Zero-width joiner | `ChrW(8205)` | |
| LTR / RTL mark | `ChrW(8206/8207)` | |
| Narrow no-break space | `ChrW(8239)` | SAP EU locale |
| Medium math space | `ChrW(8287)` | |
| Soft hyphen | `ChrW(173)` | |
| BOM | `ChrW(65279)` | |

## Requirements

- Microsoft Word (any version supporting VBA)
- Macro execution must be enabled
- Run on the active document

## Installation

1. Press **Alt + F11** in Word.
2. Right-click your document → **Import File**.
3. Select `WordResizeBorderImagesCleanlines.bas` from this folder.
4. Run `ResizeImagesAndCleanDocument` from **View → Macros**.

## License

See [LICENSE](../LICENSE) in the repository root.
