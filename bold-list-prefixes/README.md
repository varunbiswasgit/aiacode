# BoldListPrefixesOutlook

A VBA macro that bolds the prefix of every bulleted or numbered list item in the active document, up to and including the first colon (`:`) or dash (`-`), whichever appears first. Works identically in **Microsoft Word** and the **Outlook message editor**.

## Compatibility

| Host application | Supported |
|---|---|
| Microsoft Word (any version with VBA) | Yes |
| Microsoft Outlook — compose/reply inspector | Yes |

## How It Works

1. Resolves the active document — tries `Application.ActiveDocument` first; falls back to `Application.ActiveInspector.WordEditor` for Outlook.
2. If no editable document is found, shows a `MsgBox` and exits cleanly.
3. Loops over every `Paragraph` in the document.
4. Skips paragraphs where `ListFormat.ListType = 0` (not a list item).
5. For each list paragraph, locates the first `:` and first `-` in the paragraph text.
6. Selects the earlier delimiter; skips the paragraph if neither is present.
7. Duplicates the paragraph range, trims its end to `Start + endPos - 1`, and applies `Font.Bold = True`.

## Configuration

No external configuration. All behaviour is driven by paragraph content at run time.

| Implicit setting | Value | Notes |
|---|---|---|
| Delimiter priority | Earliest of `:` or `-` | Whichever appears first in the paragraph text |
| List types processed | All (`ListType <> 0`) | Bullet, numbered, and outline levels |
| Minimum prefix length | 2 characters | `endPos > 1` guard prevents single-char false matches |

## Installation

1. Open the Word document or the Outlook compose/reply window.
2. Open the VBA editor (`Alt+F11`).
3. Import `bold-list-prefixes/BoldListPrefixesOutlook.bas` into any standard module.
4. Run `BoldListPrefixesOutlook` via `Alt+F8 → Run`.

## Error Handling

| Condition | Behaviour |
|---|---|
| No active document or inspector | `MsgBox` prompt; macro exits |
| List item with no `:` or `-` | Paragraph skipped silently |
| `endPos = 1` (delimiter is first character) | Paragraph skipped (guard: `endPos > 1`) |

## Version History

| Version | Summary |
|---|---|
| v1 | Initial release — bold prefix up to first `:` or `-` in list items; Word + Outlook support |
| v1.1 | Renamed from `BoldListPrefixes` to `BoldListPrefixesOutlook` to clarify Outlook context |

## License

See [LICENSE](../LICENSE) in the repository root.
