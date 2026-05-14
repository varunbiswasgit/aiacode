# Bold List Prefixes — Word & Outlook

Bolds the **prefix** of every bulleted or numbered list item in the active document, up to and including the first colon (`:`) or dash (`-`), whichever appears first. Works identically in **Microsoft Word** and the **Outlook message editor**.

## Files

| File | Purpose |
|------|---------|
| `BoldListPrefixesWordOutlook.bas` | Main VBA macro |
| `Test_BoldListPrefixes.bas` | Automated VBA unit test harness |

## Compatibility

| Host application | Supported |
|---|---|
| Microsoft Word (any version with VBA) | Yes |
| Microsoft Outlook — compose/reply inspector | Yes |

## Logic Flow

1. Resolve the active document — try `Application.ActiveDocument` first; fall back to `Application.ActiveInspector.WordEditor` for Outlook.
2. If no editable document is found, show a `MsgBox` and exit cleanly.
3. Loop over every `Paragraph` in the document.
4. Skip paragraphs where `ListFormat.ListType = 0` (not a list item).
5. For each list paragraph, locate the first `:` and first `-` in the paragraph text.
6. Select the earlier delimiter; if neither is present, skip the paragraph.
7. Duplicate the paragraph range, trim its end to `Start + endPos - 1`, and apply `Font.Bold = True`.

## Configuration

This macro has no external configuration. All behaviour is controlled by the paragraph content at run time.

| Implicit setting | Value | Notes |
|---|---|---|
| Delimiter priority | Earliest of `:` or `-` | Whichever appears first in the paragraph text |
| List types processed | All (`ListType <> 0`) | Bullet, numbered, and outline levels |
| Minimum prefix length | 2 characters | `endPos > 1` guard prevents single-char false matches |

## Usage

1. Open the Word document or the Outlook compose/reply window.
2. Open the VBA editor (`Alt+F11`).
3. Import or paste `BoldListPrefixesWordOutlook.bas` into any standard module.
4. Run `BoldListPrefixesWordOutlook` (`F5` or `Alt+F8 → Run`).

## Error Handling

| Condition | Behaviour |
|---|---|
| No active document or inspector | `MsgBox` prompt; macro exits |
| List item with no `:` or `-` | Paragraph skipped silently |
| `endPos = 1` (delimiter is the first character) | Paragraph skipped (guard: `endPos > 1`) |

## Version History

| Version | Summary |
|---|---|
| v1 | Initial release — bold prefix up to first `:` or `-` in list items; Word + Outlook support |
| v1.1 | Renamed from BoldListPrefixes to BoldListPrefixesOutlook to clarify Outlook context |
| v1.2 | Renamed to BoldListPrefixesWordOutlook; folder renamed to bold-list-prefixes-word-outlook to reflect Word + Outlook scope |

## License

See [LICENSE](../LICENSE) in the repository root.
