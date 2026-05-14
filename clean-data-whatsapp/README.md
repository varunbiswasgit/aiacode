# Clean-Data-WhatsApp

A PowerShell script that parses a WhatsApp exported chat text file into a structured, filterable tabular format (TSV or CSV), with optional filtering by sender and date range.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- A WhatsApp chat export in `.txt` format (exported via **Chat > Export chat > Without media**)

## Supported Timestamp Formats

| Format | Example |
|---|---|
| WhatsApp default (M/D/YY) | `[3/15/24, 10:32:00 AM]` |
| WhatsApp default (M/D/YYYY) | `[3/15/2024, 10:32:00 AM]` |
| European (DD/MM/YYYY) | `[15/03/2024, 10:32:00]` |
| ISO 8601 (YYYY-MM-DD) | `[2024-03-15, 10:32:00]` |

All formats are normalised internally to `M/D/YY` before processing.

## Usage

Run from PowerShell:

```powershell
.\Clean_data_whatsapp.ps1
```

The script will prompt for:

1. **Input file path** — full path to the exported WhatsApp `.txt` file
2. **Output file path** — full path for the cleaned output file
3. **Output format** — `T` for TSV (default) or `C` for CSV
4. **Sender filter** — partial name match; leave blank to include all senders
5. **Date FROM** — optional start date in `MM/DD/YYYY` format
6. **Date TO** — optional end date in `MM/DD/YYYY` format

## Output Columns

| Column | Description |
|---|---|
| Date | Message date, normalised to `M/D/YY` |
| Time | Message time as exported |
| Sender | Display name of the sender |
| Message | Full message text (multi-line messages are joined into a single field) |

## Output Summary

After writing the output file, the script prints a processing summary:

```
========================================
           PROCESSING SUMMARY
========================================
Output format   : TSV
Output file     : C:\output\chat.tsv
Total messages  : 342
Unique senders  : 4
Senders         : Alice, Bob, Carol, Dave
Date range      : 01/01/2024  ->  03/15/2024
========================================
```

## Notes

- Unicode whitespace characters (`\u202F`, `\u00A0`, `\u2007`) are normalised to a standard space.
- Left-to-right marks (`\u200E`) are stripped.
- Multi-line messages (continuation lines without a timestamp) are appended to the previous record.
- System messages (e.g. “Messages and calls are end-to-end encrypted”) are included as records where the sender field reflects the WhatsApp system string; filter them out by sender if needed.

## License

See [LICENSE](../LICENSE) in the repository root.
