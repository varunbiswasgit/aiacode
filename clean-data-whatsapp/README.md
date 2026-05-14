# Clean WhatsApp Chat Data

A PowerShell script that parses a WhatsApp exported chat text file and produces a clean, structured TSV or CSV file with columns: **Date**, **Time**, **Sender**, **Message**.

## Features

- Supports three timestamp formats: `M/D/YY` (WhatsApp default), `DD/MM/YYYY` (European), `YYYY-MM-DD` (ISO 8601)
- Outputs TSV (tab-separated) or CSV (comma-separated)
- Optional sender filter (substring match)
- Optional date range filter (FROM / TO, `MM/DD/YYYY`)
- Handles multi-line messages — continuation lines appended to the previous message
- Strips Unicode invisible characters common in WhatsApp exports (`\u200E`, `\u202F`, `\u00A0`, `\u2007`)
- Prints a summary report: total messages, unique senders, date range

## Usage

Run from PowerShell:

```powershell
.\Clean_data_whatsapp.ps1
```

The script prompts interactively for all inputs — no command-line arguments required.

## Prompts

| Prompt | Example input |
|--------|---------------|
| Input file path | `C:\chats\export.txt` |
| Output file path | `C:\chats\clean.tsv` |
| Output format (T/C) | `T` for TSV, `C` for CSV |
| Sender filter | `Alice` (leave blank for all) |
| FROM date | `01/01/2024` (leave blank for none) |
| TO date | `12/31/2024` (leave blank for none) |

## Requirements

- PowerShell 5.1 or later (Windows built-in)
- No external modules required

## Running Tests

```powershell
Invoke-Pester .\Clean_data_whatsapp.Tests.ps1 -Output Detailed
```

Requires [Pester](https://pester.dev/) v5+:
```powershell
Install-Module Pester -Force -SkipPublisherCheck
```

## License

See [LICENSE](../LICENSE) in the repository root.
