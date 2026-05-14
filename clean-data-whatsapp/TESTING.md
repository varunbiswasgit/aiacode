# Clean-Data-WhatsApp — Testing Readme

Test coverage for `clean-data-whatsapp/Clean_data_whatsapp.ps1`.

## Automated Tests

No automated test harness exists. The script uses interactive `Read-Host` prompts that require manual input.

## Manual Test Cases

Environment setup:

1. Export a WhatsApp chat via **Chat > Export chat > Without media**.
2. Open PowerShell and navigate to the `clean-data-whatsapp/` folder.
3. Run `.\Clean_data_whatsapp.ps1` and supply the prompted values.

---

### TC-01 · Standard WhatsApp export (M/D/YY format)

| Step | Action |
|------|--------|
| Input | WhatsApp export with timestamps like `[3/15/24, 10:32:00 AM]` |
| Format | TSV |
| Expected | Each message parsed into Date / Time / Sender / Message columns. Summary printed. |

### TC-02 · European date format (DD/MM/YYYY)

| Step | Action |
|------|--------|
| Input | Export with timestamps like `[15/03/2024, 10:32:00]` |
| Expected | Dates normalised to `M/D/YY`. Output rows match correct calendar dates. |

### TC-03 · ISO 8601 date format (YYYY-MM-DD)

| Step | Action |
|------|--------|
| Input | Export with timestamps like `[2024-03-15, 10:32:00]` |
| Expected | Dates normalised to `M/D/YY`. Output rows match correct calendar dates. |

### TC-04 · CSV output format

| Step | Action |
|------|--------|
| Format choice | `C` |
| Expected | Fields containing commas or quotes are correctly quoted per RFC 4180. |

### TC-05 · Sender filter

| Step | Action |
|------|--------|
| Sender filter | Partial name, e.g. `Alice` |
| Expected | Output contains only messages where Sender matches `*Alice*`. Summary shows filtered count. |

### TC-06 · Date range filter

| Step | Action |
|------|--------|
| FROM | `01/01/2024` |
| TO | `01/31/2024` |
| Expected | Only messages from January 2024 appear in output. |

### TC-07 · Multi-line message continuations

| Step | Action |
|------|--------|
| Input | A message that spans multiple lines in the export file |
| Expected | Continuation lines are appended to the previous record’s Message field with a space separator. |

### TC-08 · Input file not found

| Step | Action |
|------|--------|
| Input path | Non-existent file path |
| Expected | Error message: `"Error: The input file '...' does not exist. Exiting..."` Script exits. |

### TC-09 · Invalid FROM date

| Step | Action |
|------|--------|
| FROM | `not-a-date` |
| Expected | Warning: `"Could not parse FROM date 'not-a-date'. Date filter ignored."` Script continues without date filter. |

### TC-10 · Blank filters (no filtering)

| Step | Action |
|------|--------|
| Sender filter | *(blank)* |
| FROM / TO | *(blank)* |
| Expected | All parsed messages written to output. Summary shows full count. |

### TC-11 · Unicode whitespace normalisation

| Step | Action |
|------|--------|
| Input | Messages containing `\u202F` (narrow no-break space) or `\u00A0` (non-breaking space) |
| Expected | These characters replaced with a standard space in the output. |

### TC-12 · Empty input file

| Step | Action |
|------|--------|
| Input | A zero-byte or whitespace-only `.txt` file |
| Expected | Output file created but empty. Summary shows `Total messages: 0`. |
