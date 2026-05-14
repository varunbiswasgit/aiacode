# Clean WhatsApp Chat Data — Testing Guide

---

## Automated Tests (Pester)

**File:** `Clean_data_whatsapp.Tests.ps1`

Run from the `clean-data-whatsapp/` folder:

```powershell
Invoke-Pester .\Clean_data_whatsapp.Tests.ps1 -Output Detailed
```

Requires Pester v5+:
```powershell
Install-Module Pester -Force -SkipPublisherCheck
```

### Automated Test Coverage

| Test | Covers |
|------|--------|
| Parses AM/PM and 24-hour timestamps | TC-PS-04 |
| Strips surrounding quotes from file paths | TC-PS-09 |
| Removes Unicode LRM / non-standard spaces | TC-PS-08 |
| Preserves colons inside message body | TC-PS-07 |
| Exits gracefully for missing input file | TC-PS-10 |
| Creates empty output for empty input file | TC-PS-11 |
| Merges multi-line continuation into Column D | TC-PS-03 |
| Parses DD/MM/YYYY format → M/D/YY | TC-PS-04 |
| Parses ISO 8601 YYYY-MM-DD → M/D/YY | TC-PS-04 |
| CSV output — plain field | TC-PS-02 |
| CSV output — field with comma quoted | TC-PS-02 |
| Sender filter returns only matching rows | TC-PS-05 |
| Date range filter returns rows in range | TC-PS-06 |

---

## Manual Test Cases

Run from the `clean-data-whatsapp/` folder:
```powershell
.\Clean_data_whatsapp.ps1
```

### TC-PS-01 · Standard export — TSV output

| Field | Detail |
|-------|--------|
| Setup | Text file with 3 messages in WhatsApp default format |
| Input | Format → `T` (TSV) |
| Expected | Three tab-separated rows: Date, Time, Sender, Message |
| Pass criteria | Opens correctly in Excel via Data → From Text |

### TC-PS-06 · Date range filter

| Field | Detail |
|-------|--------|
| Setup | Five messages dated 1/1/23 through 1/5/23 (one per day) |
| Input | FROM `01/02/2023` TO `01/04/2023`. TSV output. |
| Expected | Three rows returned (2nd, 3rd, 4th only) |
| Pass criteria | Output contains exactly three data rows |

### TC-PS-12 · Processing summary report printed

| Field | Detail |
|-------|--------|
| Setup | Five messages from two senders over three days |
| Input | TSV output, no filters |
| Expected | Terminal prints: total messages, unique senders, sender list, date range |
| Pass criteria | All four summary fields visible in terminal output |
