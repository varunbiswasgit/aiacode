# Clean_data_whatsapp.ps1 — Testing Readme

Test coverage for `scripts/Clean_data_whatsapp.ps1`.

## Automated Tests

**File:** `tests/Clean_data_whatsapp.Tests.ps1` (Pester)

Run from the repo root:
```powershell
Invoke-Pester .\tests\Clean_data_whatsapp.Tests.ps1 -Output Detailed
```

| Pester Describe Block | Covers |
|-----------------------|--------|
| Timestamp parsing — M/D/YY | TC-PS-04 |
| Timestamp parsing — DD/MM/YYYY | TC-PS-04 |
| Timestamp parsing — ISO 8601 | TC-PS-04 |
| Multi-line merge | TC-PS-03 |
| Sender filter | TC-PS-05 |
| Unicode stripping | TC-PS-08 |
| Empty file handling | TC-PS-11 |

## Manual Test Cases

Run the script from the `scripts/` folder:
```powershell
.\Clean_data_whatsapp.ps1
```

---

### TC-PS-01 · Standard export — TSV output

| Field | Detail |
|-------|--------|
| Setup | Text file with 3 messages: `1/5/23, 10:00 - Alice: Hello`, `1/5/23, 10:01 - Bob: Hi there`, `1/5/23, 10:02 - Alice: How are you?` |
| Input | Format → `T` (TSV) |
| Expected | Three tab-separated rows with columns: Date, Time, Sender, Message |
| Pass criteria | Open in Excel via Data → From Text. Three rows render in four columns correctly. |

---

### TC-PS-02 · Standard export — CSV output

| Field | Detail |
|-------|--------|
| Setup | Same 3-message file as TC-PS-01 |
| Input | Format → `C` (CSV) |
| Expected | Three comma-separated rows. Fields with commas quoted per RFC 4180. |
| Pass criteria | Open in Excel. Three rows, four columns, no data corruption. |

---

### TC-PS-03 · Multi-line message merged into single row

| Field | Detail |
|-------|--------|
| Setup | Message spanning two lines: `1/5/23, 10:00 - Alice: Line one` followed by `continuation of line one` (no timestamp) |
| Input | TSV output |
| Expected | Single output row; Message = `Line one continuation of line one` |
| Pass criteria | Output has one data row for this message, not two. |

---

### TC-PS-04 · Multiple timestamp formats normalised

| Field | Detail |
|-------|--------|
| Setup | Three messages using formats `1/5/23`, `05/01/2023`, and `2023-01-05` |
| Input | TSV output |
| Expected | All Date values normalised to `M/D/YY` format |
| Pass criteria | All Date column values match the `M/D/YY` pattern. |

---

### TC-PS-05 · Sender filter — only matching sender returned

| Field | Detail |
|-------|--------|
| Setup | Three messages from Alice, Bob, Alice |
| Input | Sender filter → `Alice`. TSV output. |
| Expected | Two rows returned (Alice's messages only) |
| Pass criteria | Row count = 2. No rows where Sender ≠ Alice. |

---

### TC-PS-06 · Date range filter

| Field | Detail |
|-------|--------|
| Setup | Five messages dated 1/1/23 through 1/5/23 (one per day) |
| Input | FROM `01/02/2023` TO `01/04/2023`. TSV output. |
| Expected | Three rows returned (2nd, 3rd, 4th) |
| Pass criteria | Output contains exactly three data rows within range. |

---

### TC-PS-07 · Message containing colons not mis-split

| Field | Detail |
|-------|--------|
| Setup | Message: `1/5/23, 10:00 - Alice: Time is 10:30:00 today` |
| Input | TSV output |
| Expected | Sender = `Alice`. Message = `Time is 10:30:00 today` |
| Pass criteria | Message column contains full string with colons intact. |

---

### TC-PS-08 · Unicode LRM characters stripped

| Field | Detail |
|-------|--------|
| Setup | Input file contains Left-to-Right Mark (U+200F) in sender names and message text |
| Input | TSV output |
| Expected | No LRM characters in output |
| Pass criteria | `Select-String -Pattern '\u200F' output.txt` returns no matches. |

---

### TC-PS-09 · Path with surrounding quotes stripped

| Field | Detail |
|-------|--------|
| Setup | Valid input at `C:\temp\chat.txt` |
| Input | Path entered as `"C:\temp\chat.txt"` (with double-quotes, simulating drag-and-drop) |
| Expected | Script accepts path and processes normally |
| Pass criteria | Output file produced. No "file not found" error. |

---

### TC-PS-10 · Non-existent input file

| Field | Detail |
|-------|--------|
| Setup | No file at provided path |
| Input | Path → `C:\temp\nonexistent.txt` |
| Expected | Error message printed; script exits without output |
| Pass criteria | No output file created. Error references the missing path. |

---

### TC-PS-11 · Empty input file

| Field | Detail |
|-------|--------|
| Setup | Input file exists but contains zero bytes |
| Input | TSV output |
| Expected | Warning printed (no messages found). Output empty or not created. |
| Pass criteria | No data rows in output. No unhandled exception. |

---

### TC-PS-12 · Processing summary report printed

| Field | Detail |
|-------|--------|
| Setup | Five messages from two senders over three days |
| Input | TSV output, no filters |
| Expected | Terminal prints: total messages, unique senders, sender list, date range |
| Pass criteria | All four summary fields visible in terminal output. |
