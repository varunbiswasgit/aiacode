# Clean_data_whatsapp.ps1 — Manual Test Cases

All tests supplement the automated Pester suite (`Clean_data_whatsapp.Tests.ps1`).  
Run each test in a PowerShell terminal from the `scripts/` folder:
```powershell
.\Clean_data_whatsapp.ps1
```

---

## Environment Setup

1. PowerShell 5.1 or later.
2. Sample `.txt` files as described per test case — create them in a temp folder.
3. Run `.\Clean_data_whatsapp.ps1` and enter paths at the prompts.

---

## TC-PS-01 · Standard WhatsApp export — TSV output

| Field | Detail |
|-------|--------|
| Setup | Text file with 3 messages: `1/5/23, 10:00 - Alice: Hello`, `1/5/23, 10:01 - Bob: Hi there`, `1/5/23, 10:02 - Alice: How are you?` |
| Input | Input path → file above. Output path → `output.txt`. Format → `T` (TSV). |
| Expected | Three tab-separated rows, columns: Date, Time, Sender, Message. No extra blank lines. |
| Pass criteria | Open output in Excel (Data → From Text). Three rows render correctly in four columns. |

---

## TC-PS-02 · Standard WhatsApp export — CSV output

| Field | Detail |
|-------|--------|
| Setup | Same 3-message file as TC-PS-01. |
| Input | Format → `C` (CSV). |
| Expected | Three comma-separated rows. Fields containing commas are quoted per RFC 4180. |
| Pass criteria | Open output in Excel. Three rows, four columns, no data corruption. |

---

## TC-PS-03 · Multi-line message merged into single row

| Field | Detail |
|-------|--------|
| Setup | Message spanning two lines: `1/5/23, 10:00 - Alice: Line one` followed by `continuation of line one` (no timestamp). |
| Input | TSV output. |
| Expected | Single output row for Alice; Message field contains `Line one continuation of line one`. |
| Pass criteria | Output file has one data row for this message, not two. |

---

## TC-PS-04 · Multiple timestamp formats normalised

| Field | Detail |
|-------|--------|
| Setup | Three messages using formats `1/5/23`, `05/01/2023`, and `2023-01-05` respectively. |
| Input | TSV output. |
| Expected | All three Date values normalised to `1/5/23` format in output. |
| Pass criteria | All Date column values match the `M/D/YY` pattern. |

---

## TC-PS-05 · Sender filter — only matching sender returned

| Field | Detail |
|-------|--------|
| Setup | Three messages from Alice, Bob, Alice. |
| Input | Sender filter → `Alice`. TSV output. |
| Expected | Only two rows returned (Alice's messages). Bob's message absent. |
| Pass criteria | Output row count = 2. No rows where Sender ≠ Alice. |

---

## TC-PS-06 · Date range filter

| Field | Detail |
|-------|--------|
| Setup | Five messages dated 1/1/23 through 1/5/23 (one per day). |
| Input | Date range → `01/02/2023` to `01/04/2023`. TSV output. |
| Expected | Three rows returned (2nd, 3rd, 4th). |
| Pass criteria | Output contains exactly three data rows within the specified range. |

---

## TC-PS-07 · Message containing colons not mis-split

| Field | Detail |
|-------|--------|
| Setup | Message: `1/5/23, 10:00 - Alice: Time is 10:30:00 today`. |
| Input | TSV output. |
| Expected | Sender = `Alice`. Message = `Time is 10:30:00 today` (colons in message text preserved). |
| Pass criteria | Message column contains the full string with colons intact. Sender is not truncated. |

---

## TC-PS-08 · Unicode LRM characters stripped

| Field | Detail |
|-------|--------|
| Setup | Input file contains Left-to-Right Mark (U+200F) embedded in sender names and message text. |
| Input | TSV output. |
| Expected | Output file contains no LRM characters. Sender names and messages are clean ASCII/UTF-8. |
| Pass criteria | `Select-String -Pattern '\u200F' output.txt` returns no matches. |

---

## TC-PS-09 · Path with surrounding quotes stripped

| Field | Detail |
|-------|--------|
| Setup | Valid input file at `C:\temp\chat.txt`. |
| Input | Enter path as `"C:\temp\chat.txt"` (with surrounding double-quotes, simulating drag-and-drop). |
| Expected | Script accepts the path without error and processes the file normally. |
| Pass criteria | Output file produced. No "file not found" error. |

---

## TC-PS-10 · Non-existent input file

| Field | Detail |
|-------|--------|
| Setup | No file at provided path. |
| Input | Input path → `C:\temp\nonexistent.txt`. |
| Expected | Script prints an error message and exits without producing output. |
| Pass criteria | No output file created. Error message references the missing path. |

---

## TC-PS-11 · Empty input file

| Field | Detail |
|-------|--------|
| Setup | Input file exists but contains zero bytes. |
| Input | TSV output. |
| Expected | Script prints a warning (no messages found). Output file is empty or not created. |
| Pass criteria | No data rows in output. No unhandled exception. |

---

## TC-PS-12 · Processing summary report printed

| Field | Detail |
|-------|--------|
| Setup | Five messages from two senders over three days. |
| Input | TSV output, no sender or date filter. |
| Expected | After processing, terminal prints: total message count, unique sender count, sender list, and date range. |
| Pass criteria | All four summary fields visible in terminal output. |

---

## Automated Coverage Reference

The following test IDs are covered by `Clean_data_whatsapp.Tests.ps1` (Pester):

| Pester Describe Block | Equivalent manual TC |
|-----------------------|----------------------|
| Timestamp parsing — M/D/YY | TC-PS-04 |
| Timestamp parsing — DD/MM/YYYY | TC-PS-04 |
| Timestamp parsing — ISO 8601 | TC-PS-04 |
| Multi-line merge | TC-PS-03 |
| Sender filter | TC-PS-05 |
| Unicode stripping | TC-PS-08 |
| Empty file handling | TC-PS-11 |

Manual test cases TC-PS-01, TC-PS-02, TC-PS-06, TC-PS-07, TC-PS-09, TC-PS-10, and TC-PS-12 cover scenarios not automated in the Pester suite.
