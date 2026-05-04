# Clean WhatsApp Chat Export Script

## Overview
The `Clean_data_whatsapp.ps1` script processes WhatsApp chat exports saved as text files. It cleans and structures unorganised text data, enabling users to easily import it into tools like Excel for advanced querying and table creation.

While optimised for WhatsApp chats, this script also supports generic use cases such as cleaning logs or any text file that follows a similar timestamp-based format.

---

## Key Features
1. **Date and Time Detection**
   - Matches messages based on timestamps using a regular expression.
   - Supports multiple timestamp formats: `M/D/YY`, `DD/MM/YYYY`, and `YYYY-MM-DD` (ISO 8601).
   - All formats are normalised to a consistent `M/D/YY` output.

2. **Unicode Character Handling**
   - Removes problematic characters such as `Left-to-Right Mark (LRM)` and unusual spaces for better readability.

3. **Multi-line Message Support**
   - Combines broken multi-line messages into a single line, maintaining proper structure.

4. **Flexible Output Format**
   - Choose between **tab-separated (TSV)** for direct Excel import or **CSV** for broader tool compatibility.
   - CSV fields containing commas are automatically quoted per RFC 4180.

5. **Colon Handling**
   - Correctly parses messages even when the message text itself contains colons.

6. **Path Quote Stripping**
   - Automatically removes surrounding quotes from file paths, supporting drag-and-drop input in PowerShell terminals.

7. **Sender Filter**
   - Optionally filter output to messages from a specific sender (partial name match supported).

8. **Date Range Filter**
   - Optionally filter output to messages within a specified date range (`MM/DD/YYYY` format).

9. **Processing Summary Report**
   - Prints a summary after processing: total messages, unique senders, sender list, and date range covered.

---

## Usage Instructions

### 1. Prepare Your Input File
- Export your WhatsApp chat as a `.txt` file.
- For generic use cases, ensure lines start with a supported timestamp and follow the `[timestamp] Name: Message` structure.

### 2. Run the Script
1. Open a PowerShell terminal.
2. Run the script:
   ```PowerShell
   .\Clean_data_whatsapp.ps1
   ```
3. Answer each prompt:

| Prompt | Description | Example |
|---|---|---|
| Input file path | Full path to the `.txt` export | `C:\chats\export.txt` |
| Output file path | Full path for the cleaned output | `C:\chats\output.txt` |
| Output format | `T` for tab-separated, `C` for CSV | `T` |
| Sender filter | Partial name to filter by, or blank for all | `John` or *(blank)* |
| From date | Start of date range (`MM/DD/YYYY`), or blank | `01/01/2023` or *(blank)* |
| To date | End of date range (`MM/DD/YYYY`), or blank | `06/30/2023` or *(blank)* |

### 3. Review the Summary
After processing, the script prints a summary to the terminal:
```
========================================
           PROCESSING SUMMARY
========================================
Output format   : TSV
Output file     : C:\chats\output.txt
Total messages  : 142
Unique senders  : 3
Senders         : Alice, Bob, System
Date range      : 01/01/2023  ->  06/30/2023
========================================
```

### 4. Analyse the Output
Open the output file in Excel or another tool. Columns are:
1. **Date** — normalised to `M/D/YY`
2. **Time** — as exported, including AM/PM where present
3. **Name** — sender name
4. **Message** — full message text, including any continuation lines

---

## Supported Timestamp Formats

| Format | Example | Notes |
|---|---|---|
| `M/D/YY` | `[1/12/23, 9:45:00 PM]` | WhatsApp default (US) |
| `M/D/YYYY` | `[1/12/2023, 9:45:00 PM]` | WhatsApp with 4-digit year |
| `DD/MM/YYYY` | `[25/12/2023, 10:00:00]` | European / international |
| `YYYY-MM-DD` | `[2023-06-15, 14:30:00]` | ISO 8601 |

All formats support both 12-hour (`AM`/`PM`) and 24-hour time.

---

## Example

### Input File
```
[1/12/23, 9:45:00 PM] John Doe: Hello!
How are you doing?
[25/12/2023, 10:00:00 AM] Jane Smith: Merry Christmas!
[2023-06-15, 14:30:00] System: Server restarted
```

### Output (Tab-Separated)
```
1/12/23    9:45:00 PM   John Doe     Hello! How are you doing?
12/25/23   10:00:00 AM  Jane Smith   Merry Christmas!
6/15/23    14:30:00     System       Server restarted
```

---

## Testing

The script is covered by a [Pester](https://pester.dev/) test suite located at `tests/Clean_data_whatsapp.Tests.ps1`.

### Test Cases
| # | Test | What It Validates |
|---|---|---|
| 1 | AM/PM parsing | 12-hour and 24-hour time formats parse correctly |
| 2 | Quote stripping | `Clean-Path` removes `"` from drag-dropped file paths |
| 3 | Unicode cleanup | LRM (`\u200E`) removed; `\u202F` replaced with a space |
| 4 | Colon in message | Message body containing colons is preserved intact in Column D |
| 5 | Missing input file | Script exits gracefully; no output file created |
| 6 | Empty input file | Output file created but empty — no crash |
| 7 | Continuation lines | Multi-line messages merged with a space into Column D |
| 8 | DD/MM/YYYY format | European date format parsed and normalised to `M/D/YY` |
| 9 | YYYY-MM-DD format | ISO 8601 date format parsed and normalised to `M/D/YY` |
| 10 | CSV output | Comma-separated output produced when `C` is selected |
| 11 | CSV comma quoting | Fields containing commas are wrapped in quotes |
| 12 | Sender filter | Only messages from the specified sender are included |
| 13 | Date range filter | Only messages within the specified date range are included |

### Run the Tests
```PowerShell
# Install Pester if not already installed
Install-Module -Name Pester -Force -Scope CurrentUser

# Run all tests with detailed output
Invoke-Pester -Path ".\tests\Clean_data_whatsapp.Tests.ps1" -Output Detailed
```

---

## Generic Use Cases

### 1. Log File Processing
Clean and structure server or application log files with timestamped entries.
- **Example input:** Log files with entries like `[M/D/YY, H:MM:SS]`.

### 2. Survey or Interview Data
Organise chat-like or time-annotated responses from participants.
- **Example input:** Data exported from chat platforms or manually recorded timestamps.

### 3. Collaboration Tools Data Cleanup
Process exported conversation data from platforms like Slack, Teams, or Google Chat.

### 4. Generic Data Parsing
Handle any text file where messages are associated with timestamps and identifiers.

---

## Error Handling
- The script terminates with a clear error message if the input file does not exist.
- Continuation lines are only appended to valid existing records.
- Invalid date filter values produce a warning and are skipped gracefully.

---

## Contributing
Contributions are welcome! Please see the [CONTRIBUTING.md](../CONTRIBUTING.md) for details.

---

## License
This project is licensed under the [GNU General Public License v3.0](../LICENSE).
