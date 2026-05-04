# Clean WhatsApp Chat Export Script

## Overview
The `Clean_data_whatsapp.ps1` script is designed to process WhatsApp chat exports saved as text files. It cleans and structures unorganized text data, enabling users to easily import it into tools like Excel for advanced querying and table creation.

While optimized for WhatsApp chats, this script can also be used for other generic purposes, such as cleaning up and structuring logs or text files that follow a similar timestamp-based format.

---

## Key Features
1. **Date and Time Detection:**
   - Matches messages based on timestamps using a regular expression (e.g., `[1/12/23, 9:45:00 PM]`).
   - Groups' messages with their associated participants.

2. **Unicode Character Handling:**
   - Removes problematic characters, such as `Left-to-Right Mark (LRM)` and unusual spaces, for better readability.

3. **Multi-line Message Support:**
   - Combines broken multi-line messages into a single line, maintaining proper structure.

4. **Tab-Separated Output:**
   - Saves cleaned data in a tab-separated format for easy import into Excel or other tools.

5. **Colon Handling:**
   - Correctly parses messages even when the text itself contains colons.

6. **Path Quote Stripping:**
   - Automatically removes surrounding quotes from file paths, supporting drag-and-drop input in PowerShell terminals.

---

## Usage Instructions
### 1. Prepare Your Input File
- If using WhatsApp, export your chat as a `.txt` file.
- For generic use cases, ensure your file contains:
  - Lines starting with timestamps.
  - Messages structured with participants' names followed by text.

### 2. Run the Script
1. Open a PowerShell terminal.
2. Run the script:
   ```PowerShell
   .\Clean_data_whatsapp.ps1
   ```
3. Follow the prompts:
   - **Input File:** Provide the full path of the input text file.
   - **Output File:** Provide the full path where you want to save the cleaned file.

### 3. Analyze the Output
- Open the output file in Excel or other tools.
- The data will be structured into the following columns:
  1. **Date**: The date of the message.
  2. **Time**: The time of the message.
  3. **Name**: The participant or identifier associated with the message.
  4. **Message**: The text of the message.

---

## Example
### Input File (WhatsApp Export)
```
[1/12/23, 9:45:00 PM] John Doe: Hello!
How are you doing?
[1/12/23, 9:46:30 PM] Jane Smith: I'm good, thank you!
```

### Generic Input File
```
[1/12/23, 09:00:00] System: Server started
[1/12/23, 09:10:00] User: Login attempt successful
[1/12/23, 09:15:00] User: Action performed
Another action detail is on the next line.
```

### Output File
**Tab-separated for Excel or Other Tools:**
```
Date       Time       Name        Message
1/12/23    9:45:00 PM John Doe    Hello! How are you doing?
1/12/23    9:46:30 PM Jane Smith  I'm good, thank you!
1/12/23    09:00:00  System       Server started
1/12/23    09:10:00  User         Login attempt successful
1/12/23    09:15:00  User         Action performed Another action detail on the next line.
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
| 5 | Missing input file | Script exits gracefully with an error message; no output file created |
| 6 | Empty input file | Output file is created but empty — no crash or silent failure |
| 7 | Continuation lines | Multi-line messages are merged with a space into the preceding Column D entry |

### Run the Tests
```PowerShell
# Install Pester if not already installed
Install-Module -Name Pester -Force -Scope CurrentUser

# Run all tests with detailed output
Invoke-Pester -Path ".\tests\Clean_data_whatsapp.Tests.ps1" -Output Detailed
```

---

## Generic Use Cases

### 1. **Log File Processing**
   Purpose: Clean and structure server or application log files with timestamped entries to improve readability and analysis.
   - **Example Input:** Log files with entries like `[M/D/YY, H:MM:SS]`.

### 2. **Survey or Interview Data**
   - **Purpose:** Organize chat-like or time-annotated responses from participants in a survey or interview.
   - **Example Input:** Data exported from chat platforms or manually recorded timestamps.

### 3. **Collaboration Tools Data Cleanup**
   - **Purpose:** Process exported conversation data from platforms like Slack, Teams, or Google Chat, which often have timestamps and participant names.

### 4. **Generic Data Parsing**
   - **Purpose:** Handle any text file where messages are associated with timestamps and identifiers, turning unstructured data into structured records.

---

## Future Improvements
The following enhancements are planned to make this script even more robust and user-friendly:
- [ ] Add support for additional WhatsApp timestamp formats (e.g., `DD/MM/YYYY` and `YYYY-MM-DD`).
- [ ] Add a `-Silent` switch to suppress interactive prompts and accept file paths as parameters for automation pipelines.
- [ ] Support CSV output format as an alternative to tab-separated output.
- [ ] Add an option to filter messages by sender name or date range.
- [ ] Provide a summary report (total messages, unique senders, date range) after processing.

---

## Error Handling
- The script terminates with an error message if the input file does not exist.
- Ensures continuation lines only append to valid records.

---

## Contributing
Contributions are welcome! Please see the [CONTRIBUTING.md](../CONTRIBUTING.md) for details on how to contribute.

---

## License
This project is licensed under the [GNU General Public License v3.0](../LICENSE).
