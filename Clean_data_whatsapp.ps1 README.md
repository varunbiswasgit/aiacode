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

---

## Usage Instructions
### 1. Prepare Your Input File
- If using WhatsApp, export your chat as a `.txt` file.
- For generic use cases, ensure your file contains:
  - Lines starting with timestamps.
  - Messages structured with participants’ names followed by text.

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
The following enhancements are planned to  make this script even more robust and user-friendly:

### 1. **Error Handling**
   - [ ] Add more descriptive error messages to help users improve and resolve issues.
   - [ ] Implement exception handling for unexpected scenarios, such as file permission errors or incorrect input formats.

### 2. **Parameterization**
   - [ ] Allow input and output file paths to be passed as command-line parameters to improve flexibility and enable pipeline automation.

### 3. **Logging**
   - [ ] Introduce logging capabilities to track script execution and facilitate troubleshooting.

### 4. **Code Comments**
   - [ ] Enhance code documentation by adding comments explaining each section's purpose for better readability.

### 5. **Performance Optimization**
   - [ ] Optimize regular expressions and loops for faster execution, especially for large input files.

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
