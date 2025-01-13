---

### **README.md for `Clean_data_whatsapp.ps1` (Script-Specific)**

```markdown
# Clean_data.ps1

## Overview
The `Clean_data.ps1` script is designed to clean and process text files by:
- Removing extra quotes from file paths.
- Validating the existence of input and output files.
- Extracting and correcting lines matching specific patterns, such as dates enclosed in square brackets.

## Usage
1. Run the script in a PowerShell terminal:
   ```powershell
   .\Clean_data.ps1
   ```
2. Follow the prompts:
   - Provide the full path to the input file.
   - Provide the full path to the output file.

## Example
**Input File:**
```
"[01/01/2023, 12:00:00 PM] Some text here"
"Invalid data here"
```

**Output File:**
```
"[01/01/2023, 12:00:00 PM] Some text here"
```

## Contributing
If you have ideas to improve this script or find issues, refer to the [CONTRIBUTING.md](../CONTRIBUTING.md).

---


