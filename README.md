# AI Assisted Microsoft Macro, Bash, PowerShell, and Python Scripts

## Overview
This repository contains a collection of PowerShell, MS Macros, Bash, and Python scripts developed using AI tools. The scripts have been tested and worked to the best of my knowledge at the time of posting.

## Scripts Index
1. **Clean_data_Whatsapp.ps1**: A PowerShell script for cleaning and processing text files, such as removing extra quotes, validating paths, and extracting data matching specific patterns.
2. **WordResizeBorderImagesCleanlines.bas**: A VBA script that resizes images, adds a border, and cleans the document in Microsoft Word.
3. **SplitExcelByManager.bas**: A VBA macro that splits an Excel workbook into separate sheets or files, organized by manager name.
4. **split_excel_by_manager.py**: A Python script that replicates the split-by-manager logic from the VBA macro, offering a cross-platform alternative using pandas or openpyxl.
5. **ExcelFormatting.bas**: A unified VBA macro for Excel data formatting and cleanup, with four interactive modes covering simple formatting, advanced formatting, optional column splitting, and SAP output processing.
6. **NormalizeTable.bas**: A lightweight, single-pass Word VBA macro (`NormalizeTables_Light`) that standardizes all tables in the active document — normalizing widths to 100%, clearing row/column constraints, and applying Arial 10pt formatting. Designed as a fast daily-driver; use `StandardizeTables_TwoPass_AllStories` for documents with embedded images or nested tables.

---

## Features
- Modular scripts with clear functionality.
- Easy-to-use prompts and customizable logic.
- Well-documented examples and use cases (WIP).

---

## Getting Started
### 1. Clone the repository:
   ```
   git clone https://github.com/varunbiswasgit/aiacode.git
   ```
### 2. Navigate to the specific script directory:
   ```
   cd aiacode
   ```
### 3. Follow instructions in each script's individual README file.

---

## Script-Specific README Files
Each script has a dedicated README file named `<filename.extension README.md>`, e.g., `Clean_data_whatsapp.ps1 README.md`. These files contain usage instructions, configuration options, known limitations, and task lists for tracking planned improvements.

### Task List Syntax
To add a task list in a script README, use:
```
- [ ] Task description
- [x] Completed task description
```

Task lists can also link directly to GitHub Issues or PRs:
```
- [ ] [#4 Optimize regular expressions](https://github.com/varunbiswasgit/aiacode/issues/4)
```

---

## Running Tests
See [tests/README.md](tests/README.md) for instructions on running the test suite.

---

## Contributing
Contributions are welcome! Please read our [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

---

## License
This project is licensed under the [GNU General Public License v3.0](LICENSE).

---

## Notes
- If you encounter warnings or errors during script execution, check input file formatting or verify tool installations.
