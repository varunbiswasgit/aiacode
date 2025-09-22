# Split Excel by Manager - Python Script

## Overview
This Python script reads an Excel file and creates separate workbooks for each unique manager in a specified column. The enhanced version includes comprehensive error handling, data validation, and improved user experience.

## Features

### ✨ **Core Functionality**
- **Multi-format Support**: Works with `.xlsx`, `.xls`, and `.xlsm` files
- **Automatic Column Width Adjustment**: Optimizes column widths for readability
- **Organized Output**: Creates reports in a dedicated `manager_reports/` directory
- **Cross-platform Compatibility**: Works on Windows, macOS, and Linux

### 🛡️ **Robust Error Handling**
- **File Validation**: Checks file existence, format, and accessibility
- **Data Validation**: Validates column existence and data integrity
- **Permission Handling**: Graceful handling of file access restrictions
- **Individual Processing**: One manager's failure doesn't stop others

### 📝 **Enhanced User Experience**
- **Progress Reporting**: Real-time feedback during processing
- **Detailed Error Messages**: Clear explanations when issues occur
- **Success Statistics**: Summary of completed vs. failed operations
- **Help Documentation**: Built-in usage instructions

## Installation

### Prerequisites
Install required Python packages:
```bash
pip install pandas openpyxl pathlib
```

### Dependencies
- **pandas**: Excel file reading and data manipulation
- **openpyxl**: Excel file writing and formatting
- **pathlib**: Cross-platform file path handling
- **re**: Regular expressions for filename sanitization

## Usage

### Basic Usage
```bash
python split_excel_by_manager.py <excel_file> [manager_column]
```

### Parameters
- **`excel_file`** (required): Path to the input Excel file
- **`manager_column`** (optional): Column name containing manager data (default: 'Manager')

### Examples

#### Example 1: Default Manager Column
```bash
python split_excel_by_manager.py employees.xlsx
```
Uses the 'Manager' column to split the data.

#### Example 2: Custom Column Name
```bash
python split_excel_by_manager.py staff_data.xlsx "Supervisor"
```
Uses the 'Supervisor' column to split the data.

#### Example 3: File with Spaces in Path
```bash
python split_excel_by_manager.py "C:\\Data Files\\employee list.xlsx" "Team Lead"
```
Handles file paths with spaces and custom column names.

## Output Structure

### Directory Organization
```
project_directory/
├── split_excel_by_manager.py
├── your_input_file.xlsx
└── manager_reports/
    ├── John_Smith_report.xlsx
    ├── Sarah_Johnson_report.xlsx
    ├── Michael_Brown_report.xlsx
    └── ...
```

### File Naming Convention
- **Pattern**: `{Manager_Name}_report.xlsx`
- **Sanitization**: Special characters replaced with underscores
- **Length Limits**: Filenames truncated if too long
- **Reserved Names**: Handles Windows reserved names (CON, PRN, etc.)

## Error Handling Features

### 📝 **File System Errors**
- File not found
- Permission denied (file open in another app)
- Invalid file extensions
- Directory creation failures

### 📋 **Data Validation**
- Empty Excel files
- Missing or invalid column names
- Null/empty manager data
- Corrupted Excel files

### 💾 **Processing Errors**
- Individual manager processing failures
- Memory limitations
- Disk space issues
- Keyboard interruption (Ctrl+C)

## Advanced Features

### Filename Sanitization
```python
# Special characters are automatically handled:
"John & Mary's Team" → "John___Mary_s_Team_report.xlsx"
"Sales/Marketing Dept" → "Sales_Marketing_Dept_report.xlsx"
```

### Column Width Optimization
- **Minimum width**: 8 characters
- **Maximum width**: 50 characters
- **Dynamic sizing**: Based on content length
- **Readable formatting**: Adds padding for clarity

### Progress Tracking
The script provides real-time feedback:
```
Reading Excel file: employees.xlsx
Found 5 unique managers.
Creating report for: John Smith -> manager_reports/John_Smith_report.xlsx
Creating report for: Sarah Johnson -> manager_reports/Sarah_Johnson_report.xlsx
...
Processing complete!
Successfully created 5 out of 5 manager reports.
```

## Troubleshooting

### Common Issues

#### “File not found”
- Verify the file path is correct
- Use quotes around paths with spaces
- Check file permissions

#### “Column not found”
- Check the exact column name (case-sensitive)
- View available columns in the error message
- Ensure the header row exists

#### “Permission denied”
- Close the Excel file if it's open
- Check folder write permissions
- Run as administrator if necessary

### Getting Help
Run the script without arguments to see usage instructions:
```bash
python split_excel_by_manager.py
```

## Performance Considerations

### Large Files
- **Memory Usage**: Pandas loads entire file into memory
- **Processing Time**: Linear with number of rows and managers
- **Disk Space**: Requires space for all output files

### Optimization Tips
- Close other applications when processing large files
- Ensure sufficient disk space (2-3x input file size)
- Use SSD storage for faster I/O operations

## Version History

### v2.0 (Latest)
- ✅ Comprehensive error handling
- ✅ Filename sanitization
- ✅ Organized output directory
- ✅ Progress reporting
- ✅ Enhanced user interface
- ✅ Cross-platform compatibility

### v1.0 (Original)
- ✅ Basic Excel splitting functionality
- ✅ Column width adjustment
- ❌ Limited error handling
- ❌ No input validation

## Contributing

Contributions are welcome! Please see the [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to contribute.

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).