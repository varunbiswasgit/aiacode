# Split Excel by Manager

## Overview
`split_excel_by_manager.py` is a Python utility that divides a single Excel workbook into multiple files, one for each manager listed in a column (default `Manager`). Each resulting file contains only the rows associated with that manager.

## Usage
1. Ensure you have Python and `pandas` installed:
   ```bash
   pip install pandas
   ```
2. Run the script specifying the source Excel file:
   ```bash
   python split_excel_by_manager.py path/to/input.xlsx -o output_directory
   ```
   - `-o` or `--output-dir` sets the directory where individual files will be saved. The directory is created if it does not exist.
   - `-c` or `--column` lets you specify a different column name if your manager field is not called `Manager`.

## Example
```bash
python split_excel_by_manager.py staff_list.xlsx -o splits
```
This command produces one Excel file per unique manager name, stored in the `splits` directory.

## Task List for split_excel_by_manager.py
- [ ] Add option for CSV output
- [ ] Add unit tests
- [ ] Support Google Sheets input

## License
This script is released under the [GNU General Public License v3.0](LICENSE).
