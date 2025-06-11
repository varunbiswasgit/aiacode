# split_excel_by_manager.py

This Python script reads an Excel file and saves a separate workbook for each unique value in the **Manager** column. It uses `pandas` for splitting and `openpyxl` to adjust the column widths of each output file automatically.

## Usage
```bash
python split_excel_by_manager.py data.xlsx [ManagerColumn]
```
If `ManagerColumn` is omitted, it defaults to `Manager`.
