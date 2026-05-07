# split_excel_by_manager.py

A Python CLI script that splits an Excel file into one `.xlsx` workbook per unique manager value. Maintains feature parity with `SplitExcelByManager.bas` v3.0.

## Features

- Configurable manager column name (default: `Manager`)
- Configurable output directory (default: `manager_reports`)
- Sanitizes filenames: invalid characters replaced with `_`, Windows reserved names mapped to `Unknown_Manager`, length capped at 200 characters
- Auto-fits column widths, clamped 8–50 characters
- Skips blank/null manager entries
- Per-manager error handling — one failure does not abort the run
- Prints a structured summary on completion

## Usage

```bash
python split_excel_by_manager.py <file> [column] [output_dir]
```

### Arguments

| Argument | Default | Description |
|----------|---------|-------------|
| `file` | *(required)* | Path to the input Excel file (`.xlsx`, `.xls`, `.xlsm`) |
| `column` | `Manager` | Header of the manager column |
| `output_dir` | `manager_reports` | Output folder |

### Examples

```bash
python split_excel_by_manager.py data.xlsx
python split_excel_by_manager.py staff.xlsx Supervisor
python split_excel_by_manager.py staff.xlsx Supervisor ./output
```

## Requirements

```
pandas
openpyxl
```

Install with:

```bash
pip install pandas openpyxl
```

## Running Tests

```bash
pytest tests/test_split_excel_by_manager.py -v
```

The test suite covers 15 cases across sanitization, edge inputs, split correctness, and column width clamping.

## License

See [LICENSE](../LICENSE) in the repository root.
