# Split Excel by Manager

Two tools — a VBA macro and a Python CLI script — that split an Excel sheet into one `.xlsx` workbook per unique manager value. Both tools maintain full feature parity at v3.0.

## Tools

| File | Type | When to use |
|------|------|-------------|
| `SplitExcelByManager.bas` | Excel VBA macro | Run directly inside Excel; no Python required |
| `split_excel_by_manager.py` | Python CLI script | Automate from the command line or CI pipelines |

---

## SplitExcelByManager.bas

An Excel VBA macro that splits the active sheet into one `.xlsx` workbook per unique manager value, saving each file to a `manager_reports` subfolder beside the open workbook.

### Features

- Locates the manager column by header name — no hardcoded column index
- Sanitizes filenames: replaces invalid characters, guards against Windows reserved device names (`CON`, `PRN`, `NUL`, `COM1`–`COM9`, `LPT1`–`LPT9`), caps length at 200 characters
- Auto-fits column widths, clamped between 8 and 50 characters
- Skips blank manager entries
- Per-manager error handling — one failure does not abort the entire run
- Displays a completion summary: reports saved vs. total unique managers

### Configuration

Edit the two constants at the top of the macro before running:

```vb
Const MANAGER_COL_NAME  As String = "Manager"          ' Column header
Const OUTPUT_SUBFOLDER  As String = "manager_reports"  ' Output subfolder
```

### Requirements

- Microsoft Excel (any version supporting VBA)
- Macro execution must be enabled
- The workbook must be saved before running (output path is derived from `ThisWorkbook.Path`)

### Installation

1. Press **Alt + F11** in Excel.
2. Right-click your workbook in Project Explorer → **Import File**.
3. Select `SplitExcelByManager.bas`.
4. Run `SplitExcelByManager` from **Developer → Macros**.

---

## split_excel_by_manager.py

A Python CLI script that splits an Excel file into one `.xlsx` workbook per unique manager value.

### Features

- Configurable manager column name (default: `Manager`)
- Configurable output directory (default: `manager_reports`)
- Sanitizes filenames: invalid characters replaced with `_`, Windows reserved names mapped to `Unknown_Manager`, length capped at 200 characters
- Auto-fits column widths, clamped 8–50 characters
- Skips blank/null manager entries
- Per-manager error handling — one failure does not abort the run
- Prints a structured summary on completion

### Usage

```bash
python split_excel_by_manager.py <file> [column] [output_dir]
```

#### Arguments

| Argument | Default | Description |
|----------|---------|-------------|
| `file` | *(required)* | Path to the input Excel file (`.xlsx`, `.xls`, `.xlsm`) |
| `column` | `Manager` | Header of the manager column |
| `output_dir` | `manager_reports` | Output folder |

#### Examples

```bash
python split_excel_by_manager.py data.xlsx
python split_excel_by_manager.py staff.xlsx Supervisor
python split_excel_by_manager.py staff.xlsx Supervisor ./output
```

### Requirements

```
pandas
openpyxl
```

Install with:

```bash
pip install pandas openpyxl
```

### Running Tests

```bash
pytest tests/test_split_excel_by_manager.py -v
```

The test suite covers 15 cases across sanitization, edge inputs, split correctness, and column width clamping.

---

## License

See [LICENSE](../LICENSE) in the repository root.
