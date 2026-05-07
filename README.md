# aiacode — AI-Assisted Automation Scripts

A curated collection of VBA macros, PowerShell, and Python scripts developed with AI assistance. All scripts are tested and documented; individual README files and test coverage are provided for each.

---

## Repository Structure

```
aiacode/
├── scripts/          # All production scripts
├── README/           # Per-script README files
├── tests/            # Test scripts and Testing Readme files
├── CONTRIBUTING.md
├── LICENSE
└── README.md         # This file
```

---

## Scripts

| Script | Type | Description |
|--------|------|-------------|
| [`Clean_data_whatsapp.ps1`](scripts/Clean_data_whatsapp.ps1) | PowerShell | Cleans and processes exported WhatsApp chat text files — removes formatting artifacts, validates paths, and extracts structured data. |
| [`ExcelFormatting.bas`](scripts/ExcelFormatting.bas) | VBA (Excel) | Unified Excel formatting and cleanup macro with four interactive modes: simple formatting, advanced formatting, optional column splitting, and generic table extraction (formerly SAP-specific). |
| [`SplitExcelByManager.bas`](scripts/SplitExcelByManager.bas) | VBA (Excel) | Splits an Excel workbook into separate files, one per unique manager name, with sanitized filenames and configurable column targeting. |
| [`split_excel_by_manager.py`](scripts/split_excel_by_manager.py) | Python | Cross-platform equivalent of `SplitExcelByManager.bas` using pandas and openpyxl. Produces identical output filenames. |
| [`NormalizeTable.bas`](scripts/NormalizeTable.bas) | VBA (Word) | Normalizes all tables in the active Word document — sets width to 100%, clears row/cell constraints, and applies Arial 10pt. Two subroutines: `NormalizeTables_Light` (body only) and `StandardizeTables_TwoPass_AllStories` (all stories including headers, footers, and text boxes). |
| [`WordResizeBorderImagesCleanlines.bas`](scripts/WordResizeBorderImagesCleanlines.bas) | VBA (Word) | Resizes inline images to a user-specified width range, applies a configurable border (RGB or hex color), and cleans blank paragraphs including ghost bullet lines and consecutive blank lines. |

---

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/varunbiswasgit/aiacode.git
cd aiacode
```

### 2. Find the script you need

All scripts are in the [`scripts/`](scripts/) folder.

### 3. Read the script-specific README

All per-script READMEs are in the [`README/`](README/) folder, named `<script filename> README.md`.

---

## Running Tests

Test scripts and manual test case documentation are in the [`tests/`](tests/) folder. Each script has a corresponding Testing Readme named `<script filename> Testing Readme.md`.

For automated Python tests, run from the repository root:

```bash
pip install pandas openpyxl pytest
pytest tests/test_split_excel_by_manager.py -v
```

See [tests/README.md](tests/README.md) for the full test index.

---

## Contributing

Contributions are welcome. Please read [CONTRIBUTING.md](CONTRIBUTING.md) before submitting a pull request.

---

## License

This project is licensed under the [GNU General Public License v3.0](LICENSE).
