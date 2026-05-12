# aiacode — AI-Assisted Automation Scripts

A curated collection of VBA macros, PowerShell, and Python scripts developed with AI assistance. All scripts are tested and documented; individual README files and test coverage are provided for each.

---

## ⚠️ AI Contributor Policy — README Updates

**This applies to any AI agent, Copilot, or automated tool pushing changes to this repository.**

Whenever you modify a script or any file in `scripts/`, you **must** also update all four of the following README files in the same commit or PR. Updating only some of them is not acceptable.

| README file | Location | What it covers |
|-------------|----------|----------------|
| Global README | [`README.md`](README.md) | Script table, descriptions, repository overview |
| Per-script README | [`README/<script filename> README.md`](README/) | Purpose, fields, logic, configuration, version history |
| Testing README | [`tests/<script filename> Testing Readme.md`](tests/) | Manual test cases, pass criteria, test environment |
| Tests index | [`tests/README.md`](tests/README.md) | Index of all test files and their coverage |

### What triggers a README update

Any of the following changes require all four READMEs to be reviewed and updated:

- Adding, removing, or renaming a field in the `$apps` array or equivalent config structure
- Adding, changing, or removing a launch strategy or function
- Changing default configuration values (`$InitialDelaySeconds`, `$MaxRepairDepth`, etc.)
- Adding or removing a fallback or repair step in any function
- Changing what is logged to the console
- Changing how errors or failures are handled
- Any change that affects the test cases documented in the Testing README
- Adding a new script to the repository

### Minimum update requirement per README

- **Global README** — update the script description in the Scripts table if behaviour changed.
- **Per-script README** — update the relevant section (fields, logic flow, configuration table, version history). Always add a version history entry.
- **Testing README** — add, update, or remove test cases to match the new behaviour. Update the pass criteria table.
- **Tests index** — update the entry for this script if the test count or coverage changed.

### Version history

Every code change must add a new row to the version history table in the per-script README. Do not edit existing rows. Format:

```
| vN | One-line summary of what changed |
```

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
| [`OutlookKeywordSearch.bas`](scripts/OutlookKeywordSearch.bas) | VBA (Outlook) | Searches all Outlook folders and subfolders for a keyword in the email body. Single mode returns the oldest match via MsgBox and Immediate Window. Batch mode reads keywords from a user-specified Excel column and appends Match Email, Sender, and Status columns automatically. |
| [`SplitExcelByManager.bas`](scripts/SplitExcelByManager.bas) | VBA (Excel) | Splits an Excel workbook into separate files, one per unique manager name, with sanitized filenames and configurable column targeting. |
| [`split_excel_by_manager.py`](scripts/split_excel_by_manager.py) | Python | Cross-platform equivalent of `SplitExcelByManager.bas` using pandas and openpyxl. Produces identical output filenames. |
| [`WordNormalizeTable.bas`](scripts/WordNormalizeTable.bas) | VBA (Word) | Normalizes all tables in the active Word document — sets width to 100%, clears row/cell constraints, and applies Arial 10pt. Two subroutines: `NormalizeTables_Light` (body only) and `StandardizeTables_TwoPass_AllStories` (all stories including headers, footers, and text boxes). |
| [`WordResizeBorderImagesCleanlines.bas`](scripts/WordResizeBorderImagesCleanlines.bas) | VBA (Word) | Resizes inline images to a user-specified width range, applies a configurable border (RGB or hex color), and cleans blank paragraphs including ghost bullet lines and consecutive blank lines. |
| [`Win11startup.ps1`](scripts/Win11startup.ps1) | PowerShell | Self-healing Windows 11 startup launcher. Sequentially starts a curated list of applications using Win32 shortcut-based repair (depth 3, user-prompt fallback) or dynamic Appx AUMID resolution (Get-StartApps → KnownAumid verification → AppxPackage manifest). |

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
