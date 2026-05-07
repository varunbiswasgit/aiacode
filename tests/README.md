# Tests

This folder contains all test assets for the `aiacode` repository — automated test scripts and manual test case documents.

---

## Testing Methodology

### Automated Tests
Scripts that are non-interactive (no UI prompts) have automated test files:

| Tool | Command |
|------|---------|
| **Pester (PowerShell)** | `Invoke-Pester .\Clean_data_whatsapp.Tests.ps1 -Output Detailed` |
| **pytest (Python)** | `pytest tests/test_split_excel_by_manager.py -v` |

Run Pester from the `tests/` folder. Run pytest from the repo root.

### Manual Test Cases
Scripts driven by interactive UI prompts (VBA MsgBox / InputBox) cannot be automated. Each has a companion `.md` file listing step-by-step test cases with setup, input, expected outcome, and pass criteria — consistent format across all scripts.

---

## Test Assets per Script

| Script | Automated Test File | Manual Test Cases |
|--------|--------------------|-----------------|
| `Clean_data_whatsapp.ps1` | `Clean_data_whatsapp.Tests.ps1` | `Clean_data_whatsapp_test_cases.md` |
| `ExcelFormatting.bas` | *(UI-driven — not automated)* | `ExcelFormatting test cases.md` |
| `SplitExcelByManager.bas` | *(UI-driven — not automated)* | `SplitExcelByManager_test_cases.md` |
| `split_excel_by_manager.py` | `test_split_excel_by_manager.py` | *(pytest covers automation)* |
| `NormalizeTable.bas` | *(UI-driven — not automated)* | *(planned)* |
| `WordResizeBorderImagesCleanlines.bas` | *(UI-driven — not automated)* | *(planned)* |

---

## Running Automated Tests

### Pester — PowerShell (Clean_data_whatsapp.ps1)
```powershell
# From repo root
Invoke-Pester .\tests\Clean_data_whatsapp.Tests.ps1 -Output Detailed
```
Requires Pester v5+: `Install-Module Pester -Force`

### pytest — Python (split_excel_by_manager.py)
```bash
# From repo root
pip install pytest pandas openpyxl
pytest tests/test_split_excel_by_manager.py -v
```

---

## Manual Test Execution

Open the relevant `.md` file for the script under test. Each test case specifies:
- **Setup** — what data or state to prepare
- **Input** — what to enter at each prompt
- **Expected** — what the script should produce
- **Pass criteria** — how to verify the result
