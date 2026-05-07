# tests/

This folder contains all test scripts and their corresponding Testing Readme files.

## Structure

| Script under test | Test file | Testing Readme |
|---|---|---|
| `scripts/ExcelFormatting.bas` | *(manual only — see readme)* | `ExcelFormatting.bas Testing Readme.md` |
| `scripts/SplitExcelByManager.bas` | *(manual only — see readme)* | `SplitExcelByManager.bas Testing Readme.md` |
| `scripts/split_excel_by_manager.py` | `test_split_excel_by_manager.py` | `test_split_excel_by_manager.py Testing Readme.md` |
| `scripts/Clean_data_whatsapp.ps1` | `Clean_data_whatsapp.Tests.ps1` | `Clean_data_whatsapp.ps1 Testing Readme.md` |

## Naming Convention

Every Testing Readme follows the pattern:
```
<script filename> Testing Readme.md
```

## Running Automated Tests

### Python (pytest)
```bash
pytest tests/test_split_excel_by_manager.py -v
```

### PowerShell (Pester)
```powershell
Invoke-Pester .\tests\Clean_data_whatsapp.Tests.ps1 -Output Detailed
```

### VBA / Word macros
Manual only. Follow the environment setup steps in the respective Testing Readme.
