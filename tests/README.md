# tests/ — Test Index

This folder contains automated test scripts and manual Testing README files for every script in `scripts/`.

---

## Test Files

| Script under test | Testing README | Automated test file | Test count |
|---|---|---|---|
| `BoldListPrefixes.bas` | [BoldListPrefixes.bas Testing Readme.md](BoldListPrefixes.bas%20Testing%20Readme.md) | [Test_BoldListPrefixes.bas](Test_BoldListPrefixes.bas) | 7 automated / 10 manual |
| `Clean_data_whatsapp.ps1` | [Clean_data_whatsapp.ps1 Testing Readme.md](Clean_data_whatsapp.ps1%20Testing%20Readme.md) | [Clean_data_whatsapp.Tests.ps1](Clean_data_whatsapp.Tests.ps1) | See Testing README |
| `ExcelFormatting.bas` | [ExcelFormatting.bas Testing Readme.md](ExcelFormatting.bas%20Testing%20Readme.md) | Manual only | See Testing README |
| `OutlookKeywordSearch_PS.bas` | [OutlookKeywordSearch_PS.bas Testing Readme.md](OutlookKeywordSearch_PS.bas%20Testing%20Readme.md) | Manual only | See Testing README |
| `OutlookKeywordSearch_Standalone.bas` | [OutlookKeywordSearch_Standalone.bas Testing Readme.md](OutlookKeywordSearch_Standalone.bas%20Testing%20Readme.md) | Manual only | See Testing README |
| `SplitExcelByManager.bas` | [SplitExcelByManager.bas Testing Readme.md](SplitExcelByManager.bas%20Testing%20Readme.md) | Manual only | See Testing README |
| `split_excel_by_manager.py` | [split_excel_by_manager.py Testing Readme.md](split_excel_by_manager.py%20Testing%20Readme.md) | [test_split_excel_by_manager.py](test_split_excel_by_manager.py) | See Testing README |
| `Win11startup.ps1` | [Win11startup.ps1 Testing Readme.md](Win11startup.ps1%20Testing%20Readme.md) | Manual only | See Testing README |
| `WordNormalizeTable.bas` | [WordNormalizeTable.bas Testing Readme.md](WordNormalizeTable.bas%20Testing%20Readme.md) | Manual only | See Testing README |
| `WordResizeBorderImagesCleanlines.bas` | [WordResizeBorderImagesCleanlines.bas Testing Readme.md](WordResizeBorderImagesCleanlines.bas%20Testing%20Readme.md) | Manual only | See Testing README |

---

## Running Automated Tests

### VBA tests (BoldListPrefixes)

1. Open Microsoft Word.
2. Open the VBA editor (`Alt+F11`).
3. Import both `scripts/BoldListPrefixes.bas` and `tests/Test_BoldListPrefixes.bas` into any standard module.
4. Run `RunAllTests` (`Alt+F8 → RunAllTests → Run`).
5. Results appear in the **Immediate window** (`Ctrl+G`) and a summary `MsgBox`.

### Python tests (split_excel_by_manager)

```bash
pip install pandas openpyxl pytest
pytest tests/test_split_excel_by_manager.py -v
```

### PowerShell tests (Clean_data_whatsapp)

```powershell
cd tests
Invoke-Pester .\Clean_data_whatsapp.Tests.ps1 -Output Detailed
```

---

## Manual Test Execution

For scripts without automated tests, follow the step-by-step test cases in the corresponding Testing README. Each test case specifies:
- **Setup** — document or data state required before running
- **Action** — what to run
- **Expected** — what the output should be
- **Pass criteria** — unambiguous pass/fail condition
