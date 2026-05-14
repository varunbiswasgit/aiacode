# Repo Reorganization Task List

## Target Folder Structure

```
aiacode/
├── win11-startup/          ✅ already created
├── outlook-bold-list/      (BoldListPrefixesOutlook)
├── outlook-keyword-search/ (OutlookKeywordSearch PS + Standalone)
├── whatsapp-cleaner/       (Clean_data_whatsapp)
├── excel-tools/            (ExcelFormatting + SplitExcelByManager)
├── word-tools/             (WordNormalizeTable + WordResizeBorderImagesCleanlines)
```

---

## Tasks

### win11-startup/ (already created — cleanup only)
- [ ] Delete `scripts/Win11startup.ps1` (duplicate)
- [ ] Delete `README/Win11startup.ps1 README.md` (merged into win11-startup/README.md)
- [ ] Delete `tests/Win11startup.ps1 Testing Readme.md` (merged into win11-startup/TESTING.md)

### outlook-bold-list/
- [ ] Move `scripts/BoldListPrefixesOutlook.bas` → `outlook-bold-list/BoldListPrefixesOutlook.bas`
- [ ] Move `tests/Test_BoldListPrefixes.bas` → `outlook-bold-list/Test_BoldListPrefixes.bas`
- [ ] Move `README/BoldListPrefixesOutlook.bas README.md` → `outlook-bold-list/README.md`
- [ ] Move `tests/BoldListPrefixes.bas Testing Readme.md` → `outlook-bold-list/TESTING.md`
- [ ] Delete originals from scripts/, README/, tests/

### outlook-keyword-search/
- [ ] Move `scripts/OutlookKeywordSearch_PS.bas` → `outlook-keyword-search/OutlookKeywordSearch_PS.bas`
- [ ] Move `scripts/OutlookKeywordSearch_PS.ps1` → `outlook-keyword-search/OutlookKeywordSearch_PS.ps1`
- [ ] Move `scripts/OutlookKeywordSearch_Standalone.bas` → `outlook-keyword-search/OutlookKeywordSearch_Standalone.bas`
- [ ] Move `README/OutlookKeywordSearch_PS.bas README.md` → `outlook-keyword-search/README_PS.md`
- [ ] Move `README/OutlookKeywordSearch_Standalone.bas README.md` → `outlook-keyword-search/README_Standalone.md`
- [ ] Move `tests/OutlookKeywordSearch_PS.bas Testing Readme.md` → `outlook-keyword-search/TESTING_PS.md`
- [ ] Move `tests/OutlookKeywordSearch_Standalone.bas Testing Readme.md` → `outlook-keyword-search/TESTING_Standalone.md`
- [ ] Delete originals from scripts/, README/, tests/

### whatsapp-cleaner/
- [ ] Move `scripts/Clean_data_whatsapp.ps1` → `whatsapp-cleaner/Clean_data_whatsapp.ps1`
- [ ] Move `tests/Clean_data_whatsapp.Tests.ps1` → `whatsapp-cleaner/Clean_data_whatsapp.Tests.ps1`
- [ ] Move `README/Clean_data_whatsapp.ps1 README.md` → `whatsapp-cleaner/README.md`
- [ ] Move `tests/Clean_data_whatsapp.ps1 Testing Readme.md` → `whatsapp-cleaner/TESTING.md`
- [ ] Delete originals from scripts/, README/, tests/

### excel-tools/
- [ ] Move `scripts/ExcelFormatting.bas` → `excel-tools/ExcelFormatting.bas`
- [ ] Move `scripts/SplitExcelByManager.bas` → `excel-tools/SplitExcelByManager.bas`
- [ ] Move `scripts/split_excel_by_manager.py` → `excel-tools/split_excel_by_manager.py`
- [ ] Move `tests/test_split_excel_by_manager.py` → `excel-tools/test_split_excel_by_manager.py`
- [ ] Move `README/ExcelFormatting.bas README.md` → `excel-tools/README_ExcelFormatting.md`
- [ ] Move `README/SplitExcelByManager.bas README.md` → `excel-tools/README_SplitExcelByManager.md`
- [ ] Move `README/split_excel_by_manager.py.README.md` → `excel-tools/README_split_excel_by_manager_py.md`
- [ ] Move `tests/ExcelFormatting.bas Testing Readme.md` → `excel-tools/TESTING_ExcelFormatting.md`
- [ ] Move `tests/SplitExcelByManager.bas Testing Readme.md` → `excel-tools/TESTING_SplitExcelByManager.md`
- [ ] Move `tests/split_excel_by_manager.py Testing Readme.md` → `excel-tools/TESTING_split_excel_by_manager_py.md`
- [ ] Delete originals from scripts/, README/, tests/

### word-tools/
- [ ] Move `scripts/WordNormalizeTable.bas` → `word-tools/WordNormalizeTable.bas`
- [ ] Move `scripts/WordResizeBorderImagesCleanlines.bas` → `word-tools/WordResizeBorderImagesCleanlines.bas`
- [ ] Move `README/WordNormalizeTable.bas README.md` → `word-tools/README_WordNormalizeTable.md`
- [ ] Move `README/WordResizeBorderImagesCleanlines.bas README.md` → `word-tools/README_WordResizeBorderImagesCleanlines.md`
- [ ] Move `tests/WordNormalizeTable.bas Testing Readme.md` → `word-tools/TESTING_WordNormalizeTable.md`
- [ ] Move `tests/WordResizeBorderImagesCleanlines.bas Testing Readme.md` → `word-tools/TESTING_WordResizeBorderImagesCleanlines.md`
- [ ] Delete originals from scripts/, README/, tests/

### Cleanup (after all moves)
- [ ] Delete empty `scripts/` folder (all files moved)
- [ ] Delete empty `README/` folder (all files moved)
- [ ] Delete empty `tests/` folder (all files moved)
- [ ] Delete `tests/README.md` (general env setup — content absorbed into individual TESTING.md files)
- [ ] Delete `temp.md` (this file)
