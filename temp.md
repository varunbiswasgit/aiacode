# Repo Reorganization — Task List (Refreshed May 14, 2026)

## Current State
You reorganized the repo yourself. Below is what is DONE and what still remains.

---

## Folders: Status

### ✅ bold-list-prefixes-outlook/ — COMPLETE
All 4 files present:
- `BoldListPrefixesOutlook.bas`
- `README.md`
- `TESTING.md`
- `Test_BoldListPrefixes.bas`

### ✅ clean-data-whatsapp/ — COMPLETE
All 4 files present:
- `Clean_data_whatsapp.ps1`
- `Clean_data_whatsapp.Tests.ps1`
- `README.md`
- `TESTING.md`

### ✅ excel-formatting/ — COMPLETE (1 stray file to delete)
Files present:
- `ExcelFormatting.bas`
- `README.md`
- `TESTING.md`
- ⚠️ `ExcelFormatting Testing Readme.md` — stray file with spaces in name; DELETE

### ✅ outlook-keyword-search/ — COMPLETE
All 5 files present:
- `OutlookKeywordSearch_PS.bas`
- `OutlookKeywordSearch_PS.ps1`
- `OutlookKeywordSearch_Standalone.bas`
- `README.md`
- `TESTING.md`

### ✅ split-excel-by-manager/ — COMPLETE
All 4 files present:
- `SplitExcelByManager.bas`
- `split_excel_by_manager.py`
- `README.md`
- `TESTING.md`

### ✅ win11-startup/ — COMPLETE
All 3 files present:
- `Win11startup.ps1`
- `README.md`
- `TESTING.md`

---

## Remaining Tasks

### word-tools/ — NOT CREATED YET
Scripts exist in `scripts/` but no dedicated folder created yet.
- [ ] Create `word-tools/WordNormalizeTable.bas` (from `scripts/WordNormalizeTable.bas`)
- [ ] Create `word-tools/WordResizeBorderImagesCleanlines.bas` (from `scripts/WordResizeBorderImagesCleanlines.bas`)
- [ ] Create `word-tools/README-WordNormalizeTable.md` (from `README/WordNormalizeTable.bas README.md`)
- [ ] Create `word-tools/README-WordResizeBorderImages.md` (from `README/WordResizeBorderImagesCleanlines.bas README.md`)
- [ ] Create `word-tools/TESTING-WordNormalizeTable.md` (from `tests/WordNormalizeTable.bas Testing Readme.md`)
- [ ] Create `word-tools/TESTING-WordResizeBorderImages.md` (from `tests/WordResizeBorderImagesCleanlines.bas Testing Readme.md`)

### excel-formatting/ — Stray file
- [ ] Delete `excel-formatting/ExcelFormatting Testing Readme.md`

### Root Cleanup — Delete legacy folders
Delete all files in these folders (GitHub requires deleting file-by-file):

**scripts/** (11 files to delete):
- [ ] `scripts/BoldListPrefixesOutlook.bas`
- [ ] `scripts/Clean_data_whatsapp.ps1`
- [ ] `scripts/ExcelFormatting.bas`
- [ ] `scripts/OutlookKeywordSearch_PS.bas`
- [ ] `scripts/OutlookKeywordSearch_PS.ps1`
- [ ] `scripts/OutlookKeywordSearch_Standalone.bas`
- [ ] `scripts/SplitExcelByManager.bas`
- [ ] `scripts/Win11startup.ps1`
- [ ] `scripts/WordNormalizeTable.bas`
- [ ] `scripts/WordResizeBorderImagesCleanlines.bas`
- [ ] `scripts/split_excel_by_manager.py`

**tests/** (13 files to delete):
- [ ] `tests/BoldListPrefixes.bas Testing Readme.md`
- [ ] `tests/Clean_data_whatsapp.Tests.ps1`
- [ ] `tests/Clean_data_whatsapp.ps1 Testing Readme.md`
- [ ] `tests/ExcelFormatting.bas Testing Readme.md`
- [ ] `tests/OutlookKeywordSearch_PS.bas Testing Readme.md`
- [ ] `tests/OutlookKeywordSearch_Standalone.bas Testing Readme.md`
- [ ] `tests/README.md`
- [ ] `tests/SplitExcelByManager.bas Testing Readme.md`
- [ ] `tests/Test_BoldListPrefixes.bas`
- [ ] `tests/Win11startup.ps1 Testing Readme.md`
- [ ] `tests/WordNormalizeTable.bas Testing Readme.md`
- [ ] `tests/WordResizeBorderImagesCleanlines.bas Testing Readme.md`
- [ ] `tests/split_excel_by_manager.py Testing Readme.md`
- [ ] `tests/test_split_excel_by_manager.py`

**README/** (10 files to delete):
- [ ] `README/BoldListPrefixesOutlook.bas README.md`
- [ ] `README/Clean_data_whatsapp.ps1 README.md`
- [ ] `README/ExcelFormatting.bas README.md`
- [ ] `README/OutlookKeywordSearch_PS.bas README.md`
- [ ] `README/OutlookKeywordSearch_Standalone.bas README.md`
- [ ] `README/SplitExcelByManager.bas README.md`
- [ ] `README/Win11startup.ps1 README.md`
- [ ] `README/WordNormalizeTable.bas README.md`
- [ ] `README/WordResizeBorderImagesCleanlines.bas README.md`
- [ ] `README/split_excel_by_manager.py.README.md`

### Root README.md — Update index table
- [ ] Update root `README.md` to reflect the 7 final folders

### temp.md — Delete when done
- [ ] Delete `temp.md` after all tasks are complete
