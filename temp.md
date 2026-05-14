# Repo Reorganization — Task List

## Plan
Flatten all files (source + README + tests) into one functional subfolder per tool. No nested `tests/` subdirectories. Remove the `README/` and `tests/` folders at the root after moving all files.

---

## Target Structure
```
aiacode/
├── outlook-bold-list/
│   ├── BoldListPrefixesOutlook.bas
│   ├── README.md
│   ├── TESTING.md
│   └── Test_BoldListPrefixes.bas
├── outlook-keyword-search/
│   ├── OutlookKeywordSearch_PS.bas
│   ├── OutlookKeywordSearch_PS.ps1
│   ├── OutlookKeywordSearch_Standalone.bas
│   ├── README-PS.md
│   ├── README-Standalone.md
│   ├── TESTING-PS.md
│   └── TESTING-Standalone.md
├── whatsapp-cleaner/
│   ├── Clean_data_whatsapp.ps1
│   ├── README.md
│   ├── TESTING.md
│   └── Clean_data_whatsapp.Tests.ps1
├── excel-tools/
│   ├── ExcelFormatting.bas
│   ├── SplitExcelByManager.bas
│   ├── split_excel_by_manager.py
│   ├── README-ExcelFormatting.md
│   ├── README-SplitExcelByManager-bas.md
│   ├── README-SplitExcelByManager-py.md
│   ├── TESTING-ExcelFormatting.md
│   ├── TESTING-SplitExcelByManager-bas.md
│   ├── TESTING-SplitExcelByManager-py.md
│   └── test_split_excel_by_manager.py
├── word-tools/
│   ├── WordNormalizeTable.bas
│   ├── WordResizeBorderImagesCleanlines.bas
│   ├── README-WordNormalizeTable.md
│   ├── README-WordResizeBorderImages.md
│   ├── TESTING-WordNormalizeTable.md
│   └── TESTING-WordResizeBorderImages.md
├── win11-startup/
│   ├── Win11startup.ps1
│   ├── README.md
│   └── TESTING.md
├── docs/
│   └── tests-overview.md
├── README.md  (updated)
├── CONTRIBUTING.md
└── LICENSE
```

---

## Tasks

### outlook-bold-list/
- [ ] Create `outlook-bold-list/BoldListPrefixesOutlook.bas`
- [ ] Create `outlook-bold-list/README.md` (from `README/BoldListPrefixesOutlook.bas README.md`)
- [ ] Create `outlook-bold-list/TESTING.md` (from `tests/BoldListPrefixes.bas Testing Readme.md`)
- [ ] Create `outlook-bold-list/Test_BoldListPrefixes.bas` (from `tests/Test_BoldListPrefixes.bas`)

### outlook-keyword-search/
- [ ] Create `outlook-keyword-search/OutlookKeywordSearch_PS.bas`
- [ ] Create `outlook-keyword-search/OutlookKeywordSearch_PS.ps1`
- [ ] Create `outlook-keyword-search/OutlookKeywordSearch_Standalone.bas`
- [ ] Create `outlook-keyword-search/README-PS.md`
- [ ] Create `outlook-keyword-search/README-Standalone.md`
- [ ] Create `outlook-keyword-search/TESTING-PS.md`
- [ ] Create `outlook-keyword-search/TESTING-Standalone.md`

### whatsapp-cleaner/
- [ ] Create `whatsapp-cleaner/Clean_data_whatsapp.ps1`
- [ ] Create `whatsapp-cleaner/README.md`
- [ ] Create `whatsapp-cleaner/TESTING.md`
- [ ] Create `whatsapp-cleaner/Clean_data_whatsapp.Tests.ps1`

### excel-tools/
- [ ] Create `excel-tools/ExcelFormatting.bas`
- [ ] Create `excel-tools/SplitExcelByManager.bas`
- [ ] Create `excel-tools/split_excel_by_manager.py`
- [ ] Create `excel-tools/README-ExcelFormatting.md`
- [ ] Create `excel-tools/README-SplitExcelByManager-bas.md`
- [ ] Create `excel-tools/README-SplitExcelByManager-py.md`
- [ ] Create `excel-tools/TESTING-ExcelFormatting.md`
- [ ] Create `excel-tools/TESTING-SplitExcelByManager-bas.md`
- [ ] Create `excel-tools/TESTING-SplitExcelByManager-py.md`
- [ ] Create `excel-tools/test_split_excel_by_manager.py`

### word-tools/
- [ ] Create `word-tools/WordNormalizeTable.bas`
- [ ] Create `word-tools/WordResizeBorderImagesCleanlines.bas`
- [ ] Create `word-tools/README-WordNormalizeTable.md`
- [ ] Create `word-tools/README-WordResizeBorderImages.md`
- [ ] Create `word-tools/TESTING-WordNormalizeTable.md`
- [ ] Create `word-tools/TESTING-WordResizeBorderImages.md`

### win11-startup/
- [ ] Create `win11-startup/Win11startup.ps1`
- [ ] Create `win11-startup/README.md`
- [ ] Create `win11-startup/TESTING.md`

### docs/
- [ ] Create `docs/tests-overview.md` (from `tests/README.md`)

### Root cleanup
- [ ] Update root `README.md`
- [ ] Delete all files from `scripts/`
- [ ] Delete all files from `README/`
- [ ] Delete all files from `tests/`
