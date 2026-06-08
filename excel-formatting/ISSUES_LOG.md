# Issues Log — ExcelFormatting.bas

---

## [BUG-001] Method 'Range' of object '_Worksheet' Failed

| Field | Detail |
|---|---|
| **Date Reported** | 2026-06-08 |
| **Date Fixed** | 2026-06-08 |
| **Severity** | High — macro halted on Options 2 and 3 |
| **Status** | ✅ Fixed |
| **File** | `excel-formatting/ExcelFormatting.bas` |
| **Subroutine** | `ProcessSheetCore` |
| **Commit** | `123a31cd82cda5ec49b58e47c1f070ca8ce10f94` |

### Description

When running Options 2 or 3 of the Unified Data Formatter macro, a VBA runtime error popup appeared:

> **Error: Method 'Range' of object '_Worksheet' failed**

The macro stopped immediately after user selection.

### Root Cause

Invalid usage of `ws.Range()` with a `Rows()` argument inside `ProcessSheetCore`:

```vb
' ❌ INVALID — ws.Range() does not accept a Rows() object as argument
CleanTextInRange ws.Range(ws.Rows(1))
```

`ws.Range()` expects either a cell address string (e.g. `"A1:Z1"`) or two `Cells()` corner arguments. Passing `ws.Rows(1)` into it triggers the runtime error.

### Fix Applied

```vb
' ✅ CORRECT — ws.Rows(1) is already a valid Range object
CleanTextInRange ws.Rows(1)
```

`ws.Rows(1)` returns a valid `Range` object representing the entire first row and can be passed directly to `CleanTextInRange(ByVal rng As Range)`.

---
