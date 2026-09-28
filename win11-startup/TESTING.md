# Testing Guide: Windows 11 Startup Manager

This document provides instructions for executing the automated unit test suite for `Win11startup.ps1`.

---

## 1. Prerequisites

The test suite is written using the **Pester** testing framework for PowerShell.

* Windows PowerShell 5.1 or PowerShell Core 7+
* Pester module (version 5.0 or later recommended)

To install or verify Pester, run in PowerShell:
```powershell
if (-not (Get-Module -ListAvailable -Name Pester)) {
    Install-Module -Name Pester -Scope CurrentUser -Force -SkipPublisherCheck
}
```

---

## 2. Test Structure

Tests are defined in `Win11startup.Tests.ps1`. The test suite targets core business logic functions in isolation without invoking the interactive menu loop:

* **Pattern Matching**: Validates the zero-padded two-digit numbering regex (`01-99`).
* **Process Name Derivation (`Get-ProcName`)**: Confirms executable name extraction and fallback behavior for `explorer.exe` targets.
* **Sequence Calculation (`Get-NextShortcutNumber`)**: Ensures monotonic ordering and gap-filling behavior in managed folders.
* **Configuration Safety (`Load-ConfigSafe`)**: Validates defensive loading and schema resilience against missing or malformed JSON files.

---

## 3. Running the Tests

Open PowerShell in the `win11-startup` directory and run:

```powershell
Invoke-Pester .\Win11startup.Tests.ps1 -Output Detailed
```

### Expected Output
```text
Running tests from 'Win11startup.Tests.ps1'
Describing Pattern Matching & Formatting Tests
  [+] Matches zero-padded numbers 01 to 99 followed by space
  [+] Rejects invalid or unnumbered shortcut formats
Describing Get-ProcName Function Tests
  [+] Extracts standard executable name without extension
  [+] Falls back to sanitized display name when target is explorer.exe
  [+] Falls back to sanitized display name when TargetPath is empty
Describing Sequence Number Calculation Tests
  Context With mock files in temporary folder
    [+] Returns 1 (formatted as 01) when no numbered files exist
    [+] Returns the next sequential integer when shortcuts exist
    [+] Ignores non-numbered lnk files when calculating the max number
Describing Configuration Loading Tests
  Context JSON schema resilience
    [+] Returns default configuration object if JSON file is missing
    [+] Correctly parses valid configuration JSON

Tests completed in ...ms
Tests Passed: 10, Failed: 0, Skipped: 0, Total: 10
```

