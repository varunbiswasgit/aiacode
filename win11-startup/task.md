# Win11startup.ps1 — Task Tracker

All items reference the live script at `win11-startup/Win11startup.ps1`.

---

## Open Bugs

| ID | Location | Description | Status |
|----|----------|-------------|--------|
| BUG-A | `Invoke-LaunchAttempt` (×2) | `return if ($recover) {...}` is invalid PS syntax — PS does not support inline ternary `return`. Must split into `if ($recover) { return 'Retry' } else { return 'Abort' }` | ✅ Fixed |
| BUG-B | `Add-Shortcut` Win32 branch | `New-AppEntry` call passes `-AppName $appName` — neither the param name (`-AppxName`) nor the variable (`$appName` vs `$appxName`) is correct. `-ExpectedPublisher` flag is also missing its argument value. | ✅ Fixed |
| BUG-C | `Initialize-Shortcut` | Orphaned `} else { Write-Warning "shortcut creation skipped." }` with no matching `if`. The preceding `if (Test-Path ...)` block resolves `$exePath` but the `else` branch is a dangling statement that never runs. | ✅ Fixed |

---

## Completed Items (from header comments)

- BUG-01 through BUG-06, FIX-05 through FIX-07
- LEAN-01 through LEAN-07
- ROB-01, ROB-02, ROB-04
- DUP-01, AUD-01, SYNC-01
- QOL-01 through QOL-05
- T-07, INT-01, INT-02, UX-02 through UX-04
- HARD-04, HARD-05

---

## Principles

- Always reuse existing modules: `Test-ExeAcceptable`, `New-AppEntry`, `Initialize-Shortcut`, `Invoke-ShortcutRepair`
- No dead branches, unused variables, or uncalled helpers
- All `.lnk` writes go through `New-AppShortcut` (LEAN-02)
- All exe validation goes through `Test-ExeAcceptable` (LEAN-01)
