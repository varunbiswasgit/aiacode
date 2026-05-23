# Win11startup.ps1 — Task Tracker

All items reference the live script at `win11-startup/Win11startup.ps1`.

---

## Open Bugs

| ID | Location | Description | Status |
|----|----------|-------------|--------|
| BUG-A | `Invoke-LaunchAttempt` (×2) | `return if ($recover) {...}` is invalid PS syntax — PS does not support inline ternary `return`. Must split into `if ($recover) { return 'Retry' } else { return 'Abort' }` | ✅ Fixed |
| BUG-B | `Add-Shortcut` Win32 branch | `New-AppEntry` call passes `-AppName $appName` — neither the param name (`-AppxName`) nor the variable (`$appName` vs `$appxName`) is correct. `-ExpectedPublisher` flag is also missing its argument value. | ✅ Fixed |
| BUG-C | `Initialize-Shortcut` | Orphaned `} else { Write-Warning \"shortcut creation skipped.\" }` with no matching `if`. The preceding `if (Test-Path ...)` block resolves `$exePath` but the `else` branch is a dangling statement that never runs. | ✅ Fixed |
| BUG-D | `Invoke-LaunchAttempt` (×2) | `return if ($recover) { 'Retry' } else { 'Abort' }` — inline `return if` is still invalid in PowerShell 5.1 (Windows default). Will throw a parse error on any machine without PS 7+. Fix: `if ($recover) { return 'Retry' } else { return 'Abort' }` | ⬜ Open |
| BUG-E | `Repair-ShortcutArguments` | AUMID fragment regex: `(($Matches[1]) -split '_', 2)[1]` extracts only the publisher token from the PFN, not the full suffix needed to match WindowsApps folder names. Fix: use `$Matches[1]` directly as the search fragment. | ⬜ Open |
| BUG-F | `Get-AppPresenceMode` | No guard for blank `$ProcessName`. `Get-Process -Name ''` enumerates all processes, causing false positives in presence detection. Fix: add `if ([string]::IsNullOrWhiteSpace($ProcessName)) { return $null }` at top of function. | ⬜ Open |
| BUG-G | `Sync-AppsFromStartMenu` | `$expectedExe = $leafName` is set to `''` when shortcut target is empty (broken/URL shortcuts). This passes empty `ExpectedExe` into `New-AppEntry`, which `Import-AppsConfig` throws on at next load. Fix: add `continue` guard after the `$leafName -notlike '*.exe'` warning branch. | ⬜ Open |

---

## Open Test Gaps

| ID | Location | Description | Status |
|----|----------|-------------|--------|
| TEST-GAP-01 | `Win11startup.Tests.ps1` | `Invoke-LaunchAttempt` — no direct unit tests. Core retry logic (Success/Retry/Abort paths) untested in isolation. | ⬜ Open |
| TEST-GAP-02 | `Win11startup.Tests.ps1` | `Start-Win32App` retry loop (0–2 attempts) — max-attempts guard and loop exit untested. | ⬜ Open |
| TEST-GAP-03 | `Win11startup.Tests.ps1` | `Wait-ForWindowByTitle` (BUG-06 path) — title-match window polling logic unverified. | ⬜ Open |
| TEST-GAP-04 | `Win11startup.Tests.ps1` | `Sync-AppsFromStartMenu` with empty-target shortcut — unguarded `$scArgs` (BUG-G) not caught by any test. | ⬜ Open |
| TEST-GAP-05 | `Win11startup.Tests.ps1` | `Resolve-ConfigPath` user-input branches — custom JSON path and new-file creation paths untested. | ⬜ Open |
| TEST-GAP-06 | `Win11startup.Tests.ps1` | `Show-FailureMenu` choice `'4'` (Delete entry, UX-04) — `Invoke-FailureRecovery` switch branch for `'4'` has no test. | ⬜ Open |
| TEST-GAP-07 | `Win11startup.Tests.ps1` | `TEST-02` (Get-RelativeDepth) tests a function removed from the main script (AUD-01). Tests exercise dead code and should be removed. | ⬜ Open |

---

## Completed Items (from header comments)

- BUG-01 through BUG-08, FIX-05 through FIX-07
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
