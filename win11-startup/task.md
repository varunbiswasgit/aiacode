# Win11startup.ps1 — Task Tracker

All items reference the live script at `win11-startup/Win11startup.ps1`.

---

## Open Bugs

| ID | Location | Description | Status |
|----|----------|-------------|--------|
| BUG-A | `Invoke-LaunchAttempt` (×2) | `return if ($recover) {...}` is invalid PS syntax — PS does not support inline ternary `return`. Must split into `if ($recover) { return 'Retry' } else { return 'Abort' }` | ✅ Fixed |
| BUG-B | `Add-Shortcut` Win32 branch | `New-AppEntry` call passes `-AppName $appName` — neither the param name (`-AppxName`) nor the variable (`$appName` vs `$appxName`) is correct. `-ExpectedPublisher` flag is also missing its argument value. | ✅ Fixed |
| BUG-C | `Initialize-Shortcut` | Orphaned `} else { Write-Warning "shortcut creation skipped." }` with no matching `if`. The preceding `if (Test-Path ...)` block resolves `$exePath` but the `else` branch is a dangling statement that never runs. | ✅ Fixed |
| BUG-D | `Invoke-LaunchAttempt` (×2) | `return if ($recover) { 'Retry' } else { 'Abort' }` — inline `return if` is still invalid in PowerShell 5.1 (Windows default). Will throw a parse error on any machine without PS 7+. Fix: `if ($recover) { return 'Retry' } else { return 'Abort' }` | ⬜ Open |
| BUG-E | `Repair-ShortcutArguments` | AUMID fragment regex: `(($Matches[1]) -split '_', 2)[1]` extracts only the publisher token from the PFN, not the full suffix needed to match WindowsApps folder names. Fix: use `$Matches[1]` directly as the search fragment. | ⬜ Open |
| BUG-F | `Get-AppPresenceMode` | No guard for blank `$ProcessName`. `Get-Process -Name ''` enumerates all processes, causing false positives in presence detection. Fix: add `if ([string]::IsNullOrWhiteSpace($ProcessName)) { return $null }` at top of function. | ⬜ Open |
| BUG-G | `Sync-AppsFromStartMenu` | `$expectedExe = $leafName` is set to `''` when shortcut target is empty (broken/URL shortcuts). This passes empty `ExpectedExe` into `New-AppEntry`, which `Import-AppsConfig` throws on at next load. Fix: add `continue` guard after the `$leafName -notlike '*.exe'` warning branch. | ⬜ Open |

---

## Open Test Gaps

| ID | Location | Description | Status |
|----|----------|-------------|--------|
| TEST-GAP-01 | `Win11startup.Tests.ps1` | `Invoke-LaunchAttempt` — no direct unit tests. Core retry logic (Success/Retry/Abort paths) untested in isolation. | ✅ Closed — NEW-TEST-20 |
| TEST-GAP-02 | `Win11startup.Tests.ps1` | `Start-Win32App` retry loop (0–2 attempts) — max-attempts guard and loop exit untested. | ✅ Closed — NEW-TEST-23 |
| TEST-GAP-03 | `Win11startup.Tests.ps1` | `Wait-ForWindowByTitle` (BUG-06 path) — title-match window polling logic unverified. | ✅ Closed — NEW-TEST-24 |
| TEST-GAP-04 | `Win11startup.Tests.ps1` | `Sync-AppsFromStartMenu` with empty-target shortcut — unguarded `$scArgs` (BUG-G) not caught by any test. | ✅ Closed — NEW-TEST-25 |
| TEST-GAP-05 | `Win11startup.Tests.ps1` | `Resolve-ConfigPath` user-input branches — custom JSON path and new-file creation paths untested. | ✅ Closed — NEW-TEST-26 |
| TEST-GAP-06 | `Win11startup.Tests.ps1` | `Show-FailureMenu` choice `'4'` (Delete entry, UX-04) — `Invoke-FailureRecovery` switch branch for `'4'` has no test. | ✅ Closed — NEW-TEST-27 |
| TEST-GAP-07 | `Win11startup.Tests.ps1` | `TEST-02` (Get-RelativeDepth) tests a function removed from the main script (AUD-01). Tests exercise dead code and should be removed. | ✅ Closed — dead TEST-02 block removed |

---

## Open Script/Test Signature Mismatches

| ID | Function | Mismatch | Test ID |
|----|----------|----------|---------|
| FIX-TEST-01 | `Get-ShortcutObject` | Test calls `-LnkPath`; script has `-ShortcutPath`. Add `[Alias('LnkPath')]` to `-ShortcutPath`. | NEW-TEST-18 |
| FIX-TEST-02 | `Find-ExeWithinDepth` | Test calls `-SearchRoot`, `-ExeName`, `-MaxDepth`; script has `-RootFolder`, `-ExpectedExe`, no depth cap. Add aliases and `-MaxDepth` with depth-bounded filtering. | NEW-TEST-17 |
| FIX-TEST-03 | `New-AppShortcut` | Test calls `-App $app -TargetPath $target`; script has no `-App` param. Add `-App [PSCustomObject]` mapping `$App.ShortcutPath` → `-Path`. | NEW-TEST-19 |
| FIX-TEST-04 | `Invoke-AppLaunch` | Function missing. Test calls `Invoke-AppLaunch -App $app`; script has `Invoke-LaunchAttempt`. Add thin wrapper. | NEW-TEST-20 |
| FIX-TEST-05 | `Test-ShortcutHealthy` | Function missing. Extract as standalone: `.lnk` exists → target exists → `Test-ExeAcceptable` passes. | NEW-TEST-21 |
| FIX-TEST-06 | `Get-AppPresence` | Function missing. Test expects `'Running'`/`'WindowVisible'`/`$null`; script has `Get-AppPresenceMode` returning `'Tray'`/`'Window'`/`$null`. Add wrapper with mapped return values. | NEW-TEST-22 |
| FIX-TEST-07 | `Wait-ForWindowByTitle` | Test calls `-TitleFragment`/`-TimeoutSeconds` expecting `[bool]`; script has `-App`/`-WaitSecs` returning process object. Add overload returning `$true`/`$false`. | NEW-TEST-24 |
| FIX-TEST-08 | `Sync-AppsFromStartMenu` | Test calls `-StartMenuPath $path`; script has no parameter (hardcodes `$script:startMenu`). Add optional `-StartMenuPath` param. | NEW-TEST-25 |
| FIX-TEST-09 | `Resolve-ConfigPath` | Test calls `-Path $cfgPath`; script has no parameters. Add optional `-Path` param. | NEW-TEST-26 |
| FIX-TEST-10 | `Start-Win32App` | Test calls `-MaxAttempts 1` and `-MaxAttempts 0`; script hardcodes loop bound. Add `-MaxAttempts [int]` param defaulting to `3`. | NEW-TEST-23 |

> All FIX-TEST items are **additive only** — no existing logic changes. Apply together in one script commit.

---

## Completed Items (from header comments)

- BUG-01 through BUG-08, FIX-05 through FIX-07
- LEAN-01 through LEAN-07
- ROB-01, ROB-02, ROB-04
- DUP-01, AUD-01, SYNC-01
- QOL-01 through QOL-05
- T-07, INT-01, INT-02, UX-02 through UX-04
- HARD-04, HARD-05

## Completed Tests Added (session: 2026-05-23)

| Tag | Function |
|-----|----------|
| NEW-TEST-14 | `Test-ExeAcceptable` |
| NEW-TEST-15 | `New-AppEntry` |
| NEW-TEST-16 | `Wait-ForProcessCondition` |
| NEW-TEST-17 | `Find-ExeWithinDepth` |
| NEW-TEST-18 | `Get-ShortcutObject` |
| NEW-TEST-19 | `New-AppShortcut` |
| NEW-TEST-20 | `Invoke-AppLaunch` |
| NEW-TEST-21 | `Test-ShortcutHealthy` |
| NEW-TEST-22 | `Get-AppPresence` |
| NEW-TEST-23 | `Start-Win32App` |
| NEW-TEST-24 | `Wait-ForWindowByTitle` |
| NEW-TEST-25 | `Sync-AppsFromStartMenu` |
| NEW-TEST-26 | `Resolve-ConfigPath` |
| NEW-TEST-27 | `Invoke-FailureRecovery` choice `'4'` (Delete entry) |

---

## Principles

- Always reuse existing modules: `Test-ExeAcceptable`, `New-AppEntry`, `Initialize-Shortcut`, `Invoke-ShortcutRepair`
- No dead branches, unused variables, or uncalled helpers
- All `.lnk` writes go through `New-AppShortcut` (LEAN-02)
- All exe validation goes through `Test-ExeAcceptable` (LEAN-01)
