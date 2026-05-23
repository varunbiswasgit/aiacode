# Win11 Startup Manager — Task Backlog

---

## Open

| ID | Category | Description | Priority |
|---|---|---|---|
| **README-01** | Docs | README outdated: references `apps.json` and `$MaxRepairDepth`; missing menu `[6] Sync` workflow; missing all user menu scenario / workflow descriptions | Low |
| **LEAN-03** | Refactor | Extract `Invoke-MenuAction` to unify `Show-AppPicker` + dispatch used by main menu and `Invoke-FailureRecovery` choice `[2]` | Low |

## Closed

| ID | Category | Description | Resolution |
|---|---|---|---|
| **LEAN-01** | Refactor | Collapse `Test-ExePathAllowed` + `Test-ExeSignatureTrusted` into single `Test-ExeAcceptable` wrapper | Done |
| **LEAN-02** | Refactor | Remove inline `New-AppShortcut` from `Add-Shortcut`; delegate to `Initialize-Shortcut` so shortcut creation has one owner | Done |
| **LEAN-04** | Refactor | Extract `Invoke-LaunchAttempt` from `Start-Win32App`; collapse timeout + exception `Invoke-FailureRecovery` calls into one | Done — returns `'Success'`/`'Retry'`/`'Abort'`; `Start-Win32App` loop is a thin `switch` |
| **LEAN-05** | Refactor | Move zero-shortcut `Write-Warning` inside `Sync-AppsFromStartMenu`; callers check bool only | Done |
| **LEAN-06** | Refactor | `Invoke-ShortcutRepair` owns `Get-ShortcutObject` + `.Save()` envelope; both repair functions are thin scriptblock delegates; `Update-ShortcutTarget` retired | Done |
| **LEAN-07** | Design | Decide and document whether `Start-AppxApp` gets failure recovery on timeout | Done — intentional asymmetry documented in header and inline comment above `Start-AppxApp` |
| **SYNC-01** | Feature | Main menu `[6]` calls `Sync-AppsFromStartMenu`; auto-triggered on first run when JSON is missing | Done |
| **AUD-01** | Audit | `Get-RelativeDepth` removed from main script; retained in test file only | Done |
| **BUG-01** | Bug | `Test-AppAlreadyOpen` non-RequireWindow tail collapsed to single `return $true` | Done |
| **BUG-02** | Bug | `Show-AppList` uses bare string for status assignment | Done |
| **BUG-03** | Bug | `Start-Win32App` derives `$requireWin` from cached `$App.PresenceMode` | Done |
| **BUG-04** | Bug | `Prompt-ForExactExePath` max-attempts warning fixed from `$AppName:` to `${AppName}` | Done |
| **BUG-05** | Bug | `Sync-AppsFromStartMenu` now keeps `LaunchType=Win32` for `explorer.exe+shell:appsFolder` entries | Done |
| **BUG-06** | Bug | `Wait-ForWindowByTitle` replaces `Resolve-ProcessName`; back-fills `ProcessName` from matched window and persists to JSON | Done |
| **BUG-07** | Bug | `Edit-Shortcut` update-existing branch bypassed `Invoke-ShortcutRepair`. Routed through `Initialize-Shortcut` to honor LEAN-02 contract. | Done — sees BUG-07 fix commit |
| **BUG-08** | Bug | `Add-Shortcut` Win32 no-args path called `New-AppShortcut` outside `Initialize-Shortcut`. Refactored to pass `$exePath` into `Initialize-Shortcut`. | Done — sees BUG-08 fix commit |
| **BUG-09** | Bug | `Invoke-LaunchAttempt` args-validity check used `-notlike` wildcard on `ExpectedArguments`. Replaced with strict `-ine` equality. | Done — sees BUG-09 fix commit |
| **BUG-10** | Bug | `Invoke-LaunchAttempt` set `$App.PresenceMode` in memory after detection but never called `Export-AppsConfig`. Added `Export-AppsConfig` call after persisting. | Done — sees BUG-10 fix commit |
| **DUP-01** | Refactor | `Wait-ForProcessCondition` extracted from `Wait-ForAppReady` | Done |
| **FIX-04** | Fix | `Test-AppAlreadyOpen` accepts `-RequireWindow` switch | Done |
| **FIX-05** | Fix | `Repair-ShortcutArguments` uses `$pkg.PackageFamilyName` directly from `Get-AppxPackage` | Done |
| **FIX-06** | Fix | `Initialize-Shortcut` uses `$env:SystemRoot` for Appx shortcut creation | Done |
| **FIX-07** | Fix | `Write-ErrorLog` defined before `trap` block; boot block prompts for custom JSON path when default not found | Done |
| **HARD-04** | Hardening | `Prompt-ForExactExePath` limits retries to 3 attempts | Done |
| **HARD-05** | Hardening | `Add-Shortcut` validates shortcut number is 1-2 digits | Done |
| **INT-01** | Integration | `Repair-ShortcutArguments` uses `Join-Path $env:ProgramFiles` for WindowsApps path | Done |
| **INT-02** | Integration | `Start-Win32App` re-reads shortcut after `Repair-ShortcutTarget` | Done |
| **QOL-01** | Quality | `Show-AppPicker` pre-computes shortcut existence before display loop | Done |
| **QOL-02** | Quality | `Resolve-Aumid` logs all-paths failure via `Write-ErrorLog` | Done |
| **QOL-03** | Quality | `Import-AppsConfig` checks for top-level `schemaVersion` field | Done |
| **QOL-04** | Quality | `Start-Win32App` retry loop bounded to 3 attempts | Done |
| **QOL-05** | Quality | Main menu `[5]` prints formatted table via `Show-AppList` | Done |
| **ROB-01** | Robustness | `Export-AppsConfig` wraps `Set-Content` in try/catch | Done |
| **ROB-02** | Robustness | `Edit-Shortcut` argument-repair branch guards `Get-ShortcutObject` | Done |
| **ROB-04** | Robustness | `Start-Win32App` uses `Invoke-FailureRecovery` with bounded for-loop | Done |
| **T-07** | Testing | `Get-AppPresenceMode` and `Wait-ForAppReady` use `[System.Diagnostics.Stopwatch]` | Done |
| **UX-02** | UX | `Remove-Shortcut` uses combined prompt when shortcut is missing | Done |
| **UX-03** | UX | `Start-Win32App` catch block calls `Invoke-FailureRecovery` | Done |
| **UX-04** | UX | `Show-FailureMenu` adds `[4] Delete entry`; `Invoke-FailureRecovery` `'4'` branch calls `Remove-Shortcut` | Done |

## Notes
- All LEAN tasks complete through LEAN-07 (LEAN-03 deprioritised — low impact).
- BUG-07 through BUG-10 resolved. LEAN-02/LEAN-06 contracts now consistently enforced across `Edit-Shortcut` and `Add-Shortcut`.
- `Invoke-LaunchAttempt` args check is strict (`-ine`); `PresenceMode` persistence triggers `Export-AppsConfig`.
- `Win11startup.Tests.ps1` should be reviewed for `Invoke-ShortcutRepair` delegate coverage.
