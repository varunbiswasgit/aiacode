# Win11 Startup Manager — Task Log

Tracks all completed tasks, pending items, and identified defects for `Win11startup.ps1`.

---

## Completed Tasks

| ID | Category | Description | Status |
|----|----------|-------------|--------|
| LEAN-01 | Lean | `Test-ExeAcceptable` wraps `Test-ExePathAllowed` + `Test-ExeSignatureTrusted` into single gate. Both repair paths call `Test-ExeAcceptable` only. | ✅ Done |
| LEAN-02 | Lean | `Add-Shortcut` (new entry) delegates to `Initialize-Shortcut`; `WshShell.CreateShortcut` invoked in exactly one place (`New-AppShortcut`). | ✅ Done |
| LEAN-04 | Lean | `Invoke-LaunchAttempt` extracted from `Start-Win32App`. Encapsulates one repair+launch+wait+recovery cycle; returns `'Success'`, `'Retry'`, or `'Abort'`. | ✅ Done |
| LEAN-05 | Lean | Zero-shortcut `Write-Warning` moved inside `Sync-AppsFromStartMenu`; callers check return bool only. | ✅ Done |
| LEAN-06 | Lean | `Invoke-ShortcutRepair` owns `Get-ShortcutObject` + `.Save()` envelope. Both `Repair-ShortcutTarget` and `Repair-ShortcutArguments` are thin scriptblock delegates. `Update-ShortcutTarget` retired. | ✅ Done |
| LEAN-07 | Lean | `Start-AppxApp` intentionally has no failure recovery on timeout (by design; documented in header). | ✅ Done |
| SYNC-01 | Feature | Main menu `[6]` calls `Sync-AppsFromStartMenu`; auto-triggered on first run when JSON is missing. | ✅ Done |
| AUD-01 | Audit | `Get-RelativeDepth` removed from main script; retained in test file only. | ✅ Done |
| BUG-01 | Bug | `Test-AppAlreadyOpen` non-RequireWindow tail collapsed to single `return $true`. | ✅ Done |
| BUG-02 | Bug | `Show-AppList` uses bare string for status assignment. | ✅ Done |
| BUG-03 | Bug | `Start-Win32App` derives `$requireWin` from cached `$App.PresenceMode`. | ✅ Done |
| BUG-04 | Bug | `Prompt-ForExactExePath` max-attempts warning fixed from `$AppName:` to `${AppName}`. | ✅ Done |
| BUG-05 | Bug | `Sync-AppsFromStartMenu` now keeps `LaunchType=Win32` for `explorer.exe+shell:appsFolder` entries. | ✅ Done |
| BUG-06 | Bug | `Wait-ForWindowByTitle` replaces `Resolve-ProcessName`; back-fills `ProcessName` from matched window and persists to JSON. | ✅ Done |
| DUP-01 | Dup | `Wait-ForProcessCondition` extracted from `Wait-ForAppReady`. | ✅ Done |
| FIX-04 | Fix | `Test-AppAlreadyOpen` accepts `-RequireWindow` switch. | ✅ Done |
| FIX-05 | Fix | `Repair-ShortcutArguments` uses `$pkg.PackageFamilyName` directly from `Get-AppxPackage`. | ✅ Done |
| FIX-06 | Fix | `Initialize-Shortcut` uses `$env:SystemRoot` for Appx shortcut creation. | ✅ Done |
| FIX-07 | Fix | `Write-ErrorLog` defined before `trap` block; boot block prompts for custom JSON path when default not found. | ✅ Done |
| HARD-04 | Hardening | `Prompt-ForExactExePath` limits retries to 3 attempts. | ✅ Done |
| HARD-05 | Hardening | `Add-Shortcut` validates shortcut number is 1-2 digits. | ✅ Done |
| INT-01 | Integration | `Repair-ShortcutArguments` uses `Join-Path $env:ProgramFiles` for WindowsApps path. | ✅ Done |
| INT-02 | Integration | `Start-Win32App` re-reads shortcut after `Repair-ShortcutTarget`. | ✅ Done |
| QOL-01 | Quality | `Show-AppPicker` pre-computes shortcut existence before display loop. | ✅ Done |
| QOL-02 | Quality | `Resolve-Aumid` logs all-paths failure via `Write-ErrorLog`. | ✅ Done |
| QOL-03 | Quality | `Import-AppsConfig` checks for top-level `schemaVersion` field. | ✅ Done |
| QOL-04 | Quality | `Start-Win32App` retry loop bounded to 3 attempts (`attempt 0, 1, 2`). | ✅ Done |
| QOL-05 | Quality | Main menu `[5]` prints formatted table via `Show-AppList`. | ✅ Done |
| ROB-01 | Robustness | `Export-AppsConfig` wraps `Set-Content` in try/catch. | ✅ Done |
| ROB-02 | Robustness | `Edit-Shortcut` argument-repair branch guards `Get-ShortcutObject`. | ✅ Done |
| ROB-04 | Robustness | `Start-Win32App` uses `Invoke-FailureRecovery` with bounded for-loop. | ✅ Done |
| T-07 | Testing | `Get-AppPresenceMode` and `Wait-ForAppReady` use `[System.Diagnostics.Stopwatch]`. | ✅ Done |
| UX-02 | UX | `Remove-Shortcut` uses combined prompt when shortcut is missing. | ✅ Done |
| UX-03 | UX | `Start-Win32App` catch block calls `Invoke-FailureRecovery`. | ✅ Done |
| UX-04 | UX | `Show-FailureMenu` adds `[4] Delete entry`; `Invoke-FailureRecovery` `'4'` branch calls `Remove-Shortcut`. | ✅ Done |

---

## Open / Pending Tasks

| ID | Category | Description | Priority |
|----|----------|-------------|----------|
| BUG-07 | Bug | `Edit-Shortcut` update-existing branch calls `$script:WshShell.CreateShortcut()` + `.Save()` directly, bypassing `Invoke-ShortcutRepair`. Should route through `New-AppShortcut` (which already owns `.Save()`) or a dedicated `Invoke-ShortcutRepair` delegate. | High |
| BUG-08 | Bug | `Add-Shortcut` Win32 / no-args path calls `New-AppShortcut` outside `Initialize-Shortcut`. Breaks LEAN-02 contract: `Initialize-Shortcut` is not the sole `.lnk` creation path in this branch. Refactor to pass prompt-supplied `$exePath` into `Initialize-Shortcut` via a parameter or pre-populated app entry. | High |
| BUG-09 | Bug | `Invoke-LaunchAttempt` args-validity check uses `-notlike "*$($App.ExpectedArguments)*"` (loose wildcard). A stale AUMID whose old PFN is a substring of `ExpectedArguments` will not trigger repair. Replace with strict `-ine` equality check. | Medium |
| BUG-10 | Bug | `Invoke-LaunchAttempt` detects `$resolvedMode` and updates `$App.PresenceMode` in memory but never calls `Export-AppsConfig`. Presence mode is lost on next run. Add `Export-AppsConfig` after setting `$App.PresenceMode`. | Medium |

---

## Notes

- All LEAN tasks complete through LEAN-07. No further LEAN items identified.
- BUG-07 and BUG-08 are consistency violations against the LEAN-02/LEAN-06 contracts already in place.
- BUG-09 and BUG-10 are logic gaps in `Invoke-LaunchAttempt` introduced during the LEAN-04 extraction.
- Test file (`Win11startup.Tests.ps1`) should be reviewed for coverage of the `Invoke-ShortcutRepair` delegate pattern after BUG-07/BUG-08 are resolved.
