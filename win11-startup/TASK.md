# Win11 Startup Manager — Task Backlog

---

## Open

| ID | Category | Description |
|---|---|---|
| **BUG-A** | Bug | `Add-Shortcut` Win32 no-args branch: `New-AppEntry` called with `-ExpectedPublisher` missing value (used as switch), undefined `$appName` variable, and non-existent `-AppName` param. Must be `-ExpectedPublisher $expectedPublisher`, `-AppxName $appxName`. |
| **BUG-B** | Bug | `Initialize-Shortcut` has dangling `else` with no matching `if`. The `New-AppShortcut` + `Write-Host` lines sit outside any branch; the trailing `} else { Write-Warning …skipped }` is unanchored. Wrap in `if ($exePath) { … } else { … }`. |
| **BUG-C** | Bug | `Invoke-LaunchAttempt` uses `return if ($recover) { 'Retry' } else { 'Abort' }` — invalid PowerShell syntax. `return` exits with `$null` before `if` is evaluated, breaking the retry loop in `Start-Win32App`. Replace with explicit `if ($recover) { return 'Retry' } else { return 'Abort' }`. |
| **BUG-D** | Bug | `Win11startupapps.json` Sticky Notes entry is misconfigured as `LaunchType=Win32 / ProcessName=ONENOTE / ExpectedExe=ONENOTE.EXE`. Sticky Notes is a UWP app; it must be `LaunchType=Appx`, `ProcessName=StickyNotes`, `ExpectedExe=explorer.exe`, `KnownAumid=Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App`. Also causes OneNote false-skip (shared ProcessName). |
| **DEAD-01** | Cleanup | Audit `Win11startup.ps1` for any functions/variables defined but never called in the startup sequence (reuse existing modules; delete unused code). Cross-check against `Win11startup.Tests.ps1` to retain test-only stubs intentionally. |
| **FIX-TEST-01** | Test Fix | `Get-ShortcutObject`: test (NEW-TEST-18) calls `-LnkPath`; script has `-ShortcutPath`. Add `[Alias('LnkPath')]` to the `-ShortcutPath` parameter. |
| **FIX-TEST-02** | Test Fix | `Find-ExeWithinDepth`: test (NEW-TEST-17) calls `-SearchRoot`, `-ExeName`, `-MaxDepth`; script has `-RootFolder`, `-ExpectedExe`, no depth cap. Add aliases `-SearchRoot`/`-ExeName` and a new `-MaxDepth` parameter with depth-bounded filtering. |
| **FIX-TEST-03** | Test Fix | `New-AppShortcut`: test (NEW-TEST-19) calls `-App $app -TargetPath $target`; script has no `-App` parameter. Add `-App [PSCustomObject]` that maps `$App.ShortcutPath` → `-Path` when `-Path` is not supplied. |
| **FIX-TEST-04** | Test Fix | `Invoke-AppLaunch`: test (NEW-TEST-20) calls `Invoke-AppLaunch -App $app`; script has `Invoke-LaunchAttempt`. Add thin wrapper `Invoke-AppLaunch` that delegates to `Invoke-LaunchAttempt`. |
| **FIX-TEST-05** | Test Fix | `Test-ShortcutHealthy`: test (NEW-TEST-21) calls `Test-ShortcutHealthy -App $app`; function does not exist. Extract as standalone: check `.lnk` exists → target exists → `Test-ExeAcceptable` passes. |
| **FIX-TEST-06** | Test Fix | `Get-AppPresence`: test (NEW-TEST-22) calls `Get-AppPresence -ProcessName` expecting `'Running'`/`'WindowVisible'`/`$null`; script has `Get-AppPresenceMode` returning `'Tray'`/`'Window'`/`$null`. Add `Get-AppPresence` wrapper with mapped return values. |
| **FIX-TEST-07** | Test Fix | `Wait-ForWindowByTitle`: test (NEW-TEST-24) calls `-TitleFragment`/`-TimeoutSeconds` and expects `[bool]`; script has `-App`/`-WaitSecs` returning a process object. Add `-TitleFragment` + `-TimeoutSeconds` overload that returns `$true`/`$false`. |
| **FIX-TEST-08** | Test Fix | `Sync-AppsFromStartMenu`: test (NEW-TEST-25) calls `-StartMenuPath $path`; script has no parameter (hardcodes `$script:startMenu`). Add optional `-StartMenuPath` parameter defaulting to `$script:startMenu`. |
| **FIX-TEST-09** | Test Fix | `Resolve-ConfigPath`: test (NEW-TEST-26) calls `-Path $cfgPath`; script has no parameters. Add optional `-Path` parameter. |
| **FIX-TEST-10** | Test Fix | `Start-Win32App`: test (NEW-TEST-23) calls `-MaxAttempts 1` and `-MaxAttempts 0`; script hardcodes `$attempt -le 2`. Add `-MaxAttempts [int]` parameter with default `3`; replace hardcoded loop bound. |

## Closed

| ID | Category | Description | Resolution |
|---|---|---|---|
| **LEAN-01** | Refactor | Collapse `Test-ExePathAllowed` + `Test-ExeSignatureTrusted` into single `Test-ExeAcceptable` wrapper | Done |
| **LEAN-02** | Refactor | Remove inline `New-AppShortcut` from `Add-Shortcut`; delegate to `Initialize-Shortcut` so shortcut creation has one owner | Done |
| **LEAN-03** | Refactor | Extract `Invoke-MenuAction` to unify `Show-AppPicker` + dispatch used by main menu and `Invoke-FailureRecovery` choice `[2]` | Done — evaluated; patterns are contextually different enough that extraction adds marginal value. Intentionally deprioritised. |
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
| **BUG-07** | Bug | `Edit-Shortcut` update-existing branch bypassed `Invoke-ShortcutRepair`. Routed through `Initialize-Shortcut` to honor LEAN-02 contract. | Done |
| **BUG-08** | Bug | `Add-Shortcut` Win32 no-args path called `New-AppShortcut` outside `Initialize-Shortcut`. Refactored to pass `$exePath` into `Initialize-Shortcut`. | Done |
| **BUG-09** | Bug | `Invoke-LaunchAttempt` args-validity check used `-notlike` wildcard on `ExpectedArguments`. Replaced with strict `-ine` equality. | Done |
| **BUG-10** | Bug | `Invoke-LaunchAttempt` set `$App.PresenceMode` in memory after detection but never called `Export-AppsConfig`. Added `Export-AppsConfig` call after persisting. | Done |
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
| **README-01** | Docs | README outdated: references `apps.json` and `$MaxRepairDepth`; missing menu `[6] Sync` workflow | Done — full rewrite delivered in commit 8d05861 |

## Notes
- All LEAN tasks complete through LEAN-07.
- BUG-A/B/C are code bugs; BUG-D is a JSON config error. All four are blocking correct operation of `Add-Shortcut`, `Initialize-Shortcut`, `Invoke-LaunchAttempt`, and Sticky Notes launch.
- DEAD-01: reuse existing helpers (`Initialize-Shortcut`, `Invoke-ShortcutRepair`, `New-AppEntry`) — do not duplicate logic. Delete only functions with zero callers confirmed by grep.
- FIX-TEST-01 through FIX-TEST-10: all are **additive** — no existing logic changes. Two new functions (`Invoke-AppLaunch`, `Test-ShortcutHealthy`, `Get-AppPresence`), plus parameter aliases/additions on existing functions. Apply together in one commit.
- `Win11startup.Tests.ps1` intentionally retains `Get-RelativeDepth` stub (AUD-01); do not delete from test file.
