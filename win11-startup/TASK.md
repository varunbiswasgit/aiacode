# Win11 Startup Launcher — Task Backlog

Tasks are ordered by risk/value within each category.
Move each item to **Done** after its commit lands.

---

## To Do

### Testing

- [ ] **TEST-08** — Unit test `Test-AppAlreadyOpen` with `-RequireWindow` switch (FIX-04).
- [ ] **TEST-09** — Unit test `Export-AppsConfig` error path: verify `Write-ErrorLog` is called and a warning is emitted when `Set-Content` throws.
- [ ] **TEST-10** — Unit test `Resolve-Aumid` logs to error log when all three resolution paths fail (QOL-02).
- [ ] **TEST-11** — Unit test `Invoke-FailureRecovery`: verify each branch returns the correct value, calls `Show-FailureMenu`, invokes `PreRetryAction` only on choice `'1'`, and calls `Edit-Shortcut` / skips correctly on other choices.
- [ ] **TEST-12** — Unit test `Show-AppList`: verify header and one row of output are emitted for a single-entry `$script:apps`.
- [ ] **TEST-13** — Unit test `Import-AppsConfig` schemaVersion branch: verify warning on missing version, warning on wrong version, and clean load on correct version (QOL-03).

---

## Removed (Not Achievable Portably)

| Task | Reason |
|---|---|
| **TEST-05** — Unit test `Wait-ForAppReady` | Requires mocking `Get-Process` and `Start-Sleep` in PS 5.1 without an external mock library; not reliable in a portable single-file context. |
| **TEST-06** — Unit test `Repair-ShortcutArguments` | Depends on `C:\Program Files\WindowsApps` ACL structure; not reproducible portably without admin access on a real Windows machine. |
| **TEST-07** — Unit test `Test-AppAlreadyOpen` | Mocking live `MainWindowHandle` requires a real running GUI process; not suitable for headless unit tests. |
| **INT-03** — Integration test repair flow | `Repair-ShortcutTarget` auto-discovery depends on a real broken install path on the target machine; not reproducible portably. |

---

## Done

- [x] **QOL-05** — Add menu option [5] `List startup apps` — prints formatted table (Name, Type, Shortcut status, Process) via `Show-AppList` then exits
- [x] **QOL-03** — `Import-AppsConfig` checks `schemaVersion`; warns if absent or mismatched. `Export-AppsConfig` writes `{schemaVersion:1, apps:[...]}` wrapper
- [x] **QOL-02** — `Resolve-Aumid` logs all-paths failure to `startup-error.log` via `Write-ErrorLog`
- [x] **FIX-04** — `Test-AppAlreadyOpen` gains `-RequireWindow` switch; `Start-Win32App` passes it for Window-mode apps
- [x] **QOL-01** — Pre-compute shortcut-existence status in `Show-AppPicker` before display loop; eliminate repeated `Test-Path` calls per input attempt
- [x] **HARD-05** — Validate shortcut number is 1-2 digits in `Add-Shortcut`; re-prompt on blank or non-numeric input
- [x] **HARD-04** — Add 3-attempt cap to `Prompt-ForExactExePath`; return `$null` after exhaustion
- [x] **QOL-04** — Extract `Invoke-FailureRecovery`; eliminate duplicated switch blocks in `Start-Win32App` (landed with ROB-04)
- [x] **ROB-04** — Replace unbounded self-recursion in `Start-Win32App` with bounded for-loop (max 2 retries) via `Invoke-FailureRecovery`
- [x] **ROB-02** — Guard `Edit-Shortcut` argument-repair branch with `Test-Path` before `Get-ShortcutObject`
- [x] **ROB-01** — Wrap `Export-AppsConfig` `Set-Content` in `try/catch`; log and warn on write failure
- [x] **FIX-07** — Remove redundant `Get-ShortcutObject` call after `Repair-ShortcutTarget` in `Start-Win32App`
- [x] **FIX-06** — Replace hardcoded `C:\Windows` with `$env:SystemRoot` in `Initialize-Shortcut`
- [x] **FIX-05** — Use `$pkg.PackageFamilyName` from `Get-AppxPackage` in `Repair-ShortcutArguments`; remove fragile regex PFN reconstruction
- [x] **T-10** — Update Pester header/inventory after all refactor tasks
- [x] **T-09** — Unit test `Show-FailureMenu` output and return value
- [x] **T-08** — Unit test `Get-ParentFolder`
- [x] **T-07** — Replace manual `$elapsed` counters with `Stopwatch`
- [x] **T-06** — Remove dead `MaxDepth` filter in `Find-ExeWithinDepth`
- [x] **T-05** — Single `Get-AppxPackage` pass in `Resolve-Aumid`
- [x] **T-04** — Extract `New-AppShortcut`; consolidate 5 `CreateShortcut` blocks
- [x] **T-03** — Extract `Invoke-AppLaunchWait`; unify launch-wait tail
- [x] **T-02** — Extract `Show-FailureMenu` helper
- [x] **T-01** — Replace `Get-NearestExistingParent` with `Get-ParentFolder`; remove `$MaxRepairDepth`
- [x] **SEC-01** — Allowlist exe repair paths
- [x] **SEC-02** — Authenticode signature check
- [x] **SEC-03** — Publisher allowlist
- [x] **SEC-04** — Process-name collision guard
- [x] **HARD-01** — Anchored `Repair-ShortcutArguments` regex
- [x] **HARD-02** — `$script:` scope on all shared vars
- [x] **HARD-03** — Safer XML manifest loading
- [x] **TEST-01** — Pester scaffold
- [x] **TEST-02** — Unit test `Get-RelativeDepth`
- [x] **TEST-03** — Unit test `Find-MisnumberedShortcut`
- [x] **TEST-04** — Unit test `Test-ExePathAllowed` + `Test-ExeSignatureTrusted`
- [x] **INT-01** — Integration harness
- [x] **INT-02** — Shortcut bootstrap smoke test
- [x] **REF-01** — Externalize `$script:apps` to `apps.json`
- [x] **REF-02** — Write-back Add/Delete to `apps.json`
- [x] **FIX-01** — Add-menu Appx support
- [x] **FIX-02** — Add-flow cancel path
- [x] **FIX-03** — Phase-2 timeout math
