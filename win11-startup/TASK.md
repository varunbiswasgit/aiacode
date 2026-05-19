# Win11 Startup Launcher — Task Backlog

All tasks complete. No outstanding items.

---

## Removed (Not Achievable Portably)

| Task | Reason |
|---|---|
| **TEST-05** — Unit test `Wait-ForAppReady` | Requires mocking `Get-Process` and `Start-Sleep` in PS 5.1 without an external mock library; not reliable in a portable single-file context. |
| **TEST-06** — Unit test `Repair-ShortcutArguments` | Depends on `C:\Program Files\WindowsApps` ACL structure; not reproducible portably without admin access on a real Windows machine. |
| **TEST-07** — Unit test `Test-AppAlreadyOpen` (live MainWindowHandle) | Mocking live `MainWindowHandle` requires a real running GUI process; not suitable for headless unit tests. |
| **INT-03** — Integration test repair flow | `Repair-ShortcutTarget` auto-discovery depends on a real broken install path on the target machine; not reproducible portably. |

---

## Done

- [x] **NEW-TEST-13** — Unit test `Import-AppsConfig` schemaVersion: clean load on v1, warning on missing, warning on mismatch
- [x] **NEW-TEST-12** — Unit test `Show-AppList`: no throw, header emitted, app name appears in output
- [x] **NEW-TEST-11** — Unit test `Invoke-FailureRecovery`: choice 1 returns `$true` + calls `PreRetryAction`; choice 3/default returns `$false`; no PreRetryAction provided is safe
- [x] **NEW-TEST-10** — Unit test `Resolve-Aumid`: all-paths failure writes to `startup-error.log`
- [x] **NEW-TEST-09** — Unit test `Export-AppsConfig` error path: invalid destination does not throw
- [x] **NEW-TEST-08** — Unit test `Test-AppAlreadyOpen -RequireWindow`: non-running returns `$false`; switch accepted without throw
- [x] **QOL-05** — Add menu option [5] `List startup apps` via `Show-AppList`; renumber Exit to [6]
- [x] **QOL-03** — `schemaVersion` wrapper in `Import-AppsConfig` / `Export-AppsConfig`
- [x] **QOL-02** — `Resolve-Aumid` logs all-paths failure to `startup-error.log`
- [x] **FIX-04** — `Test-AppAlreadyOpen -RequireWindow`; `Start-Win32App` passes it for Window-mode apps
- [x] **QOL-01** — Pre-compute shortcut-existence in `Show-AppPicker`
- [x] **HARD-05** — Shortcut number 1-2 digit validation in `Add-Shortcut`
- [x] **HARD-04** — 3-attempt cap in `Prompt-ForExactExePath`
- [x] **QOL-04** — Extract `Invoke-FailureRecovery` (landed with ROB-04)
- [x] **ROB-04** — Bounded retry loop in `Start-Win32App`
- [x] **ROB-02** — `Test-Path` guard in `Edit-Shortcut`
- [x] **ROB-01** — `try/catch` on `Set-Content` in `Export-AppsConfig`
- [x] **FIX-07** — Remove stale `Get-ShortcutObject` after `Repair-ShortcutTarget`
- [x] **FIX-06** — `$env:SystemRoot` in `Initialize-Shortcut`
- [x] **FIX-05** — PFN from `Get-AppxPackage` in `Repair-ShortcutArguments`
- [x] **T-10** through **T-01** — All refactor tasks
- [x] **SEC-01** through **SEC-04** — All security gates
- [x] **HARD-01** through **HARD-03** — All hardening tasks
- [x] **TEST-01** through **TEST-04** — Original Pester scaffold and unit tests
- [x] **INT-01** / **INT-02** — Integration harness and bootstrap smoke test
- [x] **REF-01** / **REF-02** — Externalize and write-back `apps.json`
- [x] **FIX-01** through **FIX-03** — Appx add support, cancel path, phase timeout math
