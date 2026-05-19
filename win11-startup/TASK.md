# Win11 Startup Launcher — Task Backlog

Tasks are ordered by severity within each category.
Move each item to **Done** after its commit lands.

---

## To Do

### Bugs
- [ ] **BUG-02** — `Show-AppList`: replace `Write-Output 'OK'/'MISSING'` inside `if` assignment with bare string values `'OK'`/`'MISSING'`. `Write-Output` sends to pipeline, not to `$status`. *(Medium)*
- [ ] **BUG-03** — `Start-Win32App`: `$requireWin` is computed before the app launches via `Get-AppPresenceMode -SettleSecs 0`. On a cold start the process is absent so mode is always `$null` → `$requireWin = $false`. Fix: only pass `-RequireWindow` when a prior presence mode is known (i.e. skip the pre-launch mode check; derive from `Wait-ForAppReady` result instead). *(Medium)*
- [ ] **BUG-01** — `Test-AppAlreadyOpen`: the non-`RequireWindow` tail has a dead `if` branch — both paths return `$true`. Collapse to a single `return $true`. *(Low)*

### Integration
- [ ] **INT-02** — `Start-Win32App`: `$shortcut` object is stale after `Repair-ShortcutTarget` succeeds. The in-memory object still holds the old (broken) arguments, which can incorrectly trigger a second `Repair-ShortcutArguments` call. Fix: re-read `$shortcut` via `Get-ShortcutObject` after a successful repair. *(High)*
- [ ] **INT-01** — `Repair-ShortcutArguments`: hardcoded `"C:\Program Files\WindowsApps"` path. Replace with `Join-Path $env:ProgramFiles 'WindowsApps'` for portability on non-English Windows installs. *(Medium)*

### UX
- [ ] **UX-03** — `Start-Win32App` catch block: launch exception returns `$false` immediately with no recovery offered. Offer `Invoke-FailureRecovery` before returning, consistent with the other failure paths. *(Medium)*
- [ ] **UX-02** — `Remove-Shortcut`: when the shortcut file is already missing, a warning is printed but execution falls through to the "remove from apps.json?" prompt with no explicit user acknowledgement. Add a `return` after the warning, or ask "shortcut not found — still remove from apps.json?" as a single combined prompt. *(Medium)*
- [ ] **UX-01** — Header comment block: `# - List apps` still references `[6]` — should be `[5]`. *(Low)*

### Duplication
- [ ] **DUP-01** — `Wait-ForAppReady`: phase-2 Stopwatch + poll loop is duplicated in the `$null` mode branch and the `Window` mode branch (~12 lines). Extract a private `Wait-ForProcessCondition` helper. *(Low)*
- [ ] **DUP-02** — `New-AppShortcut -WorkingDirectory (Split-Path -Path $exePath -Parent)` repeated in 4 callers. Optional: default the `WorkingDirectory` derivation inside `New-AppShortcut` when omitted and `TargetPath` is a file. *(Low / optional)*

### Dead Code
- [ ] **DEAD-01** — `Get-RelativeDepth` is no longer called in production code (retained for test coverage per T-06). Acceptable as-is; remove only if a strict zero-dead-code policy is desired. *(Info)*

---

## Removed (Not Achievable Portably)

| Task | Reason |
|---|---|
| **TEST-05** | Mocking `Get-Process`/`Start-Sleep` not reliable without external mock library in PS 5.1. |
| **TEST-06** | Depends on `WindowsApps` ACL structure; not reproducible portably. |
| **TEST-07** | Mocking live `MainWindowHandle` requires a real GUI process. |
| **INT-03** | `Repair-ShortcutTarget` auto-discovery depends on a real broken install path. |

---

## Done

- [x] **NEW-TEST-13** — `Import-AppsConfig` schemaVersion unit tests
- [x] **NEW-TEST-12** — `Show-AppList` unit tests
- [x] **NEW-TEST-11** — `Invoke-FailureRecovery` unit tests
- [x] **NEW-TEST-10** — `Resolve-Aumid` error log unit test
- [x] **NEW-TEST-09** — `Export-AppsConfig` error path unit test
- [x] **NEW-TEST-08** — `Test-AppAlreadyOpen -RequireWindow` unit test
- [x] **QOL-05** — `Show-AppList` + menu option [5] List startup apps; Exit renumbered to [6]
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
- [x] **INT-01** / **INT-02** (original) — Integration harness and bootstrap smoke test
- [x] **REF-01** / **REF-02** — Externalize and write-back `apps.json`
- [x] **FIX-01** through **FIX-03** — Appx add support, cancel path, phase timeout math
