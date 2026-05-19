# Win11 Startup Launcher — Task Backlog

Tasks are ordered by severity within each category.
Move each item to **Done** after its commit lands.

---

## To Do

### Bugs
- [ ] **BUG-03** — `Start-Win32App`: `$requireWin` is computed before the app launches via `Get-AppPresenceMode -SettleSecs 0`. On a cold start the process is absent so mode is always `$null` → `$requireWin = $false`. Fix: only pass `-RequireWindow` when a prior presence mode is known. *(Medium)*
- [ ] **BUG-01** — `Test-AppAlreadyOpen`: the non-`RequireWindow` tail has a dead `if` branch — both paths return `$true`. Collapse to a single `return $true`. *(Low)*

### UX
- [ ] **UX-03** — `Start-Win32App` catch block: launch exception returns `$false` immediately with no recovery offered. Offer `Invoke-FailureRecovery` before returning. *(Medium)*
- [ ] **UX-02** — `Remove-Shortcut`: when shortcut is already missing, execution falls through to the JSON removal prompt without explicit acknowledgement. Add `return` after warning, or combine into a single prompt. *(Medium)*
- [ ] **UX-01** — Header comment block: `# - List apps` still references `[6]` — should be `[5]`. *(Low)*

### Duplication
- [ ] **DUP-01** — `Wait-ForAppReady`: phase-2 Stopwatch + poll loop duplicated in `$null` mode and `Window` mode branches. Extract a private `Wait-ForProcessCondition` helper. *(Low)*
- [ ] **DUP-02** — `New-AppShortcut -WorkingDirectory (Split-Path ...)` repeated in 4 callers. Optional: default derivation inside `New-AppShortcut`. *(Low / optional)*

### Dead Code
- [ ] **DEAD-01** — `Get-RelativeDepth` not called in production (retained for test coverage per T-06). Remove only if strict zero-dead-code policy is desired. *(Info)*

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

- [x] **INT-02** — Re-read `$shortcut` via `Get-ShortcutObject` after `Repair-ShortcutTarget` in `Start-Win32App`
- [x] **INT-01** — Replace hardcoded `C:\Program Files\WindowsApps` with `Join-Path $env:ProgramFiles 'WindowsApps'`
- [x] **BUG-02** — `Show-AppList`: bare string `'OK'`/`'MISSING'` replaces `Write-Output` inside `if` assignment
- [x] **NEW-TEST-13** — `Import-AppsConfig` schemaVersion unit tests
- [x] **NEW-TEST-12** — `Show-AppList` unit tests
- [x] **NEW-TEST-11** — `Invoke-FailureRecovery` unit tests
- [x] **NEW-TEST-10** — `Resolve-Aumid` error log unit test
- [x] **NEW-TEST-09** — `Export-AppsConfig` error path unit test
- [x] **NEW-TEST-08** — `Test-AppAlreadyOpen -RequireWindow` unit test
- [x] **QOL-05** — `Show-AppList` + menu option [5]; Exit to [6]
- [x] **QOL-03** — `schemaVersion` wrapper
- [x] **QOL-02** — `Resolve-Aumid` error log
- [x] **FIX-04** — `Test-AppAlreadyOpen -RequireWindow`
- [x] **QOL-01** — Pre-compute shortcut-existence in `Show-AppPicker`
- [x] **HARD-05** — Shortcut number validation
- [x] **HARD-04** — 3-attempt cap in `Prompt-ForExactExePath`
- [x] **QOL-04** / **ROB-04** — `Invoke-FailureRecovery` + bounded retry loop
- [x] **ROB-02** — `Test-Path` guard in `Edit-Shortcut`
- [x] **ROB-01** — `try/catch` on `Export-AppsConfig`
- [x] **FIX-07** / **FIX-06** / **FIX-05** — Stale object, SystemRoot, PFN source
- [x] **T-10** through **T-01** — All refactor tasks
- [x] **SEC-01** through **SEC-04** — All security gates
- [x] **HARD-01** through **HARD-03** — All hardening tasks
- [x] **TEST-01** through **TEST-04** — Pester scaffold and unit tests
- [x] **INT-01** / **INT-02** (original) — Integration harness and bootstrap smoke test
- [x] **REF-01** / **REF-02** — Externalize and write-back `apps.json`
- [x] **FIX-01** through **FIX-03** — Appx add, cancel path, phase timeout math
