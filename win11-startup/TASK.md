# Win11 Startup Launcher — Task Backlog

All backlog items resolved. No open tasks.

---

## Removed (Not Achievable Portably)

| Task | Reason |
|---|---|
| **TEST-05** | Mocking `Get-Process`/`Start-Sleep` not reliable without external mock library in PS 5.1. |
| **TEST-06** | Depends on `WindowsApps` ACL structure; not reproducible portably. |
| **TEST-07** | Mocking live `MainWindowHandle` requires a real GUI process. |
| **INT-03** | `Repair-ShortcutTarget` auto-discovery depends on a real broken install path. |
| **DUP-02** | Optional — `WorkingDirectory` derivation pattern; deferred by design. |
| **DEAD-01** | `Get-RelativeDepth` retained for test coverage per T-06; acceptable as-is. |

---

## Done

- [x] **SYNC-01** — `Sync-AppsFromStartMenu` inlined into `Win11startup.ps1`; menu `[6]`; auto-triggers on missing `Win11startupapps.json`; `Sync-AppsJson.ps1` retired
- [x] Config file renamed from `apps.json` to `Win11startupapps.json` throughout
- [x] **BUG-01** — `Test-AppAlreadyOpen` dead branch removed
- [x] **UX-01** — Header comment menu references corrected
- [x] **DUP-01** — `Wait-ForProcessCondition` extracted
- [x] **UX-03** — catch block recovery in `Start-Win32App`
- [x] **UX-02** — `Remove-Shortcut` combined missing-shortcut prompt
- [x] **BUG-03** — `$requireWin` from cached `$App.PresenceMode`
- [x] **INT-02** / **INT-01** — Stale shortcut re-read; WindowsApps path fix
- [x] **BUG-02** — `Show-AppList` bare string status
- [x] **QOL-05** / **QOL-03** / **QOL-02** / **QOL-01** — Show-AppList, schemaVersion, AUMID log, picker perf
- [x] **FIX-04** through **FIX-07** — RequireWindow, SystemRoot, PFN source, stale object
- [x] **ROB-01** / **ROB-02** / **ROB-04** — Export guard, Edit guard, bounded retry
- [x] **HARD-04** / **HARD-05** — Path attempts cap, shortcut number validation
- [x] **T-10** through **T-01** — All refactor tasks
- [x] **SEC-01** through **SEC-04** — All security gates
- [x] **HARD-01** through **HARD-03** — All hardening tasks
- [x] **TEST-01** through **TEST-04** — Pester scaffold and unit tests
- [x] **REF-01** / **REF-02** — Externalize and write-back config
- [x] **FIX-01** through **FIX-03** — Appx add, cancel path, phase timeout math
