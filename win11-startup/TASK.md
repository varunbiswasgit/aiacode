# Win11 Startup Launcher — Task Backlog

---

## Open

| ID | Category | Description |
|---|---|---|
| **AUD-01** | Dead code | `Get-RelativeDepth` is never called in the startup sequence — only exists for Pester. Move into test file or remove; update `Find-ExeWithinDepth` comment |
| **AUD-02** | Bug | `Repair-ShortcutTarget` allowlist/signature failure inner block falls through without `return $null` after writing the warning — subsequent code may attempt a manual prompt unnecessarily |
| **AUD-03** | Fluff | `Sync-AppsFromStartMenu` sets `$launchType = 'Win32'` in both the `elseif` and the catch-all `else` branch — redundant; set once as default before the `if` block |
| **AUD-04** | Fluff | `Sync-AppsFromStartMenu` `else` catch-all duplicates `$processName` and `$expectedExe` derivation from the `elseif` Win32 branch — merge the two branches |
| **AUD-05** | Comment | `Start-Win32App` retry loop comment says "max 2 retries" but `for ($attempt = 0; $attempt -le 2)` is 3 iterations — clarify comment to say "up to 3 attempts" |
| **AUD-06** | Reusability | `Add-Shortcut` new-entry block and `Sync-AppsFromStartMenu` both construct a `[PSCustomObject]` app entry with identical field sets — extract `New-AppEntry` private helper |
| **AUD-07** | Line reduction | Minor one-liner opportunities in `Export-AppsConfig` and `Import-AppsConfig` success/warning messages — consolidate without losing clarity |
| **README-01** | Docs | README outdated: references `apps.json` and `$MaxRepairDepth`; missing menu `[6] Sync` workflow; missing all user menu scenario / workflow descriptions |

---

## Removed (Not Achievable Portably)

| Task | Reason |
|---|---|
| **TEST-05** | Mocking `Get-Process`/`Start-Sleep` not reliable without external mock library in PS 5.1. |
| **TEST-06** | Depends on `WindowsApps` ACL structure; not reproducible portably. |
| **TEST-07** | Mocking live `MainWindowHandle` requires a real GUI process. |
| **INT-03** | `Repair-ShortcutTarget` auto-discovery depends on a real broken install path. |
| **DUP-02** | Optional — `WorkingDirectory` derivation pattern; deferred by design. |
| **DEAD-01** | `Get-RelativeDepth` superseded by AUD-01 — now tracked as open. |

---

## Done

- [x] **SYNC-01** — `Sync-AppsFromStartMenu` inlined into `Win11startup.ps1`; menu `[6]`; auto-triggers on missing `Win11startupapps.json`; `Sync-AppsJson.ps1` deleted
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
