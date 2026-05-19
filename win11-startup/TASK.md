# Win11 Startup Launcher — Task Backlog

Tasks are ordered by risk/value within each category.
Move each item to **Done** after its commit lands.

---

## To Do

_All tasks complete._

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

- [x] **T-10** — Update Pester header/inventory after all refactor tasks _(this commit)_
- [x] **T-09** — Unit test `Show-FailureMenu` output and return value _(this commit)_
- [x] **T-08** — Unit test `Get-ParentFolder` _(this commit)_
- [x] **T-07** — Replace manual `$elapsed` counters with `Stopwatch` _(this commit)_
- [x] **T-06** — Remove dead `MaxDepth` filter in `Find-ExeWithinDepth` _(this commit)_
- [x] **T-05** — Single `Get-AppxPackage` pass in `Resolve-Aumid` _(349b44e)_
- [x] **T-04** — Extract `New-AppShortcut`; consolidate 5 `CreateShortcut` blocks _(c0bcc98)_
- [x] **T-03** — Extract `Invoke-AppLaunchWait`; unify launch-wait tail _(c156a5c)_
- [x] **T-02** — Extract `Show-FailureMenu` helper _(795e918)_
- [x] **T-01** — Replace `Get-NearestExistingParent` with `Get-ParentFolder`; remove `$MaxRepairDepth` _(76cd6b9)_
- [x] **SEC-01** — Allowlist exe repair paths _(v10)_
- [x] **SEC-02** — Authenticode signature check _(v11)_
- [x] **SEC-03** — Publisher allowlist _(v12)_
- [x] **SEC-04** — Process-name collision guard _(v13)_
- [x] **HARD-01** — Anchored `Repair-ShortcutArguments` regex _(v14)_
- [x] **HARD-02** — `$script:` scope on all shared vars _(v15)_
- [x] **HARD-03** — Safer XML manifest loading _(v11)_
- [x] **TEST-01** — Pester scaffold _(v16)_
- [x] **TEST-02** — Unit test `Get-RelativeDepth` _(v16)_
- [x] **TEST-03** — Unit test `Find-MisnumberedShortcut` _(v16)_
- [x] **TEST-04** — Unit test `Test-ExePathAllowed` + `Test-ExeSignatureTrusted` _(v16)_
- [x] **INT-01** — Integration harness _(v16)_
- [x] **INT-02** — Shortcut bootstrap smoke test _(v16)_
- [x] **REF-01** — Externalize `$script:apps` to `apps.json` _(v17)_
- [x] **REF-02** — Write-back Add/Delete to `apps.json` _(v17)_
- [x] **FIX-01** — Add-menu Appx support _(v18)_
- [x] **FIX-02** — Add-flow cancel path _(v18)_
- [x] **FIX-03** — Phase-2 timeout math _(v18)_
