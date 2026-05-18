# Win11 Startup Launcher — Task Backlog

Tasks are ordered by risk/value within each category.
Move each item to **Done** after its commit lands.

---

## To Do

### Refactor

- [ ] **T-05** — Simplify `Resolve-Aumid` AppxPackage enumeration
  `Resolve-Aumid` calls `Get-AppxPackage` twice (KnownAumid verification +
  AppxName fallback). Collect the full package list once into `$pkgs` at the
  top of the function and filter from that variable both times, avoiding a
  second pipeline enumeration.

- [ ] **T-06** — Remove dead `Get-RelativeDepth` depth cap in `Find-ExeWithinDepth`
  `Repair-ShortcutTarget` now passes `MaxDepth 10` (effectively unlimited),
  making the `Get-RelativeDepth` filter inside `Find-ExeWithinDepth` a
  no-op. Remove the `-MaxDepth` parameter and the `Get-RelativeDepth` call;
  let `Get-ChildItem -Recurse` do unrestricted search. Keep
  `Get-RelativeDepth` in the file (used by tests) but remove the dead filter.

- [ ] **T-07** — Replace manual `$elapsed` counters with `[Diagnostics.Stopwatch]`
  `Get-AppPresenceMode` and the phase-2 loop in `Wait-ForAppReady` both
  maintain a manual `$elapsed++` counter. Replace with
  `[System.Diagnostics.Stopwatch]::StartNew()` so elapsed time reflects wall
  clock rather than iteration count (which drifts when `Get-Process` is slow).

### Tests

- [ ] **T-08** — Unit test `Get-ParentFolder` (replaces removed `Get-NearestExistingParent`)
  Add Pester tests for: null/empty input returns `$null`; parent of a
  two-level path that exists returns correct folder; parent folder does not
  exist returns `$null`.

- [ ] **T-09** — Unit test `Show-FailureMenu` output and return value
  Mock `Read-Host` to return '1', '2', '3' in turn; assert the function
  returns the correct string and that `Write-Host` was called with the
  expected lines.

- [ ] **T-10** — Update `Win11startup.Tests.ps1` header and function inventory
  After T-01 through T-09 land, update the Pester file's top-comment
  inventory to list `Get-ParentFolder` (added), remove
  `Get-NearestExistingParent` (deleted), and add `Show-FailureMenu`,
  `Invoke-AppLaunchWait`, and `New-AppShortcut` to the tested-functions table.

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

- [x] **T-04** — Extract `New-AppShortcut`; consolidate 5 `CreateShortcut` blocks _(next commit)_
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
