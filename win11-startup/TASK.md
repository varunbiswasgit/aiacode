# Win11 Startup Launcher — Task Backlog

Tasks are ordered by risk/value within each category.
Move each item to **Done** after its commit lands.

---

## To Do

### Correctness

- [ ] **FIX-04** — `Test-AppAlreadyOpen` returns `$true` for any running process even when it has no visible window; the final `return $true` at the bottom ignores window-vs-tray distinction. Add a `$RequireWindow` switch so callers that need a visible window can enforce it.

### Hardening

- [ ] **HARD-04** — `Prompt-ForExactExePath` loops forever with no escape except an empty Enter. Add a max-attempt counter (e.g. 3) and return `$null` after exhausting retries so the caller can handle it gracefully.
- [ ] **HARD-05** — `Add-Shortcut` (new-entry flow) does not validate that the shortcut number entered by the user is numeric or padded correctly (e.g. `09`). A blank or non-numeric number produces a malformed `.lnk` filename silently. Add validation with a re-prompt.
- [ ] **HARD-06** — `Remove-Shortcut` now filters `$script:apps` by object identity (`$_ -ne $App`). No action needed unless Pester test coverage is added for the exact-match case.

### Quality / Clarity

- [ ] **QOL-01** — `Show-AppPicker` re-evaluates `Test-Path` inside `Write-Host` on every loop iteration. Extract the status string before the loop to avoid repeated filesystem calls when the list is long.
- [ ] **QOL-02** — `Resolve-Aumid` silently returns `$null` with only a `Write-Warning` when all three resolution paths fail. Log the failure to `startup-error.log` via `Write-ErrorLog` so it appears in the post-run diagnostic file.
- [ ] **QOL-03** — `apps.json` schema has no version field. Add a top-level `"schemaVersion": 1` wrapper so future breaking changes can be detected at load time in `Import-AppsConfig`.

### Testing

- [ ] **TEST-08** — Unit test `Test-AppAlreadyOpen` with `$RequireWindow` switch (once FIX-04 lands).
- [ ] **TEST-09** — Unit test `Export-AppsConfig` error path (ROB-01 landed): verify `Write-ErrorLog` is called and a warning is emitted when `Set-Content` throws.
- [ ] **TEST-10** — Unit test `Resolve-Aumid` logs to error log when all three resolution paths fail (once QOL-02 lands).
- [ ] **TEST-11** — Unit test `Invoke-FailureRecovery`: verify each branch returns the correct value, calls `Show-FailureMenu`, invokes `PreRetryAction` only on choice `'1'`, and calls `Edit-Shortcut` / skips correctly on other choices.

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
