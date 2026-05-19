# Win11 Startup Launcher — Task Backlog

Tasks are ordered by risk/value within each category.
Move each item to **Done** after its commit lands.

---

## To Do

### Correctness

- [ ] **FIX-04** — `Test-AppAlreadyOpen` returns `$true` for any running process even when it has no visible window; the final `return $true` at the bottom ignores window-vs-tray distinction. Add a `$RequireWindow` switch so callers that need a visible window can enforce it.
- [ ] **FIX-05** — `Repair-ShortcutArguments` reconstructs the PFN by stripping the version segment with a fragile regex (`_\d+\.\d+\.\d+\.\d+_[^_]+__`). Use `$pkg.PackageFamilyName` from `Get-AppxPackage` instead — it is already the canonical PFN and never needs regex reconstruction.
- [ ] **FIX-06** — `Initialize-Shortcut` hardcodes `C:\Windows\explorer.exe` and `C:\Windows` for Appx shortcuts instead of using `$env:SystemRoot`. Replace both literals.
- [ ] **FIX-07** — `Start-Win32App` re-reads the shortcut after argument repair (`Get-ShortcutObject` is called again) but still uses the old `$shortcut` object for `WshShell.Run`. The path passed to `Run` is always `$App.ShortcutPath`, so this is benign today — but the stale `$shortcut` variable is misleading. Remove the redundant second read.

### Robustness

- [ ] **ROB-01** — `Export-AppsConfig` calls `Set-Content` with no `-ErrorAction`; a write failure (locked file, read-only share) silently discards the update. Wrap in `try/catch` and call `Write-ErrorLog` on failure.
- [ ] **ROB-02** — `Get-ShortcutObject` throws a terminating error if the shortcut is missing. Every caller already guards with `Test-Path` except `Edit-Shortcut` (argument-repair branch), which calls it without a prior existence check. Add the guard.
- [ ] **ROB-03** — `Wait-ForAppReady` calls `Get-AppPresenceMode` with `$phase1Secs` but does not propagate the actual time spent inside that call to the phase-2 remaining calculation. The Stopwatch starts *after* `Get-AppPresenceMode` returns, so phase-2 timeout is correctly based on remaining time — document this clearly with an inline comment so future edits do not break it.
- [ ] **ROB-04** — `Start-Win32App` recurses into itself on failure-menu choice `'1'` (missing shortcut) and `'1'` (launch timeout). Deep recursion is unlikely but unbounded. Replace with a `for` loop (max 2 retries) and `break`/`continue` instead.

### Hardening

- [ ] **HARD-04** — `Prompt-ForExactExePath` loops forever with no escape except an empty Enter. Add a max-attempt counter (e.g. 3) and return `$null` after exhausting retries so the caller can handle it gracefully.
- [ ] **HARD-05** — `Add-Shortcut` (new-entry flow) does not validate that the shortcut number entered by the user is numeric or padded correctly (e.g. `09`). A blank or non-numeric number produces a malformed `.lnk` filename silently. Add validation with a re-prompt.
- [ ] **HARD-06** — `Remove-Shortcut` compares by `$_.Name -ne $App.Name` when filtering `$script:apps` after deletion. If two entries share the same `Name`, both are removed. Switch to filtering by object identity (`$_ -ne $App`) or a unique key.

### Quality / Clarity

- [ ] **QOL-01** — `Show-AppPicker` re-formats the shortcut path in every loop iteration via `Test-Path` inside `Write-Host`. Extract the status string before the loop to avoid repeated filesystem calls when the list is long. Also change `$failedApps` in the startup sequence to store only `Name` + `ProcessName` strings instead of full PSCustomObjects — the summary loop only uses those two fields.
- [ ] **QOL-02** — `Resolve-Aumid` silently returns `$null` with only a `Write-Warning` when all three resolution paths fail. Log the failure to `startup-error.log` via `Write-ErrorLog` so it appears in the post-run diagnostic file.
- [ ] **QOL-03** — `apps.json` schema has no version field. Add a top-level `"schemaVersion": 1` wrapper so future breaking changes can be detected at load time in `Import-AppsConfig`.
- [ ] **QOL-04** — `Start-Win32App` contains two near-identical `switch ($failChoice)` blocks (missing-shortcut path and launch-timeout path). Extract into a single `Invoke-FailureRecovery` helper that accepts the app, context label, and an optional pre-retry action scriptblock. This shrinks `Start-Win32App`, removes the duplicated `Show-AppPicker` + `Edit-Shortcut` calls, and keeps the failure-handling logic in one place. Coordinate with ROB-04 (recursion guard) — both touch the same switch blocks and should land in the same commit.

### Testing

- [ ] **TEST-08** — Unit test `Test-AppAlreadyOpen` with `$RequireWindow` switch (once FIX-04 lands).
- [ ] **TEST-09** — Unit test `Export-AppsConfig` error path (once ROB-01 lands): verify `Write-ErrorLog` is called when `Set-Content` throws.
- [ ] **TEST-10** — Unit test `Resolve-Aumid` logs to error log when all three resolution paths fail (once QOL-02 lands).
- [ ] **TEST-11** — Unit test `Invoke-FailureRecovery` (once QOL-04 lands): verify each branch returns the correct value and calls the right helpers.

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
