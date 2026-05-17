# Win11 Startup Launcher — Task Backlog

Tasks are ordered by risk/value within each category.
Move each item to **Done** after its commit lands.

---

## To Do

### Fixes

- [ ] **FIX-01** — Repair Add-menu Appx support mismatch: the new `Add-Shortcut` flow currently accepts `LaunchType = 'Appx'` but does not prompt for or persist the Appx-specific fields used by `Resolve-Aumid` (`StartAppName`, `KnownAumid`, `AppxName`). Either remove `Appx` from the Add flow or extend the prompt/export/import path to fully support those fields.

- [ ] **FIX-02** — Restore a true Cancel path in the Add flow: the Add menu prompt says `or [0] to add a new app entry`, but `Show-AppPicker` already returns `$null` for `0`, so the same return value currently means both cancel and create-new. Split these outcomes so Add has both a real cancel action and an explicit create-new action.

- [ ] **FIX-03** — Correct phase-2 timeout math in `Wait-ForAppReady`: the function uses `Min($script:SettleSeconds, $TimeoutSeconds)` for phase 1, but subtracts the full `$script:SettleSeconds` instead of the actual phase-1 value. Store the phase-1 settle duration in a variable and subtract that exact value when calculating remaining timeout.

---

## Removed (Not Achievable Portably)

| Task | Reason |
|---|---|
| **TEST-05** — Unit test `Wait-ForAppReady` | Requires mocking `Get-Process` and `Start-Sleep` in PS 5.1 without an external mock library; not reliable in a portable single-file context. |
| **TEST-06** — Unit test `Repair-ShortcutArguments` | Depends on `C:\\Program Files\\WindowsApps` ACL structure; not reproducible portably without admin access on a real Windows machine. |
| **TEST-07** — Unit test `Test-AppAlreadyOpen` | Mocking live `MainWindowHandle` requires a real running GUI process; not suitable for headless unit tests. |
| **INT-03** — Integration test repair flow | `Repair-ShortcutTarget` auto-discovery depends on a real broken install path on the target machine; not reproducible portably. |

---

## Done

- [x] **SEC-01** — Allowlist exe repair paths: added `$script:AllowedExeRoots` config array and `Test-ExePathAllowed` helper; both `Prompt-ForExactExePath` and `Repair-ShortcutTarget` reject paths outside allowed roots. _(v10)_
- [x] **SEC-02** — Authenticode signature check: added `Test-ExeSignatureTrusted` using `Get-AuthenticodeSignature`; both `Prompt-ForExactExePath` and `Repair-ShortcutTarget` reject executables whose signature status is not `Valid`. _(v11)_
- [x] **SEC-03** — Publisher allowlist: added optional `ExpectedPublisher` field per app entry; `Test-ExeSignatureTrusted` accepts `-ExpectedPublisher` and verifies `SignerCertificate.Subject` contains the expected string when provided. Microsoft apps use `CN=Microsoft Corporation`; Chrome uses `CN=Google LLC`. _(v12)_
- [x] **SEC-04** — Process-name collision guard: `Test-AppAlreadyOpen` now accepts `ExpectedExe`, filters matching processes by `MainModule.FileName`, and returns false if no exact executable match remains; `Start-Win32App` passes `App.ExpectedExe`. _(v13)_
- [x] **HARD-01** — Anchored `Repair-ShortcutArguments` regex: replaced loose PFN extraction with a fully anchored pattern `^shell:appsFolder\\(Microsoft\.[A-Za-z0-9._]+_[A-Za-z0-9]+)![A-Za-z0-9._-]+$` that constrains PFN to start with `Microsoft.` and requires the full argument string to match, blocking path injection via crafted `ExpectedArguments` values. _(v14)_
- [x] **HARD-02** — `$script:` scope on all shared vars: prefixed `$WshShell`, `$apps`, `$startMenu`, `$MaxRepairDepth`, `$InitialDelaySeconds`, `$LaunchTimeoutSeconds`, `$PostLaunchPauseSeconds`, `$SettleSeconds`, and `$AllowedExeRoots` with `$script:` on declaration and all references; prevents variable leak or shadowing when dot-sourced by Pester. _(v15)_
- [x] **HARD-03** — Safer XML manifest loading: `Repair-ShortcutArguments` now uses `[xml]::new(); $manifest.Load($manifestPath)` instead of `[xml]$manifest = Get-Content` to handle BOMs and large manifests. _(v11)_
- [x] **TEST-01** — Pester scaffold: added `$env:PS_STARTUP_TESTMODE` guard at bottom of `Win11startup.ps1`; skips interactive menu and startup sequence when set to `'1'`. Created `Win11startup.Tests.ps1` with `BeforeAll` that sets the env var and dot-sources the script. _(v16)_
- [x] **TEST-02** — Unit test `Get-RelativeDepth`: 5 cases — same folder (0), 1 level, 2 levels, outside base (MaxValue), empty candidate via outside-base check. All in `Describe Unit > Get-RelativeDepth`. _(v16)_
- [x] **TEST-03** — Unit test `Find-MisnumberedShortcut`: 4 cases — match found, no name match, empty folder, missing folder. Uses temp `.lnk` stubs in `$env:TEMP`; cleaned up in `AfterEach`. _(v16)_
- [x] **TEST-04** — Unit test `Test-ExePathAllowed` (4 cases: Program Files, SystemRoot, TEMP reject, Desktop reject) and `Test-ExeSignatureTrusted` (4 cases: notepad valid, correct publisher, wrong publisher, fake .exe). _(v16)_
- [x] **INT-01** — Integration harness: `Describe Integration` gated by `$env:RUN_INTEGRATION -eq '1'`; `BeforeAll` creates temp dir, `AfterAll` removes all temp `.lnk` files and the dir. _(v16)_
- [x] **INT-02** — Shortcut bootstrap smoke test: creates a real `.lnk` pointing to `notepad.exe`, calls `Initialize-Shortcut`, confirms shortcut exists and `TargetPath` resolves correctly, cleans up in `AfterAll`. _(v16)_
- [x] **REF-01** — Externalize `$script:apps` to `apps.json`: inline array removed from `Win11startup.ps1`; `Import-AppsConfig` loads and validates the file on startup, failing closed on missing/malformed entries; `apps.json` added with all 8 app entries. _(v17)_
- [x] **REF-02** — Write-back new app entries to `apps.json` via Add/Delete menus: `Add-Shortcut` extended to prompt for all fields when called with no existing app, validates through security gates, creates the `.lnk`, appends to `$script:apps`, and persists via `Export-AppsConfig`; `Remove-Shortcut` extended to remove the entry from `$script:apps` and save `apps.json` after deletion. _(v17)_
