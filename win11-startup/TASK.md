# Win11 Startup Launcher — Task Backlog

Tasks are ordered by risk/value within each category.
Move each item to **Done** after its commit lands.

---

## To Do

### Hardening

- [ ] **HARD-02** — Add `$script:` scope to shared vars: prefix `$WshShell`, `$apps`, `$startMenu`, `$MaxRepairDepth`, `$InitialDelaySeconds`, `$LaunchTimeoutSeconds`, `$PostLaunchPauseSeconds`, `$SettleSeconds`, and `$AllowedExeRoots` so dot-sourced Pester runs cannot leak or shadow globals.
  > _Achievable: search-and-replace on declarations and all references._

### Unit Tests (Pester)

- [ ] **TEST-01** — Pester scaffold: create `Win11startup.Tests.ps1`; add `BeforeAll` that sets `$env:PS_STARTUP_TESTMODE = '1'` then dot-sources the script. Add a 2-line guard in the main script that skips the menu and startup sequence when the env var is set.
  > _Achievable: small change to main script plus new test file._

- [ ] **TEST-02** — Unit test `Get-RelativeDepth`: 5 cases — same folder (0), 1 level, 2 levels, path outside base (MaxValue), empty string. Pure path math; no mocking needed.
  > _Achievable: self-contained function._

- [ ] **TEST-03** — Unit test `Find-MisnumberedShortcut`: create temp `.lnk` stubs in `$env:TEMP`; verify match, no-match, empty folder, and missing folder cases. Clean up in `AfterEach`.
  > _Achievable: filesystem only, no COM or process dependency._

- [ ] **TEST-04** — Unit test `Test-ExePathAllowed` and `Test-ExeSignatureTrusted`: allowlist check is string comparison (no mock). Signature test uses `notepad.exe` for the Valid case and a renamed `.txt` for the invalid case.
  > _Achievable: no COM or process mocking needed._

### Integration Tests

- [ ] **INT-01** — Integration harness: add a `Describe 'Integration'` block gated by `$env:RUN_INTEGRATION -eq '1'`; `AfterAll` removes all temp `.lnk` files and kills any test processes.
  > _Achievable: standard Pester pattern. Required before INT-02._

- [ ] **INT-02** — Shortcut bootstrap smoke test: create a temp `.lnk` pointing to `notepad.exe` under `$env:TEMP`, call `Initialize-Shortcut`, confirm `TargetPath` resolves, then delete. No admin rights needed.
  > _Achievable: uses real COM via `$WshShell`._

### Refactor

- [ ] **REF-01** — Externalize `$apps` to `apps.json`: move the inline `$apps` table from `Win11startup.ps1` into a sibling `apps.json` config file loaded via `Get-Content | ConvertFrom-Json`. Validate required fields (`Name`, `LaunchType`, `ShortcutPath`, `ProcessName`, `ExpectedExe`) on load and fail closed if the file is missing or malformed. Allows adding new app entries without editing the script source. All existing security gates (allowlist, signature, publisher) remain unchanged.
  > _Achievable: replace the inline array with a single loader block; all downstream code is field-name-based and requires no other changes._

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

- [x] **SEC-01** — Allowlist exe repair paths: added `$AllowedExeRoots` config array and `Test-ExePathAllowed` helper; both `Prompt-ForExactExePath` and `Repair-ShortcutTarget` reject paths outside allowed roots. _(v10)_
- [x] **SEC-02** — Authenticode signature check: added `Test-ExeSignatureTrusted` using `Get-AuthenticodeSignature`; both `Prompt-ForExactExePath` and `Repair-ShortcutTarget` reject executables whose signature status is not `Valid`. _(v11)_
- [x] **SEC-03** — Publisher allowlist: added optional `ExpectedPublisher` field per app entry; `Test-ExeSignatureTrusted` accepts `-ExpectedPublisher` and verifies `SignerCertificate.Subject` contains the expected string when provided. Microsoft apps use `CN=Microsoft Corporation`; Chrome uses `CN=Google LLC`. _(v12)_
- [x] **SEC-04** — Process-name collision guard: `Test-AppAlreadyOpen` now accepts `ExpectedExe`, filters matching processes by `MainModule.FileName`, and returns false if no exact executable match remains; `Start-Win32App` passes `App.ExpectedExe`. _(v13)_
- [x] **HARD-01** — Anchored `Repair-ShortcutArguments` regex: replaced loose PFN extraction with a fully anchored pattern `^shell:appsFolder\\(Microsoft\.[A-Za-z0-9._]+_[A-Za-z0-9]+)![A-Za-z0-9._-]+$` that constrains PFN to start with `Microsoft.` and requires the full argument string to match, blocking path injection via crafted `ExpectedArguments` values. _(v14)_
- [x] **HARD-03** — Safer XML manifest loading: `Repair-ShortcutArguments` now uses `[xml]::new(); $manifest.Load($manifestPath)` instead of `[xml]$manifest = Get-Content` to handle BOMs and large manifests. _(v11)_
