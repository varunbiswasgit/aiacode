# Win11 Startup Launcher — Task Backlog

Derived from the audit of v9 (runtime presence-mode detection). Tasks are ordered by risk/value.
Move each item to **Done** after its commit lands.

---

## To Do

### Security

- [ ] **SEC-01** — Allowlist exe repair paths: validate that a user-supplied or auto-discovered exe lives under `Program Files`, `Program Files (x86)`, or `Windows` before updating any shortcut target.
- [ ] **SEC-02** — Authenticode signature check: before persisting a repaired shortcut, call `Get-AuthenticodeSignature` and reject executables whose status is not `Valid`.
- [ ] **SEC-03** — Publisher allowlist: for each app entry, add an optional `ExpectedPublisher` field; compare the resolved exe's signer against it before writing the shortcut.
- [ ] **SEC-04** — Process-name collision guard: when `Test-AppAlreadyOpen` finds a running process, verify its executable path matches `ExpectedExe` so an unrelated same-named process cannot cause a false skip.

### Unit Tests (Pester)

- [ ] **TEST-01** — Pester scaffold: create `Win11startup.Tests.ps1`; add `BeforeAll` that dot-sources the main script with a `$TestMode` guard so the main-menu and startup sequence do not execute on import.
- [ ] **TEST-02** — Unit test `Get-RelativeDepth`: cases for same folder (depth 0), one level deep, two levels deep, path outside base (returns MaxValue), and empty input.
- [ ] **TEST-03** — Unit test `Find-MisnumberedShortcut`: matching `.lnk` with wrong number prefix found, no match when name differs, empty folder, folder missing.
- [ ] **TEST-04** — Unit test `Get-AppPresenceMode`: mock `Get-Process` to simulate window appearing on tick 2 (Window), process present but no window after settle (Tray), process never appearing ($null).
- [ ] **TEST-05** — Unit test `Wait-ForAppReady`: mock presence-mode results and verify Phase 1 / Phase 2 branching and return values.
- [ ] **TEST-06** — Unit test `Repair-ShortcutArguments`: mock filesystem with valid and ACL-blocked AppxManifest; verify AUMID reconstruction and fallback to ExpectedArguments AppId.
- [ ] **TEST-07** — Unit test `Test-AppAlreadyOpen`: process not found (false), process found no window (true), process found with window (true).

### Integration Tests

- [ ] **INT-01** — Pester integration harness: add a second describe block `Integration` gated by `$env:RUN_INTEGRATION -eq '1'`; tests create real temp `.lnk` files under `$env:TEMP` and clean up in `AfterAll`.
- [ ] **INT-02** — Integration test full Win32 shortcut bootstrap: create a valid `.lnk` pointing to `notepad.exe`, run `Start-Win32App`, confirm Notepad starts and is killed in cleanup.
- [ ] **INT-03** — Integration test repair flow: create `.lnk` with broken target, run `Start-Win32App`, confirm shortcut target is updated to discovered `notepad.exe` and process starts.

### Hardening

- [ ] **HARD-01** — Constrain `Repair-ShortcutArguments` regex: tighten the AUMID fragment extraction regex to prevent path traversal via crafted `ExpectedArguments` values.
- [ ] **HARD-02** — Add `$script:` scope guards: ensure `$WshShell`, `$apps`, and config variables are scoped to `$script:` so dot-sourced test runs cannot leak globals.
- [ ] **HARD-03** — Replace `[xml]$manifest = Get-Content` with `[xml]::new()` + `Load()` using a `FileStream` to avoid issues with BOMs and large manifests.

---

## Done

_No tasks completed yet._
