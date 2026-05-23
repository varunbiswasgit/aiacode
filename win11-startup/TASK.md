# Win11 Startup Launcher — Task Backlog

---

## Open

| ID | Category | Description |
|---|---|---|
| **README-01** | Docs | README outdated: references `apps.json` and `$MaxRepairDepth`; missing menu `[6] Sync` workflow; missing all user menu scenario / workflow descriptions |
| **LEAN-01** | Refactor | Collapse `Test-ExePathAllowed` + `Test-ExeSignatureTrusted` into single `Test-ExeAcceptable` wrapper |
| **LEAN-02** | Refactor | Remove inline `New-AppShortcut` from `Add-Shortcut`; delegate to `Initialize-Shortcut` so shortcut creation has one owner |
| **LEAN-03** | Refactor | Extract `Invoke-MenuAction` to unify `Show-AppPicker` + dispatch used by main menu and `Invoke-FailureRecovery` choice `[2]` |
| **LEAN-04** | Refactor | Extract `Invoke-LaunchAttempt` from `Start-Win32App`; collapse timeout + exception `Invoke-FailureRecovery` calls into one |
| **LEAN-05** | Refactor | Move zero-shortcut `Write-Warning` inside `Sync-AppsFromStartMenu`; callers check bool only |
| **LEAN-06** | Refactor | Wrap `Get-ShortcutObject` + `.Save()` envelope in `Invoke-ShortcutRepair`; `Repair-ShortcutTarget` and `Repair-ShortcutArguments` become thin scriptblock delegates |
| **LEAN-07** | Design | Decide and document whether `Start-AppxApp` gets failure recovery on timeout (Option A: add `Invoke-FailureRecovery`; Option B: by design asymmetry) |

---

## LEAN Task Details

### LEAN-01 — Collapse exe validation into `Test-ExeAcceptable`

**Scope:** `Test-ExePathAllowed`, `Test-ExeSignatureTrusted`, `Repair-ShortcutTarget`, `Prompt-ForExactExePath`

**Rationale:**  
`Test-ExePathAllowed` and `Test-ExeSignatureTrusted` are always called together in the same order at two call sites. The two-function sequence is a hidden contract every future caller must know about.

**Change:**  
Create `Test-ExeAcceptable -ExePath -ExpectedPublisher`. Calls both helpers internally; returns `$true` only when both pass. Replace both two-call sequences with a single `Test-ExeAcceptable` call.

**Acceptance criteria:**
- [ ] `Test-ExePathAllowed` and `Test-ExeSignatureTrusted` have no callers outside `Test-ExeAcceptable`
- [ ] `Repair-ShortcutTarget` calls `Test-ExeAcceptable` once
- [ ] `Prompt-ForExactExePath` calls `Test-ExeAcceptable` once
- [ ] Existing Pester tests updated to target `Test-ExeAcceptable`

---

### LEAN-02 — Unify shortcut creation in `Initialize-Shortcut`

**Scope:** `Add-Shortcut` (new entry branch), `Initialize-Shortcut`

**Rationale:**  
The new-entry branch in `Add-Shortcut` duplicates the `New-AppShortcut` call logic already in `Initialize-Shortcut`. When `Initialize-Shortcut` changes, `Add-Shortcut` must also change — a maintenance trap.

**Change:**  
After `Add-Shortcut` appends `$newEntry` to `$script:apps`, call `Initialize-Shortcut -App $newEntry`. Remove the inline `New-AppShortcut` block from `Add-Shortcut`.

**Acceptance criteria:**
- [ ] `New-AppShortcut` not called directly from `Add-Shortcut`
- [ ] `Initialize-Shortcut` is the sole `.lnk` creation path
- [ ] Win32 exe-path and Win32-with-args entries written correctly

---

### LEAN-03 — Extract `Invoke-MenuAction` to unify picker + dispatch

**Scope:** Main menu `[2]`/`[3]`/`[4]`, `Invoke-FailureRecovery` choice `[2]`

**Rationale:**  
`Show-AppPicker` followed by an action dispatch appears five times. The prompt string is the only variable across all five sites.

**Change:**  
Create `Invoke-MenuAction -Action <'Add'|'Delete'|'Modify'> -Prompt <string>`. Handles picker call, null-cancel guard, and dispatch. Replace all five call sites.

**Acceptance criteria:**
- [ ] `Show-AppPicker` not called directly from main switch or `Invoke-FailureRecovery`
- [ ] All three action types reachable through `Invoke-MenuAction`
- [ ] Cancel handled centrally
- [ ] Failure recovery choice `[2]` behaviour unchanged

---

### LEAN-04 — Consolidate `Invoke-FailureRecovery` call sites in `Start-Win32App`

**Scope:** `Start-Win32App` (timeout path, exception path)

**Rationale:**  
Timeout and exception paths both call `Invoke-FailureRecovery` with `Edit-Shortcut` as `PreRetryAction`; only `Context` string differs. Retry logic is scattered across three sites in one function.

**Change:**  
Extract launch-attempt body into `Invoke-LaunchAttempt -App` returning tri-state `'Success'`/`'Retry'`/`'Abort'`. The two `Edit-Shortcut` recovery calls collapse into one inside `Invoke-LaunchAttempt`.

**Acceptance criteria:**
- [ ] `Invoke-FailureRecovery` called at most twice in launch flow
- [ ] Retry loop bound (max 3 attempts) preserved
- [ ] `Start-Win32App` body shortened by ≥30 lines

---

### LEAN-05 — Move post-sync guard into `Sync-AppsFromStartMenu`

**Scope:** `Sync-AppsFromStartMenu`, `Resolve-ConfigPath`, main menu `[6]`

**Rationale:**  
Both callers check `$false` return and print their own inconsistent warning. A third caller would require a third copy.

**Change:**  
Add single `Write-Warning` for the zero-files case inside `Sync-AppsFromStartMenu`. Callers check bool and exit/return — no inline message text.

**Acceptance criteria:**
- [ ] Warning appears exactly once, from inside `Sync-AppsFromStartMenu`
- [ ] Both callers act on return bool without duplicating message

---

### LEAN-06 — Wrap shortcut read/write envelope in repair functions

**Scope:** `Repair-ShortcutTarget`, `Repair-ShortcutArguments`

**Rationale:**  
Both repair functions share identical `Get-ShortcutObject` + `.Save()` envelope; only middle logic differs.

**Change:**  
Create `Invoke-ShortcutRepair -App -RepairAction <scriptblock>`. Owns the read/write envelope; calls `.Save()` only when `RepairAction` returns non-null. Both repair functions become thin wrappers.

**Acceptance criteria:**
- [ ] `Get-ShortcutObject` and `.Save()` in exactly one place for repair
- [ ] `$null` from `RepairAction` suppresses `.Save()`
- [ ] Existing repair behaviour unchanged

---

### LEAN-07 — `Start-AppxApp` failure recovery on timeout (design decision)

**Scope:** `Start-AppxApp`

**Rationale:**  
`Start-Win32App` presents a failure menu on timeout; `Start-AppxApp` silently returns `$false`. Asymmetry leaves Appx failures unrecoverable.

**Options:**
- **A** — Add `Invoke-FailureRecovery` to `Start-AppxApp` after timeout
- **B** — Keep Appx simpler; document asymmetry as by design

**Acceptance criteria (Option A):**
- [ ] Timeout → `Show-FailureMenu` with Skip and Delete options
- [ ] No retry loop added unless explicitly scoped

---

## Dependency Order for LEAN Tasks

```
LEAN-05   (self-contained)
LEAN-01   (new wrapper, two call sites swap)
LEAN-02   (Add-Shortcut simplification)
LEAN-06   (repair envelope, no external callers change)
LEAN-04   (Start-Win32App restructure)
LEAN-03   (touches menu + failure recovery)
LEAN-07   (design decision before implementation)
```

---

## Removed (Not Achievable Portably)

| Task | Reason |
|---|---|
| **TEST-05** | Mocking `Get-Process`/`Start-Sleep` not reliable without external mock library in PS 5.1. |
| **TEST-06** | Depends on `WindowsApps` ACL structure; not reproducible portably. |
| **TEST-07** | Mocking live `MainWindowHandle` requires a real GUI process. |
| **INT-03** | `Repair-ShortcutTarget` auto-discovery depends on a real broken install path. |
| **DUP-02** | Optional — `WorkingDirectory` derivation pattern; deferred by design. |

---

## Done

- [x] **AUD-01** — `Get-RelativeDepth` removed from main script; never called in startup sequence; comment updated in `Find-ExeWithinDepth`
- [x] **AUD-02** — `Repair-ShortcutTarget` allowlist/signature failure blocks now explicitly `return $null`; manual prompt no longer reached after a blocked auto-repair
- [x] **AUD-03** — `Sync-AppsFromStartMenu` duplicate `$launchType = 'Win32'` assignments removed; default set once before the `if` block
- [x] **AUD-04** — `Sync-AppsFromStartMenu` `elseif`/`else` Win32 branches merged; `$processName`/`$expectedExe`/`$expectedArguments` derived once as defaults, overridden only for Appx
- [x] **AUD-05** — `Start-Win32App` retry loop comment corrected to "up to 3 attempts"
- [x] **AUD-06** — `New-AppEntry` private helper extracted; `Add-Shortcut` and `Sync-AppsFromStartMenu` both call it
- [x] **AUD-07** — `Export-AppsConfig` success message condensed to one line
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
