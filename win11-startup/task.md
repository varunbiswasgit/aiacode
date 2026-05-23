# Win11startup.ps1 — Lean-Down & Reusability Backlog

Derived from full flow analysis (May 2026).  
Each task has a code prefix (LEAN-xx), scope, rationale, and acceptance criteria.

---

## LEAN-01 — Collapse exe validation into `Test-ExeAcceptable`

**Status:** Open  
**Scope:** `Test-ExePathAllowed`, `Test-ExeSignatureTrusted`, `Repair-ShortcutTarget`, `Prompt-ForExactExePath`

**Rationale:**  
`Test-ExePathAllowed` and `Test-ExeSignatureTrusted` are always called together in the same order at two call sites. There is no case where one is called without the other. The two-function sequence is a hidden contract that every future caller must know about.

**Change:**  
Create `Test-ExeAcceptable -ExePath -ExpectedPublisher`. It calls `Test-ExePathAllowed` then `Test-ExeSignatureTrusted` internally and returns `$true` only when both pass. Replace the two-call sequence at both sites with a single `Test-ExeAcceptable` call.

**Acceptance criteria:**
- [ ] `Test-ExePathAllowed` and `Test-ExeSignatureTrusted` remain as private helpers (no callers outside `Test-ExeAcceptable`)
- [ ] `Repair-ShortcutTarget` calls `Test-ExeAcceptable` once
- [ ] `Prompt-ForExactExePath` calls `Test-ExeAcceptable` once
- [ ] Existing Pester tests pass unchanged or updated to target `Test-ExeAcceptable`

---

## LEAN-02 — Unify shortcut creation in `Initialize-Shortcut`

**Status:** Open  
**Scope:** `Add-Shortcut` (new entry branch), `Initialize-Shortcut`

**Rationale:**  
The new-entry branch at the bottom of `Add-Shortcut` duplicates the `New-AppShortcut` call logic that already exists in `Initialize-Shortcut`. Both branches write a `.lnk` for an exe-path entry or an arguments-based entry using nearly identical code. When `Initialize-Shortcut` changes, `Add-Shortcut` must also change — a maintenance trap.

**Change:**  
After `Add-Shortcut` builds and appends the new `$newEntry` to `$script:apps`, call `Initialize-Shortcut -App $newEntry` to write the `.lnk`. Remove the inline `New-AppShortcut` block from `Add-Shortcut`. `Initialize-Shortcut` already guards against overwriting an existing file, so the behaviour is identical.

**Acceptance criteria:**
- [ ] `New-AppShortcut` is no longer called directly from `Add-Shortcut`
- [ ] `Initialize-Shortcut` is the single `.lnk` creation path for both new entries and re-init
- [ ] Win32 exe-path entry created by `Add-Shortcut` matches the shortcut written by `Initialize-Shortcut`
- [ ] Win32-with-args (explorer.exe + shell:appsFolder) entry written correctly

---

## LEAN-03 — Extract `Invoke-MenuAction` to unify picker + dispatch

**Status:** Open  
**Scope:** Main menu `[2]`/`[3]`/`[4]` dispatch, `Invoke-FailureRecovery` choice `[2]`

**Rationale:**  
The pattern `Show-AppPicker -Prompt <string>` followed by an action call (`Edit-Shortcut`, `Remove-Shortcut`, `Add-Shortcut`) appears four times in the main menu switch and once inside `Invoke-FailureRecovery`. The prompt string is the only variable. Any change to picker or action behaviour must be replicated across all five sites.

**Change:**  
Create `Invoke-MenuAction -Action <string> -Prompt <string>` where `Action` is one of `'Add'`, `'Delete'`, `'Modify'`. It calls `Show-AppPicker` with the supplied prompt, then dispatches to the correct function. Replace all five call sites. The main menu switch and `Invoke-FailureRecovery` choice `[2]` both call `Invoke-MenuAction`.

**Acceptance criteria:**
- [ ] `Show-AppPicker` is not called directly from the main switch block or `Invoke-FailureRecovery`
- [ ] All three action types (`Add`, `Delete`, `Modify`) reachable through `Invoke-MenuAction`
- [ ] Cancel (`$null` return from picker) handled centrally inside `Invoke-MenuAction`
- [ ] Failure recovery choice `[2]` behaviour unchanged

---

## LEAN-04 — Consolidate `Invoke-FailureRecovery` call sites in `Start-Win32App`

**Status:** Open  
**Scope:** `Start-Win32App` (timeout path, exception path)

**Rationale:**  
`Invoke-FailureRecovery` is called three times inside `Start-Win32App`. The missing-shortcut path correctly passes `Initialize-Shortcut` as `PreRetryAction`. The timeout path and exception path both pass `Edit-Shortcut` as `PreRetryAction` — they differ only in `Context` string. This duplication means the retry decision logic is scattered across three nearly-identical call sites.

**Change:**  
Extract the launch-attempt body (from `Get-ShortcutObject` through the exception `catch`) into a private helper `Invoke-LaunchAttempt -App`. It returns a tri-state: `'Success'`, `'Retry'`, or `'Abort'`. `Start-Win32App` loops on `'Retry'` and returns on `'Success'` or `'Abort'`. The two `Invoke-FailureRecovery` calls with `Edit-Shortcut` collapse into one inside `Invoke-LaunchAttempt`.

**Acceptance criteria:**
- [ ] `Invoke-FailureRecovery` called at most twice in the launch flow (once for missing shortcut, once inside `Invoke-LaunchAttempt`)
- [ ] Retry loop bound (max 3 attempts) preserved
- [ ] Exception and timeout paths both reach the same `Invoke-FailureRecovery` call
- [ ] `Start-Win32App` function body shortened by at least 30 lines

---

## LEAN-05 — Move post-sync guard into `Sync-AppsFromStartMenu`

**Status:** Open  
**Scope:** `Sync-AppsFromStartMenu`, `Resolve-ConfigPath`, main menu `[6]`

**Rationale:**  
Both callers of `Sync-AppsFromStartMenu` check its `$false` return and print their own error message. The guard logic and messaging are inconsistent between the two. Adding a third caller would require a third copy of the guard.

**Change:**  
`Sync-AppsFromStartMenu` keeps its `return $false` for the no-files case. Add a single descriptive `Write-Warning` inside the function for that case (already partially present). Callers only need to check the bool and `exit` or `return` — no inline messaging. Remove the duplicated warning text from `Resolve-ConfigPath`.

**Acceptance criteria:**
- [ ] Warning message for zero-shortcut sync appears exactly once, from inside `Sync-AppsFromStartMenu`
- [ ] `Resolve-ConfigPath` and main menu `[6]` both act on the return bool without duplicating the message
- [ ] Behaviour on empty sync is identical in both callers (warning shown, no crash)

---

## LEAN-06 — Wrap shortcut read/write envelope in repair functions

**Status:** Open  
**Scope:** `Repair-ShortcutTarget`, `Repair-ShortcutArguments`

**Rationale:**  
Both repair functions open with `Get-ShortcutObject` and close with `$shortcut.Save()`. The read/write envelope is identical; only the middle logic differs. If the shortcut object model changes (e.g., switching from WshShell to a different COM approach), both functions need updating.

**Change:**  
Create `Invoke-ShortcutRepair -App -RepairAction <scriptblock>`. It calls `Get-ShortcutObject`, passes the shortcut object into `RepairAction`, then calls `.Save()` if `RepairAction` returns a non-null result. `Repair-ShortcutTarget` and `Repair-ShortcutArguments` become thin wrappers that supply the middle logic as a scriptblock.

**Acceptance criteria:**
- [ ] `Get-ShortcutObject` and `.Save()` called in exactly one place for repair operations
- [ ] `Repair-ShortcutTarget` and `Repair-ShortcutArguments` delegates contain only their unique logic
- [ ] `$null` return from `RepairAction` suppresses `.Save()` (no partial write)
- [ ] Existing repair behaviour unchanged end-to-end

---

## LEAN-07 — Add `Start-AppxApp` failure recovery on timeout

**Status:** Open (design decision required)  
**Scope:** `Start-AppxApp`

**Rationale:**  
`Start-Win32App` shows a failure menu when `Invoke-AppLaunchWait` returns `$false`. `Start-AppxApp` silently returns `$false` on timeout. This asymmetry means Appx apps that fail to start are logged as failures with no recovery opportunity. The user gets no inline prompt to retry or skip.

**Options:**
- A — Add `Invoke-FailureRecovery` to `Start-AppxApp` after timeout (mirrors Win32 path)
- B — Intentionally keep Appx simpler; document the asymmetry as by design

**Acceptance criteria (if Option A chosen):**
- [ ] `Invoke-AppLaunchWait` returns `$false` for Appx → `Show-FailureMenu` presented
- [ ] Skip and Delete options available (Add/Modify not applicable to Appx)
- [ ] `Start-AppxApp` does not gain a retry loop unless explicitly scoped

---

## Dependency Order

Safe implementation sequence to avoid merge conflicts:

```
LEAN-05   (self-contained, no callers change)
LEAN-01   (new function, two call sites swap)
LEAN-02   (Add-Shortcut simplification)
LEAN-06   (repair envelope — no external callers change)
LEAN-04   (Start-Win32App restructure)
LEAN-03   (picker unification — touches menu + recovery)
LEAN-07   (design decision first, then implement)
```

---

## Completed Tasks

| ID | Title | Commit |
|----|-------|--------|
| BUG-06 (rev) | Replace Resolve-ProcessName with Wait-ForWindowByTitle; back-fill ProcessName | `62a7129` |
| UX-04 | Add [4] Delete entry to Show-FailureMenu; reuse Remove-Shortcut | `62a7129` |
| FIX-07 | Write-ErrorLog before trap; boot-time config path prompt | prior commit |
| BUG-05 | Sync classifies explorer+shell: entries as Win32-with-args | prior commit |
