# Win11 Startup Launcher — Task Backlog

---

## Open

| ID | Category | Description |
|---|---|---|
| **README-01** | Docs | README outdated: references `apps.json` and `$MaxRepairDepth`; missing menu `[6] Sync` workflow; missing all user menu scenario / workflow descriptions |
| **LEAN-01** | Refactor | Collapse `Test-ExePathAllowed` + `Test-ExeSignatureTrusted` into single `Test-ExeAcceptable` wrapper |
| **LEAN-02** | Refactor | Remove inline `New-AppShortcut` from `Add-Shortcut`; delegate to `Initialize-Shortcut` so shortcut creation has one owner |
| **LEAN-03** | Refactor | Extract `Invoke-MenuAction` to unify `Show-AppPicker` + dispatch used by main menu and `Invoke-FailureRecovery` choice `[2]` |

---

## Closed

| ID | Category | Description | Resolution |
|---|---|---|---|
| **LEAN-04** | Refactor | Extract `Invoke-LaunchAttempt` from `Start-Win32App`; collapse timeout + exception `Invoke-FailureRecovery` calls into one | Done — `Invoke-LaunchAttempt` returns `'Success'`/`'Retry'`/`'Abort'`; `Start-Win32App` loop is a thin `switch`. |
| **LEAN-05** | Refactor | Move zero-shortcut `Write-Warning` inside `Sync-AppsFromStartMenu`; callers check bool only | Done |
| **LEAN-06** | Refactor | Wrap `Get-ShortcutObject` + `.Save()` envelope in `Invoke-ShortcutRepair`; `Repair-ShortcutTarget` and `Repair-ShortcutArguments` become thin scriptblock delegates | Done |
| **LEAN-07** | Design | Decide and document whether `Start-AppxApp` gets failure recovery on timeout | Option B — intentional asymmetry documented in header comment and inline above `Start-AppxApp`. |

---

## LEAN Task Details

### LEAN-01 — Collapse exe
