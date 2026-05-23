# Win11startup.ps1 — Task Tracker

All items reference `win11-startup/Win11startup.ps1`.

---

## Open — Script/Test Signature Mismatches

All items are **additive only** — no existing logic changes required.

| ID | Function | Required Change | Test |
|----|----------|-----------------|------|
| FIX-TEST-01 | `Get-ShortcutObject` | Add `[Alias('LnkPath')]` to `-ShortcutPath` param | NEW-TEST-18 |
| FIX-TEST-02 | `Find-ExeWithinDepth` | Add aliases `-SearchRoot`, `-ExeName`; add optional `-MaxDepth` param with depth-bounded filtering | NEW-TEST-17 |
| FIX-TEST-03 | `New-AppShortcut` | Add `-App [PSCustomObject]` param mapping `$App.ShortcutPath` → `-Path` | NEW-TEST-19 |
| FIX-TEST-04 | `Invoke-AppLaunch` | Add thin wrapper over `Invoke-LaunchAttempt` | NEW-TEST-20 |
| FIX-TEST-05 | `Test-ShortcutHealthy` | Extract standalone function: `.lnk` exists → target exists → `Test-ExeAcceptable` passes | NEW-TEST-21 |
| FIX-TEST-06 | `Get-AppPresence` | Add wrapper over `Get-AppPresenceMode`; map `'Tray'` → `'Running'`, `'Window'` → `'WindowVisible'` | NEW-TEST-22 |
| FIX-TEST-07 | `Wait-ForWindowByTitle` | Add `-TitleFragment`/`-TimeoutSeconds` overload returning `[bool]` | NEW-TEST-24 |
| FIX-TEST-08 | `Sync-AppsFromStartMenu` | Add optional `-StartMenuPath` param (currently hardcodes `$script:startMenu`) | NEW-TEST-25 |
| FIX-TEST-09 | `Resolve-ConfigPath` | Add optional `-Path` param | NEW-TEST-26 |
| FIX-TEST-10 | `Start-Win32App` | Add `-MaxAttempts [int]` param defaulting to `3` | NEW-TEST-23 |

---

## Completed

**Bugs:** BUG-01–BUG-08, BUG-A–BUG-G

**Hardening:** LEAN-01–LEAN-07, ROB-01, ROB-02, ROB-04, HARD-04, HARD-05

**Features:** SYNC-01, DUP-01, AUD-01, QOL-01–QOL-05, T-07, INT-01, INT-02, UX-02–UX-04, FIX-05–FIX-07

**Tests added (2026-05-23):** NEW-TEST-14 through NEW-TEST-27

**Test gaps closed:** TEST-GAP-01–TEST-GAP-07

---

## Principles

- Reuse `Test-ExeAcceptable`, `New-AppEntry`, `Initialize-Shortcut`, `Invoke-ShortcutRepair`
- All `.lnk` writes go through `New-AppShortcut` (LEAN-02)
- All exe validation goes through `Test-ExeAcceptable` (LEAN-01)
- No dead branches, unused variables, or uncalled helpers
