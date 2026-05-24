# Win11startup.ps1 — Task Tracker

All items reference `win11-startup/Win11startup.ps1`.

---

## Completed

**Bugs:** BUG-01–BUG-08, BUG-A–BUG-G

**Hardening:** LEAN-01–LEAN-07, ROB-01, ROB-02, ROB-04, HARD-04, HARD-05

**Features:** SYNC-01, DUP-01, AUD-01, QOL-01–QOL-05, T-07, INT-01, INT-02, UX-02–UX-04, FIX-05–FIX-07

**Tests added (2026-05-23):** NEW-TEST-14 through NEW-TEST-27

**Test gaps closed:** TEST-GAP-01–TEST-GAP-07

**Testability fixes (2026-05-24):** FIX-TEST-01–FIX-TEST-10

**Live run validated (2026-05-24):**
- Phone Link — launched via .lnk; ProcessName back-filled as `PhoneExperienceHost`; config saved ✅
- Microsoft Edge — already open, skipped ✅
- Google Chrome — already open, skipped ✅
- Startup sequence completed successfully (no failures) ✅

---

## Principles

- Reuse `Test-ExeAcceptable`, `New-AppEntry`, `Initialize-Shortcut`, `Invoke-ShortcutRepair`
- All `.lnk` writes go through `New-AppShortcut` (LEAN-02)
- All exe validation goes through `Test-ExeAcceptable` (LEAN-01)
- No dead branches, unused variables, or uncalled helpers
