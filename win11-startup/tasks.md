# Win11 Startup Manager — Task Log

All items reference `win11-startup/Win11startup.ps1`.

---

## Pending

*(no pending tasks)*

---

## Completed

### Bugs fixed

- **BUG-01–BUG-08, BUG-A–BUG-I** — various crash, silent-abort, and mis-classification bugs corrected across config load, sync, shortcut creation, and launch path.

### Hardening

- **LEAN-01–LEAN-07, ROB-01, ROB-02, ROB-04, HARD-04, HARD-05** — dead-code removal, single-responsibility rewrites, exe validation, and config guard-rails.

### Features

- **SYNC-01, DUP-01, AUD-01, QOL-01–QOL-05, T-07, INT-01, INT-02, UX-02–UX-04, FIX-05–FIX-07** — sync from Start Menu, duplicate detection, audit log, UX polish, and integration helpers.

### Tests

- **NEW-TEST-14–NEW-TEST-27** added (2026-05-23).
- **TEST-GAP-01–TEST-GAP-07** closed.
- **FIX-TEST-01–FIX-TEST-10** testability fixes (2026-05-24).

---

### BUG-H (2026-05-24) — Sync overwrote back-filled `ProcessName`
`Sync-AppsFromStartMenu` was blanking `ProcessName` on every sync for `shell:appsFolder` shortcuts. Fixed by loading existing config before the loop and preserving any non-empty `ProcessName`. ✅

### BUG-I (2026-05-24) — `Show-AppList` format crash on PS 5.1
Separator-row format string used `'-'*N` inside `-f` argument list, causing index-out-of-range on PS 5.1. Fixed by pre-computing separator strings. ✅

### Live run validated (2026-05-24)
- Phone Link — launched via `.lnk`; `ProcessName` back-filled; config saved ✅
- Microsoft Edge — already open, skipped ✅
- Google Chrome — already open, skipped ✅
- Startup sequence completed successfully ✅

---

### RUN-SEQ-01 — Startup folder gate (2026-05-27)
Added `Test-StartupFolder` as Gate 1. If `$script:startMenu` does not exist, prompts the user for a correct path; exits with error log entry if none supplied. ✅

### RUN-SEQ-02 — No-shortcuts gate (2026-05-27)
Added `Test-StartupFolderHasShortcuts` as Gate 2. Routes user to Add-Shortcut when no numbered `.lnk` files exist; re-checks after add. ✅

### RUN-SEQ-03 — Per-app recovery chain (2026-05-27)
Replaced ad-hoc launch/repair calls with `Invoke-AppWithRecovery`. Decision chain per app: `.lnk` existence → target validity (Win32) → argument validity (Appx) → pre-launch presence check → launch with timeout recovery. ✅

### RUN-SEQ-04 — Tray-aware presence check (2026-05-27)
Extended `Test-AppAlreadyOpen` with a `PresenceMode` field persisted in JSON:
- `Window` (default) — visible window + matching title required.
- `Tray` — process existence alone is sufficient (e.g. screenshot tools, file-sync clients).
- `WindowOrTray` — either condition met.
`Wait-ForAppReady` updated accordingly. ✅

### RUN-SEQ-05 — Manual launch detection (2026-05-27)
Added `Invoke-ManualLaunchDetection` as final fallback when automatic repair fails. Watches for new windowed/tray processes over 60 s; auto-matches on title or process name; prompts user to pick if multiple candidates; back-fills `ProcessName` and AUMID on match. ✅

---

### Option 4 → Modify: manual AUMID prompt on repair failure (2026-05-28)
When `Repair-ShortcutArguments` returned `$null` inside `Edit-Shortcut` or `Invoke-LaunchAttempt`, the script emitted WARNING lines and silently aborted with no user action offered.

Fix:
- Extracted shared helper `Invoke-ManualAumidPrompt` — shows `[1] Enter manually / [2] Skip` prompt, writes new AUMID to `.lnk`, clears stale `ProcessName`, saves config.
- `Edit-Shortcut` calls `Invoke-ManualAumidPrompt` when `Repair-ShortcutArguments` returns `$null`.
- `Invoke-LaunchAttempt` calls `Invoke-ManualAumidPrompt` on repair failure; returns `'Retry'` so `Start-AppSequence` re-attempts the launch (up to 2 retries). ✅

---

## Principles

- Reuse `Test-ExeAcceptable`, `New-AppEntry`, `Initialize-Shortcut`, `Invoke-ShortcutRepair`.
- All `.lnk` writes go through `New-AppShortcut` (LEAN-02).
- All exe validation goes through `Test-ExeAcceptable` (LEAN-01).
- No dead branches, unused variables, or uncalled helpers.
