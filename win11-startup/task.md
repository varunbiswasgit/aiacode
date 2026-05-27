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

**BUG-H (2026-05-24):** `Sync-AppsFromStartMenu` was overwriting back-filled `ProcessName` with blank on every sync for `shell:appsFolder` shortcuts. Fixed by loading existing config before the loop and preserving any non-empty `ProcessName`. Warning suppressed when value already known. Validated: `[6]` no longer shows warning after `[1]` has run. ✅

**BUG-I (2026-05-24):** `Show-AppList` crashed on PS5.1 with "index out of range" format error on the separator row (`'-'*N` expressions inside `-f` argument list). Fixed by pre-computing separator strings into variables. Validated: `[5]` displays clean table. ✅

---

## Completed (2026-05-27)

### RUN-SEQ-01 — Startup folder gate (`Test-StartupFolder`)
- Added `Test-StartupFolder` as Gate 1 in the run sequence.
- If `$script:startMenu` folder does not exist, prompts the user to enter a correct path.
- Updates `$script:startMenu` in-memory and proceeds; exits with error log entry if user provides no valid path. ✅

### RUN-SEQ-02 — No-shortcuts gate (`Test-StartupFolderHasShortcuts`)
- Added `Test-StartupFolderHasShortcuts` as Gate 2 in the run sequence.
- If startup folder has no numbered `.lnk` files, routes user to Add-Shortcut module.
- Re-checks after add; continues or exits cleanly. ✅

### RUN-SEQ-03 — Per-app recovery chain (`Invoke-AppWithRecovery`)
- Replaced ad-hoc launch/repair calls with a single `Invoke-AppWithRecovery` function.
- Decision chain per app:
  1. `.lnk` file exists → if not, call `Initialize-Shortcut`; skip with warning if still missing.
  2. Shortcut target valid (Win32) → if not, call `Repair-ShortcutTarget`; on failure offer manual detection or delete.
  3. Shortcut arguments valid (Appx) → if not, call `Repair-ShortcutArguments` + `Invoke-ManualAumidPrompt`; fall through to `Invoke-ManualLaunchDetection`.
  4. Pre-launch presence check (see RUN-SEQ-04).
  5. Launch; on timeout offer manual detection, skip, or delete. ✅

### RUN-SEQ-04 — Tray-aware presence check (`PresenceMode`)
- Extended `Test-AppAlreadyOpen` to honour a `PresenceMode` field (persisted in JSON):
  - `Window` (default) — visible window + matching title required. Prevents background host processes (e.g. `PhoneExperienceHost`) from falsely triggering skip.
  - `Tray` — process existence alone is sufficient (Greenshot, ShareFile).
  - `WindowOrTray` — either condition met (OneDrive: tray always present, Explorer window optional).
- `Wait-ForAppReady` updated to skip window-wait for `Tray` / `WindowOrTray` modes.
- All `Start-Win32App`, `Start-AppxApp`, and `Invoke-AppWithRecovery` call sites pass `$App.PresenceMode`. ✅
- JSON updated: OneDrive → `WindowOrTray`; Greenshot → `Tray`; ShareFile → `Tray`; all others → `Window`. ✅

### RUN-SEQ-05 — Manual launch detection (`Invoke-ManualLaunchDetection`)
- Added `Invoke-ManualLaunchDetection` as the final fallback when automatic repair fails.
- Watches for new windowed processes AND new tray processes (no window) over a 60-second window.
- Auto-matches on window title or process name against `$App.Name` / `$App.StartAppName`.
- If multiple candidates, prompts user to pick the correct process.
- On match: back-fills `ProcessName`; for Appx apps captures AUMID via `Get-StartApps` and saves to shortcut + config. ✅

---

## Principles

- Reuse `Test-ExeAcceptable`, `New-AppEntry`, `Initialize-Shortcut`, `Invoke-ShortcutRepair`
- All `.lnk` writes go through `New-AppShortcut` (LEAN-02)
- All exe validation goes through `Test-ExeAcceptable` (LEAN-01)
- No dead branches, unused variables, or uncalled helpers
