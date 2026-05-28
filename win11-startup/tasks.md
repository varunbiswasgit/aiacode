# Win11 Startup Manager — Task Log

## Pending

*(no pending tasks)*

## Completed

- [x] **Option 4 → Modify: no user option to fix WARNING after repair failure**
  When `Repair-ShortcutArguments` returned `$null` inside `Edit-Shortcut` (e.g. Phone Link — AUMID no longer matches any WindowsApps folder), the script emitted two WARNING lines and silently returned with no user action offered.
  When `Repair-ShortcutArguments` returned `$null` inside `Invoke-LaunchAttempt`, the app was silently aborted.
  Fix: extracted shared helper `Invoke-ManualAumidPrompt`; both `Edit-Shortcut` and `Invoke-LaunchAttempt` now call it on repair failure. `Invoke-ManualAumidPrompt` also clears `ProcessName` after a successful AUMID correction so stale process names cannot cause a false "already open" skip. `Invoke-LaunchAttempt` returns `'Retry'` when the user supplies a valid AUMID.
  Committed: 2026-05-28

- [x] **Invoke-LaunchAttempt: manual AUMID prompt on repair failure**
  When `Repair-ShortcutArguments` returns `$null` during auto-launch, the script silently aborted the app instead of offering the user a manual fix path.
  Fix: replaced the silent `return 'Abort'` block with a call to `Invoke-ManualAumidPrompt`; if the user provides a valid AUMID, returns `'Retry'` so the launch sequence re-attempts.
  Committed: 2026-05-28

- [x] **Edit-Shortcut: manual AUMID prompt on repair failure**
  Option 4 → Modify → Appx app (e.g. Phone Link) showed WARNING messages but gave no user option to correct the broken AUMID.
  Fix: extracted shared helper `Invoke-ManualAumidPrompt`; `Edit-Shortcut` now calls it when `Repair-ShortcutArguments` returns `$null`.
  Committed: 2026-05-27

- [x] **Sync-AppsFromStartMenu: LaunchType always stayed Win32 for Appx shortcuts**
  `Test-IsAppxShortcut` block ran but never set `$launchType = 'Appx'`, so Phone Link and Sticky Notes were written as Win32 in JSON after every sync.
  Fix: added `$launchType = 'Appx'` inside the detection block.

- [x] **Import-AppsConfig: auto-correct Win32→Appx on load**
  Added `Test-IsAppxShortcut` guard in the config loader so entries miscategorised as Win32 are silently corrected at load time without requiring a re-sync.

- [x] **Initialize-Shortcut: auto-correct LaunchType before shortcut work**
  Added `Test-IsAppxShortcut` guard so miscategorised Appx apps are fixed and saved the first time `Initialize-Shortcut` runs for them.
