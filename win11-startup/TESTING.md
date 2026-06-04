# Win11 Startup Launcher — Testing Guide

All tests are manual and require a live Windows 10/11 environment. The script relies on `WScript.Shell` COM objects, filesystem `.lnk` inspection, and `Get-Process` — behaviours that cannot be fully exercised outside a real Windows session.

Automated Pester unit tests covering discovery logic, config I/O, and guard conditions are in `Win11startup.Tests.ps1`.

---

## Test Environment

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- `WScript.Shell` COM object available (standard inbox)
- Numbered `.lnk` shortcut files present in the configured Start Menu folder, using the naming convention `NN AppName.lnk` (1–2 digit prefix followed by a space)
- `Win11StartupConfig.json` absent or pre-populated (see individual test cases)

---

## Test Cases

### TC-01 — First run: config file absent, user provides paths

**Setup:** Delete or rename `Win11StartupConfig.json` so it does not exist.
**Action:** Run the script. When prompted, press Enter to accept the default config path, then enter the Start Menu folder path.
**Expected:** Config file created at the default path with `StartMenuPath` and an empty `Shortcuts` array. Startup sequence proceeds.
**Pass criteria:** `Win11StartupConfig.json` created. Shortcuts discovered and launched.

---

### TC-02 — First run: user supplies a custom config path

**Setup:** Delete `Win11StartupConfig.json`.
**Action:** Run the script. When prompted, enter a custom path (e.g., `C:\MyConfig\startup.json`). Provide the Start Menu folder path.
**Expected:** Config written to the supplied path. Startup sequence proceeds using that config.
**Pass criteria:** Custom config file exists. Shortcuts launched.

---

### TC-03 — Subsequent run: valid config loaded, no prompts

**Setup:** `Win11StartupConfig.json` present with a valid `StartMenuPath` and at least one entry in `Shortcuts`.
**Action:** Run the script.
**Expected:** No prompts shown. Script reads config, discovers `.lnk` files, and launches apps immediately.
**Pass criteria:** Startup sequence begins without any user input.

---

### TC-04 — Config file present but contains invalid JSON

**Setup:** Open `Win11StartupConfig.json` and corrupt it (e.g., delete a closing brace).
**Action:** Run the script.
**Expected:** Warning: `Config file exists but could not be loaded (invalid JSON). It will be reinitialized.` Script then prompts for Start Menu folder path and recreates the config.
**Pass criteria:** Warning emitted. New valid config created. Shortcuts launched.

---

### TC-05 — Normal Win32 launch: shortcut target valid, process not running

**Setup:** At least one numbered `.lnk` file present. The target app is not running.
**Action:** Run the script.
**Expected:** `Launching [AppName]...` logged. Process appears in Task Manager within 15 seconds.
**Pass criteria:** App process running in Task Manager. No warnings.

---

### TC-06 — Skip-if-running: process already active before launch

**Setup:** Start one of the configured apps manually before running the script.
**Action:** Run the script.
**Expected:** `Skipping [AppName] (already running).` No second instance launched.
**Pass criteria:** Single process instance in Task Manager for that app.

---

### TC-07 — Store / UWP app shortcut skipped

**Setup:** Add a `.lnk` file with a numeric prefix whose `TargetPath` is `C:\Windows\explorer.exe` and `Arguments` begins with `shell:appsFolder\`.
**Action:** Run the script.
**Expected:** `Skipping (unsupported store app): [AppName]` logged. No attempt to launch the app.
**Pass criteria:** Message logged. No UWP launch attempted. Script continues to next entry.

---

### TC-08 — Shortcut with empty or unreadable TargetPath skipped

**Setup:** Create a numbered `.lnk` file with no TargetPath set (save it without a target).
**Action:** Run the script.
**Expected:** `WARNING: Skipping '[AppName]' (invalid or missing target path).`
**Pass criteria:** Warning emitted. Script continues without throwing.

---

### TC-09 — Unreadable shortcut COM object skipped

**Setup:** Place a file with a `.lnk` extension and numeric prefix but corrupt its content so `WshShell.CreateShortcut` cannot read it.
**Action:** Run the script.
**Expected:** `WARNING: Skipping unreadable shortcut: [path]`
**Pass criteria:** Warning emitted. Script continues to the next entry.

---

### TC-10 — Process does not start within timeout: file-picker dialog appears

**Setup:** Point a numbered `.lnk` to a valid-looking but non-launching exe (e.g., a dummy `.exe` that exits immediately). Set `$ProcessStartTimeout = 5` for faster testing.
**Action:** Run the script. When the dialog appears, click Cancel.
**Expected:** After the timeout: `WARNING: '[AppName]' did not start within N seconds.` Windows Open File dialog opens. On Cancel: `No executable selected for '[AppName]'. Skipping.`
**Pass criteria:** Dialog appears. Script continues after Cancel without crashing.

---

### TC-11 — Shortcut repair via file-picker: user selects correct exe

**Setup:** Same as TC-10. This time, when the dialog appears, browse to and select the real executable for the app.
**Expected:** `Repairing [AppName] shortcut to target [path]`. `Re-launching [AppName]...` App starts within 15 seconds. Config updated with new `ProcessName`.
**Pass criteria:** App running in Task Manager. `.lnk` TargetPath updated. Config reflects new process name.

---

### TC-12 — Shortcut repair: app still does not start after repair

**Setup:** Same as TC-11 but select an exe that also does not produce a detectable process.
**Expected:** After the second timeout: `WARNING: '[AppName]' still did not start after repair. Moving on.`
**Pass criteria:** Warning emitted. Script moves to the next shortcut without hanging.

---

### TC-13 — Numeric sort order: shortcuts launch in ascending numeric order

**Setup:** Three numbered shortcuts: `03 Slack.lnk`, `01 Outlook.lnk`, `02 Teams.lnk`.
**Action:** Run the script and observe console output order.
**Expected:** `Launching Outlook...` first, then `Launching Teams...`, then `Launching Slack...`.
**Pass criteria:** Launch order matches numeric prefix, not filesystem or alphabetical order.

---

### TC-14 — Config auto-sync: new shortcut added to config on first encounter

**Setup:** Config exists but does not include an entry for one of the `.lnk` files in the Start Menu folder.
**Action:** Run the script.
**Expected:** The missing shortcut is launched and its entry (`Name`, `ShortcutPath`, `ProcessName`) is written to `Win11StartupConfig.json`.
**Pass criteria:** Config file contains the new entry after the run.

---

### TC-15 — Config auto-sync: stale ProcessName updated when shortcut is repaired

**Setup:** Config entry for an app has an outdated `ProcessName`. The shortcut's TargetPath now points to a different exe (simulating an app update).
**Action:** Run the script. Allow repair to complete.
**Expected:** `ProcessName` in config updated to reflect the new exe's base name. Config saved.
**Pass criteria:** `Win11StartupConfig.json` contains the corrected `ProcessName`.

---

### TC-16 — Start Menu folder not found: script exits cleanly

**Setup:** Set `StartMenuPath` in config to a folder that does not exist.
**Action:** Run the script.
**Expected:** `Start Menu folder not found: [path]. Exiting.` (red text). Script returns without launching anything.
**Pass criteria:** Error message displayed. No exception thrown. No apps launched.

---

### TC-17 — Empty Start Menu folder: script exits cleanly

**Setup:** `StartMenuPath` points to a real but empty folder (no `.lnk` files).
**Action:** Run the script.
**Expected:** `No numbered .lnk shortcuts found in '[path]'.` Script returns.
**Pass criteria:** Message displayed. No errors thrown.

---

### TC-18 — Full sequence: all apps launch successfully

**Setup:** All numbered `.lnk` files present with valid targets. No apps pre-running. Config up to date.
**Action:** Run the script.
**Expected:** Each app logs `Launching [AppName]...` and starts within `$ProcessStartTimeout` seconds. No warnings.
**Pass criteria:** All configured processes running in Task Manager. No warnings or skips.

---

### TC-19 — Repaired shortcut reused on second run

**Setup:** Trigger TC-11 to repair a shortcut. Close the launched app.
**Action:** Run the script a second time without changes.
**Expected:** Script uses the repaired `.lnk` directly. No timeout or repair dialog triggered.
**Pass criteria:** No `did not start within N seconds` warning for the previously repaired app.

---

## Pass Criteria Summary

| TC | Scenario | Pass If |
|----|----------|---------|
| TC-01 | First run, config absent | Config created; sequence proceeds |
| TC-02 | First run, custom config path | Config written to custom path; sequence proceeds |
| TC-03 | Subsequent run, valid config | No prompts; sequence starts immediately |
| TC-04 | Config present, invalid JSON | Warning emitted; config recreated; sequence proceeds |
| TC-05 | Normal Win32 launch | App running in Task Manager within 15 s |
| TC-06 | App already running | Skipped; single process instance |
| TC-07 | Store / UWP shortcut | Skipped with message; no launch attempted |
| TC-08 | Empty TargetPath in shortcut | Warning emitted; script continues |
| TC-09 | Unreadable COM shortcut | Warning emitted; script continues |
| TC-10 | Process timeout, user cancels dialog | Dialog shown; script continues after Cancel |
| TC-11 | Process timeout, user selects correct exe | Shortcut repaired; app launched; config updated |
| TC-12 | Process still fails after repair | Warning emitted; script moves to next app |
| TC-13 | Numeric launch order | Apps launch in ascending numeric prefix order |
| TC-14 | New shortcut added to config | New entry written to config |
| TC-15 | Stale ProcessName updated after repair | Config updated with correct ProcessName |
| TC-16 | Start Menu folder missing | Error message shown; clean exit |
| TC-17 | Start Menu folder empty | No shortcuts message; clean exit |
| TC-18 | Full sequence | All apps running; no warnings |
| TC-19 | Repaired shortcut reused on second run | No repair triggered on second run |
