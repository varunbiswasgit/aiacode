# Win11 Startup Launcher — Testing Guide

All tests are manual. No automated Pester test file exists — core behaviour involves process detection, shortcut COM objects (`WshShell`), and AUMID resolution that require a live Windows environment.

---

## Test Environment

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- `Get-AppxPackage` and `Get-StartApps` available (standard on Windows 10/11)
- All Win32 shortcut files present in `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\` with numbered naming convention
- Entry 04 shortcut (`04 Sticky Notes.lnk`) Target field set to `"...\ONENOTE.EXE" /memoryWindow start`
- Entry 06 shortcut (`06 Phone Link.lnk`) Target = `C:\Windows\explorer.exe`; Arguments = `shell:appsFolder\Microsoft.YourPhone_8wekyb3d8bbwe!App`
- Phone Link installed (for TC-11 through TC-15)

---

## Test Cases

### TC-01 — Normal Win32 launch: shortcut target valid, process not running

**Setup:** Shortcut target exists. App not running.
**Action:** Run the script.
**Expected:** `[Name]: launching via shortcut: <path>.lnk` then `(presence mode: Window)` or `(presence mode: Tray)` then `[Name]: ready.`
**Pass criteria:** App process visible in Task Manager within 30 seconds.

### TC-02 — Sticky Notes: baked-in arguments honoured via WshShell.Run

**Setup:** Entry 04 shortcut Target contains `"...\ONENOTE.EXE" /memoryWindow start`. `ONENOTE` not running.
**Action:** Run the script.
**Expected:** `Sticky Notes: launching via shortcut: ...\04 Sticky Notes.lnk` then `(presence mode: Window)` then `Sticky Notes: ready.` A sticky note window (not full OneNote UI) opens.
**Pass criteria:** Sticky note window visible. `ONENOTE` process running. Full OneNote UI NOT open.

### TC-03 — Already running: process detected before launch

**Setup:** Launch one app manually before running the script.
**Action:** Run the script.
**Expected:** `[Name]: already open. Skipping.` No second instance launched.
**Pass criteria:** Only one instance of the process in Task Manager.

### TC-04 — Sticky Notes / OneNote process order: OneNote skipped after Sticky Notes

**Setup:** Neither `ONENOTE` process running. Both entries 04 and 05 present.
**Action:** Run the script.
**Expected:** Entry 04 launches and `ONENOTE` starts. Entry 05 logs `OneNote: already open. Skipping.`
**Pass criteria:** One `ONENOTE` process. Sticky note window visible. Full OneNote UI not opened separately.

### TC-05 — Win32: shortcut target broken, exe found after climbing 3 levels up

**Setup:** Change a shortcut's TargetPath to a non-existent path several folders deep (e.g. `C:\Program Files\Foo\Bar\Baz\missing.exe`). The actual exe exists somewhere under the ancestor 3 levels up from `Baz` (i.e. under `C:\Program Files\Foo`).
**Action:** Run the script.
**Expected:** `shortcut target missing or invalid` warning, then `searching for <exe> under <ancestor> (3 levels up, all subfolders)`, then `found replacement at <path>. Updating shortcut.` then `shortcut repaired. Proceeding with launch.` App launches via the updated `.lnk`.
**Pass criteria:** Shortcut target updated to the real exe path. App running.

### TC-06 — Win32: shortcut target broken, exe not found after climbing 3 levels up, user provides valid path

**Setup:** Shortcut target broken and the exe does not exist anywhere under the ancestor 3 levels up from the broken target's folder.
**Action:** When prompted, enter the full correct path to the exe.
**Expected:** `<exe> not found under <ancestor>` warning, then user prompt appears. On valid input: shortcut updated. `shortcut repaired. Proceeding with launch.` App launches.
**Pass criteria:** Shortcut updated. App running.

### TC-07 — Win32 prompt: invalid paths entered before correct one

**Setup:** Same as TC-06.
**Action:** Enter a non-existent path, then a path with the wrong filename, then the correct path.
**Expected:** Warning on each bad input. Prompt repeats until correct input.
**Pass criteria:** Prompt repeats on bad input. Succeeds on correct input.

### TC-08 — Win32 prompt: user presses Enter to skip

**Setup:** Same as TC-06.
**Action:** Press Enter at the prompt.
**Expected:** `[Name]: shortcut could not be repaired. Skipping.` App in failure list.
**Pass criteria:** Script continues. Failed app listed at end.

### TC-09 — Win32: shortcut .lnk file itself does not exist

**Setup:** Delete or rename one shortcut file.
**Action:** Run the script.
**Expected:** `WARNING: [Name]: shortcut file not found: <path>` and inline failure menu (Add / Modify / Skip).
**Pass criteria:** Script continues after user selects Skip. Remaining apps launch.

### TC-10 — Win32: process does not appear within 30 seconds

**Setup:** Point a shortcut to a valid exe that does not produce a detectable process within 30 seconds.
**Action:** Run the script.
**Expected:** After 30 seconds: `WARNING: [Name]: did not become ready within 30 seconds.` and inline failure menu.
**Pass criteria:** Script does not hang beyond 30 seconds. Failure logged.

### TC-11 — Phone Link: valid shortcut Arguments, no repair needed

**Setup:** Entry 06 shortcut Arguments field matches the installed AUMID. `PhoneExperienceHost` not running.
**Action:** Run the script.
**Expected:** `Phone Link: launching via shortcut: ...\06 Phone Link.lnk` then `(presence mode: Window)` or `(presence mode: Tray)` then `Phone Link: ready.`
**Pass criteria:** `PhoneExperienceHost` running in Task Manager.

### TC-12 — Phone Link: Arguments stale, self-healing reconstructs AUMID

**Setup:** Manually set the shortcut Arguments to an outdated AUMID (wrong version number). Phone Link installed.
**Action:** Run the script.
**Expected:** `WARNING: [Name]: shortcut Arguments missing or invalid`. Script scans WindowsApps, finds the matching package folder, reads `AppxManifest.xml`, reconstructs AUMID, updates the shortcut, and launches.
**Pass criteria:** Shortcut Arguments updated to correct AUMID. `PhoneExperienceHost` running.

### TC-13 — Phone Link: WindowsApps folder not found during argument repair

**Setup:** Temporarily rename or set the `WindowsApps` scan path to a non-existent location in the script.
**Action:** Run the script with stale Arguments.
**Expected:** `no folder matching '*<fragment>*' found in WindowsApps.` App in failure list.
**Pass criteria:** Script continues. Phone Link in final failure summary.

### TC-14 — Phone Link: already running (tray/window)

**Setup:** Phone Link already open before running the script.
**Action:** Run the script.
**Expected:** `Phone Link: already open. Skipping.` No second instance launched.
**Pass criteria:** Single `PhoneExperienceHost` process in Task Manager.

### TC-15 — Presence mode: Window app detected correctly

**Setup:** Any app that opens a visible window (e.g. Outlook, Chrome). Not running.
**Action:** Run the script.
**Expected:** Within `$SettleSeconds` (5 s), `(presence mode: Window)` logged. App confirmed ready when `MainWindowHandle != 0`.
**Pass criteria:** Mode logged as `Window`. App window visible.

### TC-16 — Presence mode: Tray app detected correctly

**Setup:** Any tray-only app (e.g. OneDrive). Not running.
**Action:** Run the script.
**Expected:** No window appears within 5 s. `(presence mode: Tray)` logged. App confirmed ready on process presence alone.
**Pass criteria:** Mode logged as `Tray`. Process visible in Task Manager. No window required.

### TC-17 — Presence mode: $SettleSeconds tuning — slow machine

**Setup:** Increase `$SettleSeconds` to 10 in the script. Use an app that normally opens a window after 6–8 s.
**Action:** Run the script.
**Expected:** Window appears before the settle period expires. App classified as `Window`, not `Tray`.
**Pass criteria:** Mode logged as `Window`. Misclassification avoided by longer settle time.

### TC-18 — Full sequence: all apps launch successfully

**Setup:** All Win32 shortcut targets valid. Phone Link installed. No apps pre-running.
**Action:** Run the script.
**Expected:** Each app launches sequentially with presence mode logged. Console ends with `Startup sequence completed successfully.`
**Pass criteria:** All processes visible in Task Manager. No warnings or failure list.

### TC-19 — Win32: repaired shortcut persists on second run

**Setup:** Trigger TC-05 or TC-06 to repair a shortcut. Close the launched app.
**Action:** Run the script a second time without changes.
**Expected:** Repaired shortcut used directly. No repair logic triggered.
**Pass criteria:** No `shortcut target missing` warning for the previously repaired app.

---

## Pass Criteria Summary

| TC | Scenario | Pass If |
|----|----------|---------|
| TC-01 | Valid Win32, not running | App launches; presence mode logged; ready within 30 s |
| TC-02 | Sticky Notes via WshShell.Run | Sticky note window opens; `/memoryWindow start` honoured from .lnk |
| TC-03 | App already running | Skipped; no second instance |
| TC-04 | Sticky Notes then OneNote process order | OneNote skipped; one ONENOTE process |
| TC-05 | Broken Win32 target, exe found after 3 levels up + full subfolder search | Shortcut updated; app launches via repaired .lnk |
| TC-06 | Broken Win32 target, exe not found after 3-level climb, user provides path | Shortcut updated; app launches via repaired .lnk |
| TC-07 | Invalid then valid paths at prompt | Prompt repeats; accepts correct input |
| TC-08 | User skips prompt | App in failure list; script continues |
| TC-09 | `.lnk` file missing | Inline failure menu shown; script continues after Skip |
| TC-10 | Win32 process timeout | Inline failure menu after 30 s; script continues |
| TC-11 | Phone Link: valid Arguments | Launches via explorer.exe shell:appsFolder; PhoneExperienceHost running |
| TC-12 | Phone Link: stale Arguments, self-healed | AUMID reconstructed from WindowsApps; shortcut updated; app launches |
| TC-13 | Phone Link: WindowsApps scan fails | Warning logged; app in failure list |
| TC-14 | Phone Link already running | Skipped; no second instance |
| TC-15 | Presence mode: Window | Mode logged as Window; window visible |
| TC-16 | Presence mode: Tray | Mode logged as Tray; process running; no window needed |
| TC-17 | Presence mode: slow machine tuning | Window classified correctly with larger $SettleSeconds |
| TC-18 | Full sequence | All apps running; success message |
| TC-19 | Repaired shortcut reused | No repair on second run |
