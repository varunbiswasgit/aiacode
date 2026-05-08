# Win11startup.ps1 — Testing Readme

This document covers manual test cases for `Win11startup.ps1`. There is no automated Pester test file for this script because its core behavior involves process detection, shortcut COM objects, and AUMID resolution that requires a live Windows environment. All tests below are manual.

For environment setup (execution policy, PowerShell version), see [tests/README.md](README.md).

---

## Test Environment

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- `Get-AppxPackage` and `Get-StartApps` available (standard on Windows 10/11)
- All Win32 shortcut files present in `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\` with the numbered naming convention
- Phone Link installed (for TC-09 through TC-12)
- Entry 06 shortcut (`06 Sticky Notes.lnk`) target pointing to `ONENOTE.EXE`

---

## Test Cases

### TC-01 — Normal Win32 launch: shortcut target valid, no Arguments, process not running

**Setup:** Shortcut target path exists. App has no `Arguments` field. App is not running.  
**Action:** Run the script.  
**Expected:** Console shows `[Name]: launching <path>` (no arguments in log) and `[Name]: '[ProcessName]' is now running.`  
**Pass criteria:** App process visible in Task Manager within 30 seconds.

---

### TC-02 — Win32 launch with Arguments field (Sticky Notes)

**Setup:** Entry 06 (Sticky Notes) shortcut target points to a valid `ONENOTE.EXE`. `ONENOTE` process not running.  
**Action:** Run the script.  
**Expected:** Console shows `Sticky Notes: launching <path> /memoryWindow start` and `Sticky Notes: 'ONENOTE' is now running.` A sticky note window (not the full OneNote UI) opens.  
**Pass criteria:** Sticky note window visible on desktop. `ONENOTE` process running. OneNote full UI is NOT open.

---

### TC-03 — Already running: process detected before launch

**Setup:** Launch one app manually before running the script.  
**Action:** Run the script.  
**Expected:** Console shows `[Name]: '[ProcessName]' already running. Skipping.` No second instance launched.  
**Pass criteria:** Only one instance of the process in Task Manager.

---

### TC-04 — Sticky Notes / OneNote process order: OneNote skipped after Sticky Notes

**Setup:** Neither `ONENOTE` process running before the script. Both entries 06 and 07 present.  
**Action:** Run the script.  
**Expected:** Entry 06 (Sticky Notes) launches and `ONENOTE` starts. Entry 07 (OneNote) finds `ONENOTE` already running and logs `OneNote: 'ONENOTE' already running. Skipping.`  
**Pass criteria:** Only one `ONENOTE` process. Sticky note window visible. OneNote full UI not opened separately.

---

### TC-05 — Win32: shortcut target missing, exe found within depth 3

**Setup:** Change a shortcut's target to a non-existent path. Actual exe exists within 3 folder levels of the nearest real parent.  
**Action:** Run the script.  
**Expected:** `shortcut target missing or invalid` warning, then `searching for <ExpectedExe>`, then `found replacement at <path>. Updating shortcut.` App launches (with `Arguments` if present).  
**Pass criteria:** Shortcut target updated. App running.

---

### TC-06 — Win32: shortcut target missing, exe not found within depth 3, user provides valid path

**Setup:** Shortcut target broken and exe not reachable within 3 levels.  
**Action:** When prompted, enter the full correct path to the exe.  
**Expected:** Shortcut updated. App launches (with `Arguments` if present).  
**Pass criteria:** Shortcut updated. App running.

---

### TC-07 — Win32 prompt: invalid paths entered before correct one

**Setup:** Same as TC-06.  
**Action:** Enter a non-existent path, then a path with the wrong filename, then the correct path.  
**Expected:** Appropriate warning on each bad input. Accepts only a valid matching path.  
**Pass criteria:** Prompt repeats on bad input. Succeeds on correct input.

---

### TC-08 — Win32 prompt: user presses Enter to skip

**Setup:** Same as TC-06.  
**Action:** Press Enter at the prompt.  
**Expected:** `[Name]: no valid executable path. Skipping.` App in failure list.  
**Pass criteria:** Script continues. Failed app listed at end.

---

### TC-09 — Win32: shortcut .lnk file itself does not exist

**Setup:** Delete or rename one shortcut file.  
**Action:** Run the script.  
**Expected:** `WARNING: [Name]: shortcut file not found: <path>`. App in failure list.  
**Pass criteria:** Script continues. Remaining apps launch.

---

### TC-10 — Win32: process does not appear within 30 seconds

**Setup:** Point a shortcut to a valid exe that does not produce a detectable process within 30 seconds.  
**Action:** Run the script.  
**Expected:** After 30 seconds: `WARNING: [Name]: '[ProcessName]' did not appear within 30 seconds.`  
**Pass criteria:** Script does not hang beyond 30 seconds per app. Failure logged.

---

### TC-11 — Appx: AUMID resolved via Get-StartApps (step 1)

**Setup:** Phone Link installed and visible in Start. `PhoneExperienceHost` not running.  
**Action:** Run the script.  
**Expected:** Console shows `AUMID resolved via Get-StartApps: <aumid>` then `launching via shell:appsFolder\<aumid>` then `'PhoneExperienceHost' is now running.`  
**Pass criteria:** Phone Link window visible. `PhoneExperienceHost` in Task Manager.

---

### TC-12 — Appx: AUMID resolved via KnownAumid verification (step 2)

**Setup:** Phone Link installed but NOT visible in `Get-StartApps`. `KnownAumid` package family present in `Get-AppxPackage`.  
**Action:** Run the script.  
**Expected:** Step 1 skipped. Console shows `KnownAumid verified as installed: <aumid>`. App launches.  
**Pass criteria:** `PhoneExperienceHost` running. AUMID matches `KnownAumid`.

---

### TC-13 — Appx: AUMID resolved via AppxPackage manifest (step 3)

**Setup:** Temporarily set `KnownAumid` to a fake package family. `Get-StartApps` does not return Phone Link.  
**Action:** Run the script.  
**Expected:** Steps 1 and 2 fail. Console shows `AUMID discovered via AppxPackage manifest: <aumid>`. App launches.  
**Pass criteria:** `PhoneExperienceHost` running. Manifest-discovered AUMID used.

---

### TC-14 — Appx: all AUMID resolution steps fail

**Setup:** Set `KnownAumid`, `AppxName`, and `StartAppName` to values matching nothing installed.  
**Action:** Run the script.  
**Expected:** All three steps fail. Console shows `no AUMID found. Skipping.` App in failure list.  
**Pass criteria:** Script continues. App in final failure summary.

---

### TC-15 — Appx: already running

**Setup:** Phone Link already open before running the script.  
**Action:** Run the script.  
**Expected:** Console shows `Phone Link: 'PhoneExperienceHost' already running. Skipping.` AUMID resolution not attempted.  
**Pass criteria:** No second instance. Single `PhoneExperienceHost` in Task Manager.

---

### TC-16 — Full sequence: all apps launch successfully

**Setup:** All Win32 shortcut targets valid, Phone Link installed, no apps pre-running.  
**Action:** Run the script.  
**Expected:** Each app launches sequentially. Sticky Notes opens a sticky note window (not full OneNote). Console ends with `Startup sequence completed successfully.`  
**Pass criteria:** All processes visible in Task Manager. No warnings or failure list.

---

### TC-17 — Win32: repaired shortcut persists on second run

**Setup:** Trigger TC-05 or TC-06 to repair a shortcut. Close the launched app.  
**Action:** Run the script a second time without changes.  
**Expected:** Repaired shortcut used directly. No repair logic triggered.  
**Pass criteria:** No `shortcut target missing` warning for the previously repaired app.

---

## Pass Criteria Summary

| TC | Scenario | Pass If |
|----|----------|---------|
| TC-01 | Valid Win32, no Arguments, not running | App launches within 30 s |
| TC-02 | Sticky Notes with Arguments | Sticky note window opens; `/memoryWindow start` in log |
| TC-03 | App already running | Skipped; no second instance |
| TC-04 | Sticky Notes then OneNote process order | OneNote skipped; one ONENOTE process |
| TC-05 | Broken Win32 target, exe in depth 3 | Shortcut updated; app launches with Arguments if present |
| TC-06 | Broken Win32 target, user provides path | Shortcut updated; app launches with Arguments if present |
| TC-07 | Invalid then valid paths at prompt | Prompt repeats; accepts correct input |
| TC-08 | User skips prompt | App in failure list; script continues |
| TC-09 | `.lnk` file missing | Warning logged; script continues |
| TC-10 | Win32 process timeout | Warning after 30 s; script continues |
| TC-11 | Appx AUMID via Get-StartApps | Correct AUMID logged; app launches |
| TC-12 | Appx AUMID via KnownAumid | KnownAumid confirmed; app launches |
| TC-13 | Appx AUMID via manifest | Manifest AUMID used; app launches |
| TC-14 | All AUMID steps fail | Warning logged; app in failure list |
| TC-15 | Appx already running | Skipped; no AUMID resolution attempted |
| TC-16 | Full sequence | All apps running; success message |
| TC-17 | Repaired shortcut reused | No repair on second run |
