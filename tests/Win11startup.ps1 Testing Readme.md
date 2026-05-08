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

---

## Test Cases

### TC-01 — Normal Win32 launch: shortcut target valid, process not running

**Setup:** Shortcut target path exists. App is not running.  
**Action:** Run the script.  
**Expected:** App launches. Console shows `[Name]: launching <path>` and `[Name]: '[ProcessName]' is now running.`  
**Pass criteria:** App process visible in Task Manager within 30 seconds.

---

### TC-02 — Already running: process detected before launch

**Setup:** Launch one app (e.g., Notepad++) manually before running the script.  
**Action:** Run the script.  
**Expected:** Console shows `[Name]: '[ProcessName]' already running. Skipping.` No second instance launched.  
**Pass criteria:** Only one instance of the process in Task Manager.

---

### TC-03 — Win32: shortcut target missing, exe found within depth 3

**Setup:** Change a shortcut's target to a non-existent path. Actual exe exists within 3 folder levels of the nearest real parent.  
**Action:** Run the script.  
**Expected:** `shortcut target missing or invalid` warning, then `searching for <ExpectedExe>`, then `found replacement at <path>. Updating shortcut.` App launches.  
**Pass criteria:** Shortcut target updated. App running.

---

### TC-04 — Win32: shortcut target missing, exe not found within depth 3, user provides valid path

**Setup:** Shortcut target broken and exe not reachable within 3 levels.  
**Action:** When prompted, enter the full correct path to the exe.  
**Expected:** Shortcut updated. App launches.  
**Pass criteria:** Shortcut updated. App running.

---

### TC-05 — Win32 prompt: invalid paths entered before correct one

**Setup:** Same as TC-04.  
**Action:** Enter a non-existent path, then a path with the wrong filename, then the correct path.  
**Expected:** Appropriate warning on each bad input. Accepts only a valid matching path.  
**Pass criteria:** Prompt repeats on bad input. Succeeds on correct input.

---

### TC-06 — Win32 prompt: user presses Enter to skip

**Setup:** Same as TC-04.  
**Action:** Press Enter at the prompt.  
**Expected:** `[Name]: no valid executable path. Skipping.` App in failure list.  
**Pass criteria:** Script continues. Failed app listed at end.

---

### TC-07 — Win32: shortcut .lnk file itself does not exist

**Setup:** Delete or rename one shortcut file.  
**Action:** Run the script.  
**Expected:** `WARNING: [Name]: shortcut file not found: <path>`. App in failure list.  
**Pass criteria:** Script continues. Remaining apps launch.

---

### TC-08 — Win32: process does not appear within 30 seconds

**Setup:** Point a shortcut to a valid exe that does not produce a detectable process within 30 seconds.  
**Action:** Run the script.  
**Expected:** After 30 seconds: `WARNING: [Name]: '[ProcessName]' did not appear within 30 seconds.`  
**Pass criteria:** Script does not hang beyond 30 seconds per app. Failure logged.

---

### TC-09 — Appx: AUMID resolved via Get-StartApps (step 1)

**Setup:** Phone Link installed and visible in Start. `PhoneExperienceHost` not running.  
**Action:** Run the script.  
**Expected:** Console shows `AUMID resolved via Get-StartApps: <aumid>` then `launching via shell:appsFolder\<aumid>` then `'PhoneExperienceHost' is now running.`  
**Pass criteria:** Phone Link window visible. `PhoneExperienceHost` in Task Manager. AUMID in log matches `Get-StartApps | Where-Object { $_.Name -like '*Phone Link*' }` output.

---

### TC-10 — Appx: AUMID resolved via KnownAumid verification (step 2)

**Setup:** Phone Link installed but NOT visible in `Get-StartApps` (rare; can simulate by temporarily removing the Start pin). `KnownAumid` package family IS present in `Get-AppxPackage`.  
**Action:** Run the script.  
**Expected:** Step 1 skipped. Console shows `KnownAumid verified as installed: <aumid>`. App launches.  
**Pass criteria:** `PhoneExperienceHost` running. AUMID matches `KnownAumid` value.

---

### TC-11 — Appx: AUMID resolved via AppxPackage manifest (step 3)

**Setup:** Simulate a stale `KnownAumid` by temporarily changing it in the script to a fake package family name. `Get-StartApps` does not return Phone Link.  
**Action:** Run the script.  
**Expected:** Steps 1 and 2 fail with warnings. Console shows `AUMID discovered via AppxPackage manifest: <aumid>`. App launches with the discovered AUMID.  
**Pass criteria:** `PhoneExperienceHost` running. Manifest-discovered AUMID used.

---

### TC-12 — Appx: all AUMID resolution steps fail

**Setup:** Simulate a fully unknown app by setting `KnownAumid`, `AppxName`, and `StartAppName` to values that match nothing installed.  
**Action:** Run the script.  
**Expected:** All three steps fail with warnings. Console shows `no AUMID found. Skipping.` App added to failure list.  
**Pass criteria:** Script continues. App in final failure summary.

---

### TC-13 — Appx: already running

**Setup:** Phone Link already open before running the script.  
**Action:** Run the script.  
**Expected:** Console shows `Phone Link: 'PhoneExperienceHost' already running. Skipping.` AUMID resolution is not attempted.  
**Pass criteria:** No second instance. Single `PhoneExperienceHost` in Task Manager.

---

### TC-14 — Full sequence: all apps launch successfully

**Setup:** All Win32 shortcut targets valid, Phone Link installed, no apps pre-running.  
**Action:** Run the script.  
**Expected:** Each app launches sequentially. Console ends with `Startup sequence completed successfully.`  
**Pass criteria:** All processes visible in Task Manager. No warnings or failure list.

---

### TC-15 — Win32: repaired shortcut persists on second run

**Setup:** Trigger TC-03 or TC-04 to repair a shortcut. After completion, close the launched app.  
**Action:** Run the script a second time without changes.  
**Expected:** Repaired shortcut target used directly (TC-01 path). No repair triggered.  
**Pass criteria:** No `shortcut target missing` warning for the previously repaired app.

---

## Pass Criteria Summary

| TC | Scenario | Pass If |
|----|----------|---------|
| TC-01 | Valid Win32 shortcut, not running | App launches within 30 s |
| TC-02 | App already running | Skipped; no second instance |
| TC-03 | Broken Win32 target, exe in depth 3 | Shortcut updated; app launches |
| TC-04 | Broken Win32 target, user provides path | Shortcut updated; app launches |
| TC-05 | Invalid then valid paths at prompt | Prompt repeats; accepts correct input |
| TC-06 | User skips prompt | App in failure list; script continues |
| TC-07 | `.lnk` file missing | Warning logged; script continues |
| TC-08 | Win32 process timeout | Warning after 30 s; script continues |
| TC-09 | Appx AUMID via Get-StartApps | Correct AUMID logged; app launches |
| TC-10 | Appx AUMID via KnownAumid verification | KnownAumid confirmed; app launches |
| TC-11 | Appx AUMID via manifest | Manifest AUMID used; app launches |
| TC-12 | All AUMID steps fail | Warning logged; app in failure list |
| TC-13 | Appx already running | Skipped; no AUMID resolution attempted |
| TC-14 | Full sequence | All apps running; success message |
| TC-15 | Repaired shortcut reused | No repair on second run |
