# Win11startup.ps1 — Testing Readme

This document covers manual test cases for `Win11startup.ps1`. There is no automated Pester test file for this script because its core behavior involves process detection, shortcut COM objects, and user prompts, which require a live Windows environment. All tests below are manual.

For environment setup (execution policy, PowerShell version), see [tests/README.md](README.md).

---

## Test Environment

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- All shortcut files present in `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\` with the numbered naming convention (`01 Outlook.lnk`, etc.)

---

## Test Cases

### TC-01 — Normal launch: shortcut target valid, process not running

**Setup:** Shortcut target path exists. App is not running.  
**Action:** Run the script.  
**Expected:** App launches. Console shows `[Name]: launching <path>` and `[Name]: '[ProcessName]' is now running.`  
**Pass criteria:** App process visible in Task Manager within 30 seconds.

---

### TC-02 — Already running: process detected before launch

**Setup:** Launch one app (e.g., Notepad++) manually before running the script.  
**Action:** Run the script.  
**Expected:** Console shows `[Name]: '[ProcessName]' already running. Skipping.` No second instance is launched.  
**Pass criteria:** Only one instance of the process appears in Task Manager.

---

### TC-03 — Shortcut target missing, exe found within depth 3

**Setup:** Open a shortcut's properties and change the target to a non-existent path (e.g., append `_old` to the folder name). Ensure the actual exe exists within 3 folder levels of the nearest real parent.  
**Action:** Run the script.  
**Expected:** Console shows `shortcut target missing or invalid`, then `searching for <ExpectedExe>`, then `found replacement at <path>. Updating shortcut.` App launches.  
**Pass criteria:** Shortcut target is updated to the found path. App process is running.

---

### TC-04 — Shortcut target missing, exe not found within depth 3, user provides valid path

**Setup:** Change shortcut target to a path where no parent folder exists on the drive, or where the exe is more than 3 levels below the nearest real parent.  
**Action:** Run the script. When prompted, type the full correct path to the exe.  
**Expected:** Console prompts `Enter the full path for [Name] ([ExpectedExe]), or press Enter to skip`. After valid input, shortcut is updated and app launches.  
**Pass criteria:** Shortcut target updated. App process running.

---

### TC-05 — User prompt: invalid path entered, then valid path entered

**Setup:** Same as TC-04.  
**Action:** At the prompt, first enter a path that does not exist, then enter a path to the correct file but with the wrong file name, then enter the correct path.  
**Expected:**
- First input: `WARNING: Path does not exist or is not a file: ...`
- Second input: `WARNING: File name must be exactly [ExpectedExe]`
- Third input: shortcut updated, app launches.

**Pass criteria:** Prompt repeats on bad input. Accepts only a path that exists, is a file, and has the correct file name.

---

### TC-06 — User prompt: user presses Enter to skip

**Setup:** Same as TC-04.  
**Action:** At the prompt, press Enter without typing a path.  
**Expected:** Console shows `[Name]: no valid executable path. Skipping.` App appears in the final failure list.  
**Pass criteria:** Script continues to the next app without crashing. Failed app listed at end.

---

### TC-07 — Shortcut .lnk file itself does not exist

**Setup:** Rename or delete one shortcut file from the Start Menu folder.  
**Action:** Run the script.  
**Expected:** Console shows `WARNING: [Name]: shortcut file not found: <path>`. App appears in the final failure list.  
**Pass criteria:** Script does not crash. Remaining apps continue to launch normally.

---

### TC-08 — Process does not appear within 30 seconds

**Setup:** Point a shortcut target to a valid executable that does not produce a detectable process within 30 seconds.  
**Action:** Run the script.  
**Expected:** After 30 seconds, console shows `WARNING: [Name]: '[ProcessName]' did not appear within 30 seconds.` App added to failure list.  
**Pass criteria:** Script does not hang beyond 30 seconds per app. Failure logged correctly.

---

### TC-09 — Full sequence: all apps launch successfully

**Setup:** All shortcut targets valid, no apps pre-running.  
**Action:** Run the script.  
**Expected:** Each app launches sequentially. Console ends with `Startup sequence completed successfully.`  
**Pass criteria:** All 12 processes visible in Task Manager. No warnings or failure list.

---

### TC-10 — Repaired shortcut persists on second run

**Setup:** Trigger TC-03 or TC-04 to repair a shortcut. After the script completes, close the launched app.  
**Action:** Run the script a second time without changing anything.  
**Expected:** On the second run, the repaired shortcut target is used directly (TC-01 path). No repair logic is triggered.  
**Pass criteria:** No `shortcut target missing` warning on the second run for the previously repaired app.

---

## Pass Criteria Summary

| TC | Scenario | Pass If |
|----|----------|---------|
| TC-01 | Valid shortcut, app not running | App launches within 30 s |
| TC-02 | App already running | Skipped; no second instance |
| TC-03 | Broken target, exe found in depth 3 | Shortcut updated; app launches |
| TC-04 | Broken target, user provides valid path | Shortcut updated; app launches |
| TC-05 | User enters bad paths before correct one | Prompt repeats correctly |
| TC-06 | User skips prompt | App in failure list; script continues |
| TC-07 | `.lnk` file missing | Warning logged; script continues |
| TC-08 | Process timeout | Warning after 30 s; script continues |
| TC-09 | Full sequence | All apps running; success message |
| TC-10 | Repaired shortcut reused | No repair triggered on second run |
