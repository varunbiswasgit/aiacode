# Win11startup.ps1

## Purpose

A curated Windows 11 startup launcher that sequentially starts a fixed list of personal applications. It replaces broad process-name guessing and UWP fallback logic with a targeted **self-healing shortcut repair** mechanism.

If a shortcut target path is missing, the script climbs upward through the folder hierarchy to find the nearest existing parent folder, then searches downward up to three levels for the expected executable. If the search still fails, the user is prompted to supply the exact executable path. In both recovery cases, the shortcut is updated before the app is launched, so the repair persists for future runs.

---

## Key Features

- Curated app list — only your chosen startup applications; no scanning of the Startup folder or Start Menu.
- Skips any app whose process is already running.
- Reads and validates each shortcut target before attempting launch.
- Climbs parent folders to recover from broken shortcut targets.
- Bounded search — depth fixed at 3 to avoid slow machine-wide crawls.
- Prompts user for exact executable path when automated repair fails; validates the path before accepting it.
- Rewrites the shortcut target and working directory after any repair.
- Logs each outcome: running (skip), launched, repaired-and-launched, or failed.
- No UWP or AppUserModelId logic — Win32 executables only.

---

## App List

Each entry in `$apps` requires four fields:

| Field | Description |
|-------|-------------|
| `Name` | Display name used in log output |
| `ShortcutPath` | Full path to the `.lnk` shortcut file |
| `ProcessName` | Process name used to check if already running (no `.exe`) |
| `ExpectedExe` | Exact executable filename used during shortcut repair and user-prompt validation |

Default entries cover: Outlook, Teams, OneDrive, ShareFile, Greenshot, Sticky Notes, OneNote, SAP GUI, Notepad++, Phone Link, Microsoft Edge, and Google Chrome.

> **Review `ExpectedExe` for each entry before first use.** The values for Teams, Sticky Notes, and Phone Link may differ depending on your installation type.

---

## Configuration

All tunable values are at the top of the script:

| Variable | Default | Description |
|----------|---------|-------------|
| `$startMenu` | System Start Menu Programs path | Base folder for all shortcut files |
| `$MaxRepairDepth` | `3` | Maximum folder depth searched during shortcut repair |
| `$InitialDelaySeconds` | `10` | Wait time after login before starting the sequence |
| `$LaunchTimeoutSeconds` | `30` | Maximum seconds to wait for a process to appear after launch |
| `$PostLaunchPauseSeconds` | `2` | Pause between apps after a successful launch |

---

## Shortcut Repair Logic

```
1. Read .lnk target path
2. Target exists?  → Launch normally
3. Target missing → Climb parent folders until an existing folder is found
4. Search that folder downward (max 3 levels) for ExpectedExe
5. Found?          → Update shortcut, launch
6. Not found?      → Prompt user for exact executable path
7. Valid input?    → Update shortcut, launch
8. Skipped?        → Log failure, continue to next app
```

---

## User Prompt Validation

When prompting for a manual path, the script rejects input if:
- The path does not exist.
- The path points to a folder rather than a file.
- The file name does not match `ExpectedExe` exactly (case-insensitive).

The prompt repeats until a valid path is entered or the user presses Enter to skip.

---

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- Execution policy set to allow local scripts:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```
- All apps in `$apps` must be Win32 executables. UWP/Store-only apps are not supported.

---

## Usage

Run from PowerShell:

```powershell
.\Win11startup.ps1
```

To run automatically at login, add a shortcut pointing to this script in the Windows Startup folder (`shell:startup`).

---

## Version History

| Version | Change |
|---------|--------|
| v1 | Direct shortcut launch; 30-second process wait |
| v2 | Added UWP fallback map, Start Menu search, broad exe search, retry prompt |
| v3 | Removed UWP; replaced broad search with self-healing shortcut repair (depth 3); added validated user-prompt fallback |
