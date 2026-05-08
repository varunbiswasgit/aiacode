# Win11startup.ps1

## Purpose

A curated Windows 11 startup launcher that sequentially starts a fixed list of personal applications. It uses a self-healing shortcut repair mechanism for Win32 apps and a packaged-app shell identity launch for apps that do not expose a reliable Win32 executable path (Phone Link).

If a Win32 shortcut target is missing, the script climbs upward through the folder hierarchy to the nearest existing parent folder, then searches downward up to three levels for the expected executable. If the search still fails, the user is prompted to supply the exact path. In both recovery cases, the shortcut is updated before launch so the repair persists.

---

## Key Features

- Curated app list — only your chosen startup applications; no scanning of the Startup folder or Start Menu.
- Two launch strategies controlled per entry by `LaunchType`:
  - `Win32` — shortcut-based with depth-3 self-healing repair and validated user-prompt fallback.
  - `Appx` — packaged app launched via `explorer.exe shell:appsFolder\...` identity.
- Skips any app whose process is already running.
- Reads and validates each Win32 shortcut target before attempting launch.
- Climbs parent folders to recover from broken Win32 shortcut targets.
- Bounded search — depth fixed at 3 to avoid slow machine-wide crawls.
- Prompts user for exact executable path when automated repair fails; validates the path before accepting.
- Rewrites the shortcut target and working directory after any repair.
- Logs each outcome: running (skip), launched, repaired-and-launched, or failed.

---

## App List

### Win32 entries

Each Win32 entry requires four fields:

| Field | Description |
|-------|-------------|
| `Name` | Display name used in log output |
| `LaunchType` | `"Win32"` |
| `ShortcutPath` | Full path to the `.lnk` shortcut file |
| `ProcessName` | Process name used to check if already running (no `.exe`) |
| `ExpectedExe` | Exact executable filename used during shortcut repair and user-prompt validation |

### Appx entries

Each Appx entry requires three fields:

| Field | Description |
|-------|-------------|
| `Name` | Display name used in log output |
| `LaunchType` | `"Appx"` |
| `AppCommand` | Shell app folder identity, e.g. `shell:appsFolder\Microsoft.YourPhone_8wekyb3d8bbwe!App` |
| `ProcessName` | Process name used to confirm launch (no `.exe`) |

### Default entries

| # | App | LaunchType | Notes |
|---|-----|------------|-------|
| 01 | Outlook | Win32 | |
| 02 | Teams | Win32 | |
| 03 | OneDrive | Win32 | |
| 04 | ShareFile | Win32 | |
| 05 | Greenshot | Win32 | |
| 06 | Sticky Notes | Win32 | Launched via `ONENOTE.EXE /memoryWindow start`; ProcessName is `ONENOTE` |
| 07 | OneNote | Win32 | Launched via `ONENOTE.EXE`; shares `ONENOTE` process name with Sticky Notes |
| 08 | SAP GUI | Win32 | |
| 09 | Notepad++ | Win32 | |
| 10 | Phone Link | Appx | Launched via `shell:appsFolder\Microsoft.YourPhone_8wekyb3d8bbwe!App` |
| 11 | Microsoft Edge | Win32 | |
| 12 | Google Chrome | Win32 | |

> **Sticky Notes note:** Both Sticky Notes (entry 06) and OneNote (entry 07) share the `ONENOTE` process name because Sticky Notes is launched via `ONENOTE.EXE /memoryWindow start`. If ONENOTE is already running when entry 06 is processed, entry 06 will be skipped. The shortcut for entry 06 should have `ONENOTE.EXE` as target and `/memoryWindow start` as arguments.

---

## Configuration

All tunable values are at the top of the script:

| Variable | Default | Description |
|----------|---------|-------------|
| `$startMenu` | System Start Menu Programs path | Base folder for all Win32 shortcut files |
| `$MaxRepairDepth` | `3` | Maximum folder depth searched during Win32 shortcut repair |
| `$InitialDelaySeconds` | `10` | Wait time after login before starting the sequence |
| `$LaunchTimeoutSeconds` | `30` | Maximum seconds to wait for a process to appear after launch |
| `$PostLaunchPauseSeconds` | `2` | Pause between apps after a successful launch |

---

## Win32 Shortcut Repair Logic

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

## User Prompt Validation (Win32 only)

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
- Win32 apps in `$apps` must expose a Win32 executable. Appx apps are launched by shell identity and do not need an exe path.

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
| v4 | Added `LaunchType` field per app; Phone Link now launched via packaged-app shell identity; `Start-Win32App` and `Start-AppxApp` split into separate functions |
