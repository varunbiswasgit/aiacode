# Win11startup.ps1

## Purpose

A curated Windows 11 startup launcher that sequentially starts a fixed list of personal applications. It uses a self-healing shortcut repair mechanism for Win32 apps, and dynamic AUMID resolution at runtime for Appx (Store/packaged) apps so that the script is not dependent on any single hardcoded identity string remaining valid across OS or app updates.

---

## Key Features

- Curated app list — only your chosen startup applications; no scanning of the Startup folder or Start Menu.
- Two launch strategies controlled per entry by `LaunchType`:
  - `Win32` — shortcut-based with depth-3 self-healing repair and validated user-prompt fallback.
  - `Appx` — AUMID resolved at runtime in three stages; `KnownAumid` is the primary candidate only, not a hardcoded dependency.
- Skips any app whose process is already running.
- Win32: reads and validates each shortcut target before launch; repairs and persists the shortcut if the target has moved.
- Appx: discovers the current AUMID dynamically; does not assume the package identity is stable across updates.
- Bounded Win32 search — depth fixed at 3 to avoid slow machine-wide crawls.
- Prompts user for exact executable path when Win32 automated repair fails; validates path before accepting.
- Logs each outcome: running (skip), launched, repaired/discovered-and-launched, or failed.

---

## App List

### Win32 entry fields

| Field | Description |
|-------|-------------|
| `Name` | Display name used in log output |
| `LaunchType` | `"Win32"` |
| `ShortcutPath` | Full path to the `.lnk` shortcut file |
| `ProcessName` | Process name used to check if already running and to confirm launch (no `.exe`) |
| `ExpectedExe` | Exact executable filename used during shortcut repair and user-prompt validation |

### Appx entry fields

| Field | Description |
|-------|-------------|
| `Name` | Display name used in log output |
| `LaunchType` | `"Appx"` |
| `KnownAumid` | Last-known AUMID (`PackageFamilyName!AppId`); verified against installed packages at runtime — not assumed to be permanently valid |
| `AppxName` | Partial package name used in `Get-AppxPackage` discovery when `KnownAumid` is stale |
| `StartAppName` | Display name pattern used in `Get-StartApps` discovery (first resolution step) |
| `ProcessName` | Process name used to confirm launch and detect if already running (no `.exe`) |

### Default entries

| # | App | LaunchType | Notes |
|---|-----|------------|-------|
| 01 | Outlook | Win32 | |
| 02 | Teams | Win32 | |
| 03 | OneDrive | Win32 | |
| 04 | ShareFile | Win32 | |
| 05 | Greenshot | Win32 | |
| 06 | Sticky Notes | Win32 | Launched via `ONENOTE.EXE`; `ProcessName` is `ONENOTE` |
| 07 | OneNote | Win32 | Shares `ONENOTE` process name with entry 06; if ONENOTE is running, entry 06 is skipped |
| 08 | SAP GUI | Win32 | |
| 09 | Notepad++ | Win32 | |
| 10 | Phone Link | Appx | `KnownAumid`: `Microsoft.YourPhone_8wekyb3d8bbwe!App`; resolved dynamically at runtime |
| 11 | Microsoft Edge | Win32 | |
| 12 | Google Chrome | Win32 | |

---

## Appx AUMID Resolution

For any `Appx` entry, the script resolves the AUMID in three steps at runtime:

```
1. Get-StartApps filtered by StartAppName
   -> Reflects current installed state; most reliable source.
   -> Returns AppID directly if found.

2. Verify KnownAumid is still installed
   -> Extracts PackageFamilyName from KnownAumid and queries Get-AppxPackage.
   -> Uses KnownAumid only if the package family is confirmed present.
   -> Warns and skips to step 3 if the family is not found.

3. Get-AppxPackage filtered by AppxName + manifest read
   -> Finds the installed package by partial name.
   -> Reads AppId from the package manifest.
   -> Constructs AUMID as PackageFamilyName!AppId.
   -> Prefers 'App' as AppId when present; falls back to first declared AppId.

If all three steps fail, the app is skipped and added to the failure list.
```

This means a Store update that changes the package version or publisher but keeps the app name will still resolve correctly via step 1 or step 3, without any script changes.

---

## Win32 Shortcut Repair Logic

```
1. Read .lnk target path
2. Target exists?  -> Launch normally
3. Target missing  -> Climb parent folders until an existing folder is found
4. Search that folder downward (max 3 levels) for ExpectedExe
5. Found?          -> Update shortcut, launch
6. Not found?      -> Prompt user for exact executable path
7. Valid input?    -> Update shortcut, launch
8. Skipped?        -> Log failure, continue to next app
```

---

## Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `$startMenu` | System Start Menu Programs path | Base folder for all Win32 shortcut files |
| `$MaxRepairDepth` | `3` | Maximum folder depth searched during Win32 shortcut repair |
| `$InitialDelaySeconds` | `10` | Wait time after login before starting the sequence |
| `$LaunchTimeoutSeconds` | `30` | Maximum seconds to wait for a process to appear after launch |
| `$PostLaunchPauseSeconds` | `2` | Pause between apps after a successful launch |

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
- `Get-AppxPackage` and `Get-StartApps` available (standard in Windows 10/11 PowerShell)

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
| v3 | Removed UWP; replaced broad search with self-healing shortcut repair (depth 3); validated user-prompt fallback |
| v4 | Added `LaunchType` per entry; Phone Link launched via packaged-app shell identity; `Start-Win32App` and `Start-AppxApp` split into separate functions |
| v5 | Appx AUMID resolved dynamically at runtime (Get-StartApps -> KnownAumid verification -> AppxPackage manifest); `KnownAumid`, `AppxName`, `StartAppName` replace static `AppCommand` |
