# Win11startup.ps1

## Purpose

A curated Windows 11 startup launcher that sequentially starts a fixed list of personal applications. It uses a `WshShell.Run`-based launch strategy for Win32 apps so that baked-in shortcut arguments are always honoured, with self-healing shortcut repair as a fallback. Appx (Store/packaged) apps are launched via AUMID resolved dynamically at runtime.

---

## Key Features

- Curated app list — only your chosen startup applications; no scanning of the Startup folder or Start Menu.
- Two launch strategies controlled per entry by `LaunchType`:
  - `Win32` — shortcut invoked via `WshShell.Run` so that any arguments baked into the `.lnk` Target field are preserved exactly as Windows would launch them. Self-healing repair and user-prompt fallback activate only when the shortcut target is broken; the repaired shortcut is then also invoked via `WshShell.Run`.
  - `Appx` — AUMID resolved at runtime in three stages; `KnownAumid` is the primary candidate only, not a hardcoded dependency.
- Skips any app whose process is already running.
- Win32: validates each shortcut target before launch; repairs and persists the shortcut if the target has moved, then invokes the updated shortcut.
- Appx: discovers the current AUMID dynamically; does not assume the package identity is stable across updates.
- Bounded Win32 search — depth fixed at 3 to avoid slow machine-wide crawls.
- Prompts user for exact executable path when Win32 automated repair fails; validates path before accepting.
- Logs each outcome: running (skip), launched, repaired/discovered-and-launched, or failed.

---

## App List

### Win32 entry fields

| Field | Required | Description |
|-------|----------|-------------|
| `Name` | Yes | Display name used in log output |
| `LaunchType` | Yes | `"Win32"` |
| `ShortcutPath` | Yes | Full path to the `.lnk` shortcut file |
| `ProcessName` | Yes | Process name used to check if already running and to confirm launch (no `.exe`) |
| `ExpectedExe` | Yes | Exact executable filename used during shortcut repair and user-prompt validation |

> No `Arguments` field exists. Any arguments needed at launch (e.g. `/memoryWindow start` for Sticky Notes) must be baked into the shortcut's Target field. The script invokes the `.lnk` via `WshShell.Run`, which passes them automatically.

### Appx entry fields

| Field | Required | Description |
|-------|----------|-------------|
| `Name` | Yes | Display name used in log output |
| `LaunchType` | Yes | `"Appx"` |
| `KnownAumid` | Yes | Last-known AUMID (`PackageFamilyName!AppId`); verified against installed packages at runtime — not assumed permanently valid |
| `AppxName` | Yes | Partial package name used in `Get-AppxPackage` discovery when `KnownAumid` is stale |
| `StartAppName` | Yes | Display name pattern used in `Get-StartApps` discovery (first resolution step) |
| `ProcessName` | Yes | Process name used to confirm launch and detect if already running (no `.exe`) |

### Default entries

| # | App | LaunchType | Notes |
|---|-----|------------|-------|
| 01 | Outlook | Win32 | |
| 02 | Teams | Win32 | |
| 03 | OneDrive | Win32 | |
| 04 | ShareFile | Win32 | |
| 05 | Greenshot | Win32 | |
| 06 | Sticky Notes | Win32 | Shortcut Target includes `/memoryWindow start`; `ExpectedExe` and `ProcessName` are `ONENOTE.EXE`/`ONENOTE` |
| 07 | OneNote | Win32 | Shares `ONENOTE` process name with entry 06; skipped if ONENOTE is already running |
| 08 | SAP GUI | Win32 | |
| 09 | Notepad++ | Win32 | |
| 10 | Phone Link | Appx | `KnownAumid`: `Microsoft.YourPhone_8wekyb3d8bbwe!App`; resolved dynamically at runtime |
| 11 | Microsoft Edge | Win32 | |
| 12 | Google Chrome | Win32 | |

> **Sticky Notes / OneNote process conflict:** Both entries share the `ONENOTE` process name. Entry 06 (Sticky Notes) fires the shortcut which carries `/memoryWindow start`. Entry 07 (OneNote) fires its own shortcut with no extra arguments. Because process detection is name-based, if Sticky Notes has already started ONENOTE by the time entry 07 is processed, OneNote is skipped with "already running". This is expected — both run in the same ONENOTE process.

---

## Win32 Launch Strategy

```
1. Check if process is already running  -> skip if yes
2. Verify .lnk file exists              -> log failure if missing
3. Read shortcut TargetPath
4. Target valid?  -> invoke .lnk via WshShell.Run (baked-in arguments preserved)
5. Target broken? -> Repair-ShortcutTarget:
     a. Climb parent folders to find nearest existing folder
     b. Search downward max 3 levels for ExpectedExe
     c. Found? -> update shortcut, invoke repaired .lnk via WshShell.Run
     d. Not found? -> prompt user for exact exe path
     e. Valid input? -> update shortcut, invoke repaired .lnk via WshShell.Run
     f. Skipped? -> log failure, continue
6. Wait up to 30 s for ProcessName to appear
7. Log success or timeout failure
```

---

## Appx AUMID Resolution

For any `Appx` entry, the script resolves the AUMID in three steps at runtime:

```
1. Get-StartApps filtered by StartAppName
   -> Reflects current installed state; most reliable source.

2. Verify KnownAumid is still installed
   -> Extracts PackageFamilyName from KnownAumid and queries Get-AppxPackage.
   -> Uses KnownAumid only if the package family is confirmed present.

3. Get-AppxPackage filtered by AppxName + manifest read
   -> Finds the installed package by partial name.
   -> Reads AppId from the package manifest.
   -> Constructs AUMID as PackageFamilyName!AppId.

If all three steps fail, the app is skipped and added to the failure list.
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

## User Prompt Validation (Win32 repair only)

When prompting for a manual path during repair, the script rejects input if:
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
| v6 | Added optional `Arguments` field to Win32 entries; Sticky Notes launched with `/memoryWindow start` via `Start-Process -ArgumentList` |
| v7 | Removed `Arguments` field; Win32 apps now invoked via `WshShell.Run` on the `.lnk` file so baked-in shortcut arguments are preserved automatically; `Resolve-LaunchPath` renamed to `Repair-ShortcutTarget` to reflect its sole purpose |
