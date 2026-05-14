# Win11 Startup Launcher

A curated Windows 11 startup launcher that sequentially starts a fixed list of personal applications. Uses a `WshShell.Run`-based launch strategy for Win32 apps so baked-in shortcut arguments are always honoured, with self-healing shortcut repair as a fallback. Appx (Store/packaged) apps are launched via AUMID resolved dynamically at runtime.

## Key Features

- Curated app list — only your chosen startup applications; no scanning of the Startup folder or Start Menu.
- Two launch strategies per entry via `LaunchType`:
  - `Win32` — shortcut invoked via `WshShell.Run` so arguments baked into the `.lnk` Target field are preserved exactly. Self-healing repair and user-prompt fallback activate only when the shortcut target is broken.
  - `Appx` — AUMID resolved at runtime in three stages; `KnownAumid` is the primary candidate only, not a hardcoded dependency.
- Skips any app whose process is already running.
- Bounded Win32 search — depth fixed at 3 to avoid slow machine-wide crawls.
- Prompts user for exact executable path when automated Win32 repair fails; validates path before accepting.
- Logs each outcome: running (skip), launched, repaired/discovered-and-launched, or failed.

## App List

### Win32 entry fields

| Field | Required | Description |
|-------|----------|-------------|
| `Name` | Yes | Display name used in log output |
| `LaunchType` | Yes | `"Win32"` |
| `ShortcutPath` | Yes | Full path to the `.lnk` shortcut file |
| `ProcessName` | Yes | Process name used to detect if running and confirm launch (no `.exe`) |
| `ExpectedExe` | Yes | Exact executable filename used during shortcut repair and user-prompt validation |

> Arguments needed at launch (e.g. `/memoryWindow start` for Sticky Notes) must be baked into the shortcut's Target field. The script invokes the `.lnk` via `WshShell.Run`, which passes them automatically.

### Appx entry fields

| Field | Required | Description |
|-------|----------|-------------|
| `Name` | Yes | Display name used in log output |
| `LaunchType` | Yes | `"Appx"` |
| `KnownAumid` | Yes | Last-known AUMID (`PackageFamilyName!AppId`); verified at runtime — not assumed permanently valid |
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
| 06 | Sticky Notes | Win32 | Shortcut Target includes `/memoryWindow start`; `ExpectedExe`/`ProcessName` = `ONENOTE.EXE`/`ONENOTE` |
| 07 | OneNote | Win32 | Shares `ONENOTE` process with entry 06; skipped if ONENOTE already running |
| 08 | SAP GUI | Win32 | |
| 09 | Notepad++ | Win32 | |
| 10 | Phone Link | Appx | `KnownAumid`: `Microsoft.YourPhone_8wekyb3d8bbwe!App`; resolved dynamically at runtime |
| 11 | Microsoft Edge | Win32 | |
| 12 | Google Chrome | Win32 | |

## Win32 Launch Strategy

```
1. Check if process is already running  -> skip if yes
2. Verify .lnk file exists              -> log failure if missing
3. Read shortcut TargetPath
4. Target valid?  -> invoke .lnk via WshShell.Run
5. Target broken? -> Repair-ShortcutTarget:
     a. Climb parent folders to find nearest existing folder
     b. Search downward max 3 levels for ExpectedExe
     c. Found? -> update shortcut, invoke repaired .lnk
     d. Not found? -> prompt user for exact exe path
     e. Valid input? -> update shortcut, invoke repaired .lnk
     f. Skipped? -> log failure, continue
6. Wait up to 30 s for ProcessName to appear
7. Log success or timeout failure
```

## Appx AUMID Resolution

```
1. Get-StartApps filtered by StartAppName  (most reliable, reflects current install state)
2. Verify KnownAumid package family is still installed via Get-AppxPackage
3. Get-AppxPackage by AppxName + read AppId from manifest
If all three fail -> app skipped, added to failure list
```

## Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `$startMenu` | System Start Menu Programs path | Base folder for all Win32 shortcut files |
| `$MaxRepairDepth` | `3` | Maximum folder depth searched during Win32 shortcut repair |
| `$InitialDelaySeconds` | `10` | Wait time after login before starting the sequence |
| `$LaunchTimeoutSeconds` | `30` | Maximum seconds to wait for a process to appear after launch |
| `$PostLaunchPauseSeconds` | `2` | Pause between apps after a successful launch |

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later
- Execution policy set to allow local scripts:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```
- `Get-AppxPackage` and `Get-StartApps` available (standard in Windows 10/11 PowerShell)

## Usage

```powershell
.\Win11startup.ps1
```

To run automatically at login, add a shortcut to this script in the Windows Startup folder (`shell:startup`).

## Version History

| Version | Change |
|---------|--------|
| v1 | Direct shortcut launch; 30-second process wait |
| v2 | Added UWP fallback map, Start Menu search, broad exe search, retry prompt |
| v3 | Removed UWP; replaced broad search with self-healing shortcut repair (depth 3); validated user-prompt fallback |
| v4 | Added `LaunchType` per entry; Phone Link launched via packaged-app shell identity |
| v5 | Appx AUMID resolved dynamically (Get-StartApps -> KnownAumid verification -> AppxPackage manifest) |
| v6 | Added optional `Arguments` field; Sticky Notes launched with `/memoryWindow start` |
| v7 | Removed `Arguments` field; Win32 apps invoked via `WshShell.Run` on `.lnk` so baked-in arguments are preserved automatically |

## License

See [LICENSE](../LICENSE) in the repository root.
