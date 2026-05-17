# Win11 Startup Launcher

A curated Windows 11 startup launcher that sequentially starts a fixed list of personal applications. Uses a `WshShell.Run`-based launch strategy for Win32 apps so baked-in shortcut arguments are always honoured, with self-healing shortcut repair as a fallback. Appx (Store/packaged) apps are launched via AUMID resolved dynamically at runtime. After each launch, the script automatically detects whether the app is window-based or tray-only — no per-app flags needed.

## Key Features

- Curated app list — only your chosen startup applications; no scanning of the Startup folder or Start Menu.
- Two launch strategies per entry via `LaunchType`:
  - `Win32` — shortcut invoked via `WshShell.Run` so arguments baked into the `.lnk` Target field are preserved exactly. Self-healing repair and user-prompt fallback activate only when the shortcut target is broken.
  - `Appx` — AUMID resolved at runtime in three stages; `KnownAumid` is the primary candidate only, not a hardcoded dependency.
- **Runtime presence-mode detection** — after launch, `Get-AppPresenceMode` polls `MainWindowHandle` for `$SettleSeconds` (default 5 s). Apps that produce a visible window are classified as `Window` mode; apps that run headless in the system tray are classified as `Tray` mode. No `WindowCheck` flag or per-app configuration is needed.
- `Test-AppAlreadyOpen` skips relaunch correctly for both Window and Tray apps — a tray app already in the process list is treated as open without requiring a visible window.
- Bounded Win32 search — depth fixed at 3 to avoid slow machine-wide crawls.
- Prompts user for exact executable path when automated Win32 repair fails; validates path before accepting.
- Inline failure menu appears when a shortcut is missing or an app times out — offering Add+retry, Modify, or Skip without restarting the sequence.
- App list externalised to `apps.json`; Add/Delete menu writes changes back automatically.
- Logs each outcome: running (skip), launched with detected mode, repaired/discovered-and-launched, or failed.

## App List

### Win32 entry fields

| Field | Required | Description |
|-------|----------|-------------|
| `Name` | Yes | Display name used in log output |
| `LaunchType` | Yes | `"Win32"` |
| `ShortcutPath` | Yes | Full path to the `.lnk` shortcut file |
| `ProcessName` | Yes | Process name used to detect if running and confirm launch (no `.exe`) |
| `ExpectedExe` | Yes | Exact executable filename used during shortcut repair and user-prompt validation |
| `ExpectedPublisher` | No | Authenticode signer CN string; verified before any repaired or user-supplied exe is persisted |
| `ExpectedArguments` | No | Expected `Arguments` field value in the `.lnk`; triggers argument self-healing when present (e.g. Phone Link AUMID) |
| `StartAppName` | No | Leave empty for Win32-only entries |
| `KnownAumid` | No | Leave empty for Win32-only entries |
| `AppxName` | No | Leave empty for Win32-only entries |

> Arguments needed at launch (e.g. `/memoryWindow start` for Sticky Notes) must be baked into the shortcut's Target field. The script invokes the `.lnk` via `WshShell.Run`, which passes them automatically.

### Appx entry fields

| Field | Required | Description |
|-------|----------|-------------|
| `Name` | Yes | Display name used in log output |
| `LaunchType` | Yes | `"Appx"` |
| `ShortcutPath` | Yes | Full path to the `.lnk` shortcut file (targets `explorer.exe`) |
| `ProcessName` | Yes | Process name used to confirm launch and detect if already running (no `.exe`) |
| `ExpectedExe` | Yes | Set to `explorer.exe` for all Appx entries |
| `KnownAumid` | Yes | Last-known AUMID (`PackageFamilyName!AppId`); verified at runtime — not assumed permanently valid |
| `AppxName` | Yes | Partial package name used in `Get-AppxPackage` discovery when `KnownAumid` is stale |
| `StartAppName` | Yes | Display name pattern used in `Get-StartApps` discovery (first resolution step) |

### Default entries

| # | App | LaunchType | Notes |
|---|-----|------------|-------|
| 01 | Outlook | Win32 | |
| 02 | Teams | Win32 | |
| 03 | OneDrive | Win32 | Tray app — classified automatically at runtime |
| 04 | Sticky Notes | Win32 | Shortcut Target includes `/memoryWindow start`; `ExpectedExe`/`ProcessName` = `ONENOTE.EXE`/`ONENOTE` |
| 05 | OneNote | Win32 | Shares `ONENOTE` process with entry 04; skipped if ONENOTE already running |
| 06 | Phone Link | Win32 | `ExpectedArguments`: `shell:appsFolder\Microsoft.YourPhone_8wekyb3d8bbwe!App`; argument self-healing scans WindowsApps for the package folder and reconstructs the AUMID from `AppxManifest.xml` |
| 07 | Microsoft Edge | Win32 | |
| 08 | Google Chrome | Win32 | |

## Win32 Launch Strategy

```
1. Test-AppAlreadyOpen (process running?)   -> skip if yes
2. Verify .lnk file exists                  -> inline failure menu if missing
3. Read shortcut TargetPath
4. Target valid?  -> check Arguments if ExpectedArguments set
   Arguments valid? -> invoke .lnk via WshShell.Run
   Arguments wrong? -> Repair-ShortcutArguments -> invoke repaired .lnk
5. Target broken? -> Repair-ShortcutTarget:
     a. Climb parent folders to find nearest existing folder
     b. Search downward max 3 levels for ExpectedExe
     c. Found? -> update shortcut, invoke repaired .lnk
     d. Not found? -> prompt user for exact exe path
     e. Valid input? -> update shortcut, invoke repaired .lnk
     f. Skipped? -> log failure, continue
6. Wait-ForAppReady:
     Phase 1 (SettleSeconds): Get-AppPresenceMode polls MainWindowHandle
       -> Window mode: continue to Phase 2
       -> Tray mode:   confirm ready immediately
     Phase 2 (remaining timeout): wait for MainWindowHandle (Window mode only)
7. Log success or timeout failure; inline failure menu on timeout
```

## Phone Link Argument Self-Healing

Phone Link is a packaged UWP app. Its `PhoneExperienceHost.exe` lives under `C:\Program Files\WindowsApps\` which is ACL-locked to `TrustedInstaller` — direct `.exe` invocation always fails. The correct and only reliable launch path is via `explorer.exe shell:appsFolder\<AUMID>`.

The script stores this as a Win32 `.lnk` targeting `explorer.exe` with the AUMID as the `Arguments` field. If the installed package version changes and the AUMID no longer matches, `Repair-ShortcutArguments` automatically:

1. Extracts the `PackageFamilyName` fragment from `ExpectedArguments`.
2. Scans `C:\Program Files\WindowsApps` for a matching folder (newest version preferred).
3. Reads `AppxManifest.xml` to confirm the `AppId`.
4. Reconstructs the AUMID and updates the shortcut `Arguments` field.

## Appx AUMID Resolution

```
1. Get-StartApps filtered by StartAppName  (most reliable, reflects current install state)
2. Verify KnownAumid package family is still installed via Get-AppxPackage
3. Get-AppxPackage by AppxName + read AppId from manifest
If all three fail -> app skipped, added to failure list
```

## Presence Mode Detection

`Get-AppPresenceMode` is called automatically after every launch. It polls `MainWindowHandle` for up to `$SettleSeconds` (default 5 s):

| Result | Meaning | Ready condition |
|--------|---------|----------------|
| `Window` | A visible window appeared within settle time | `MainWindowHandle != 0` within timeout |
| `Tray` | Process running but no window after settle time | Process presence alone is sufficient |
| `$null` | Process never appeared during settle time | Continues polling up to full timeout |

This removes the need for any `WindowCheck`, `TrayApp`, or equivalent per-app flag in the `$apps` table.

## Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `$startMenu` | System Start Menu Programs path | Base folder for all Win32 shortcut files |
| `$MaxRepairDepth` | `3` | Maximum folder depth searched during Win32 shortcut repair |
| `$InitialDelaySeconds` | `10` | Wait time after login before starting the sequence |
| `$LaunchTimeoutSeconds` | `30` | Maximum seconds to wait for a process to become ready after launch |
| `$PostLaunchPauseSeconds` | `2` | Pause between apps after a successful launch |
| `$SettleSeconds` | `5` | Seconds polled for `MainWindowHandle` to classify an app as Window or Tray mode |

## Automated Tests (Pester)

Unit tests are in `Win11startup.Tests.ps1`. They run in test mode (`$env:PS_STARTUP_TESTMODE = '1'`) which dot-sources the script without triggering the interactive menu or startup sequence.

```powershell
# Unit tests only
Invoke-Pester .\Win11startup.Tests.ps1

# Unit + integration tests (requires live Windows environment with shortcuts present)
$env:RUN_INTEGRATION = '1'; Invoke-Pester .\Win11startup.Tests.ps1
```

| Test ID | Covers |
|---------|--------|
| TEST-02 | `Get-RelativeDepth` — all depth/boundary cases |
| TEST-03 | `Find-MisnumberedShortcut` — match, no-match, empty folder, missing folder |
| TEST-04a | `Test-ExePathAllowed` — allowed roots, denied paths |
| TEST-04b | `Test-ExeSignatureTrusted` — valid sig, correct/wrong publisher, unsigned file |
| TEST-08 | `Import-AppsConfig` — valid load, optional field normalisation, missing required field, missing file |
| TEST-09 | `Get-NearestExistingParent` — deep path, empty string, immediate parent |
| TEST-10 | `Show-AppPicker -AllowNew` — `__NEW__` sentinel on `N`, `$null` on `0` |
| TEST-11 | `Add-Shortcut` dispatch — `$null` cancel, real app object re-init, `__NEW__` new-entry |
| TEST-12 | `Wait-ForAppReady` phase-1 clamping — `TimeoutSeconds < SettleSeconds`, normal case |
| INT-01 | Integration harness setup/teardown |
| INT-02 | `Initialize-Shortcut` smoke test with real `.lnk` |

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
| v8 | Removed entries 4 (ShareFile), 5 (Greenshot), 8 (SAP GUI), 9 (Notepad++) from Default Entries; remaining entries renumbered 01–08 |
| v9 | Replaced static `WindowCheck` flag with runtime presence-mode detection (`Get-AppPresenceMode`, `Test-AppAlreadyOpen`, `Wait-ForAppReady`); added `$SettleSeconds` config variable; Phone Link reclassified as Win32 with `ExpectedArguments` and argument self-healing |
| v10 | Added exe allowlist (`Test-ExePathAllowed`); restricted repair and user-prompt acceptance to paths under Program Files, Program Files (x86), or Windows |
| v11 | Added Authenticode signature gate (`Test-ExeSignatureTrusted`); added safer XML manifest loading via `[xml]::new() + Load()` |
| v12 | Added optional `ExpectedPublisher` per entry; publisher CN string verified against signer certificate subject during repair and user-prompt flows |
| v13 | Added `ExpectedExe` guard in `Test-AppAlreadyOpen`; prevents false skip caused by unrelated same-named processes |
| v14 | Anchored `Repair-ShortcutArguments` regex; pattern requires full `shell:appsFolder\<PFN>!<AppId>` form with `Microsoft.` prefix constraint |
| v15 | All shared variables moved to `$script:` scope; prevents Pester dot-source from leaking or shadowing globals |
| v16 | Added Pester test file (`Win11startup.Tests.ps1`); `$env:PS_STARTUP_TESTMODE` guard; TEST-02 through TEST-04 unit tests; INT-01/INT-02 integration harness |
| v17 | Externalised `$script:apps` to `apps.json` (`Import-AppsConfig` / `Export-AppsConfig`); Add and Delete menu flows persist changes automatically |
| v18 | FIX-01: Add-menu Appx support (`StartAppName`, `KnownAumid`, `AppxName` collected and persisted); FIX-02: `Show-AppPicker -AllowNew` with `__NEW__` sentinel; FIX-03: `Wait-ForAppReady` phase-1 clamping fixes phase-2 timeout math |
| v19 | TEST-08–12 Pester tests added (Import-AppsConfig, Get-NearestExistingParent, Show-AppPicker -AllowNew, Add-Shortcut dispatch, Wait-ForAppReady clamping); apps.json updated with Appx fields on all entries; README version history and field tables completed |

## License

See [LICENSE](../LICENSE) in the repository root.
