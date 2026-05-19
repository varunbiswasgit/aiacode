# Win11 Startup Launcher

A curated Windows 11 startup launcher that sequentially starts a fixed list of personal applications at login. Win32 apps are invoked via `WshShell.Run` on their `.lnk` shortcut so baked-in arguments are always honoured. Packaged (Store/MSIX) apps are launched via AUMID resolved dynamically at runtime. After each launch the script automatically classifies the app as window-based or tray-only — no per-app flags needed. The app list is externalised to `Win11startupapps.json` and all menu operations persist changes back to it automatically.

---

## Key Features

- Curated app list — only your chosen apps; no scanning of the Windows Startup folder.
- Two launch strategies per entry via `LaunchType`:
  - `Win32` — `.lnk` invoked via `WshShell.Run`; baked-in shortcut arguments are preserved exactly. Self-healing repair and user-prompt fallback activate only when the shortcut target is broken.
  - `Appx` — AUMID resolved at runtime in three stages; `KnownAumid` is the primary candidate, not a hardcoded dependency.
- **Runtime presence-mode detection** — `Get-AppPresenceMode` polls `MainWindowHandle` for `$SettleSeconds` (default 5 s) after every launch. Apps with a visible window → `Window` mode; apps running headless in the tray → `Tray` mode. No per-app flag required.
- `Test-AppAlreadyOpen` correctly skips both window apps and tray apps already in the process list.
- Self-healing shortcut repair — walks up to the grandparent of a broken target folder, then searches all subfolders recursively for the expected executable.
- Shortcut argument self-healing — when `ExpectedArguments` is set, stale `shell:appsFolder\<AUMID>` values in the shortcut are reconstructed automatically after a package update.
- Inline failure menu (Add+retry / Modify / Skip) appears per-app during the sequence — no need to restart.
- `Win11startupapps.json` is the single source of truth; Add, Delete, and Sync menu operations write changes back automatically.
- Logs every outcome with timestamps to `startup-error.log` in the script folder.

---

## Main Menu

```
================================================
  Win11 Startup Manager
================================================
  [1] Run startup sequence
  [2] Add shortcut
  [3] Delete shortcut
  [4] Modify shortcut
  [5] List startup apps
  [6] Sync from Start Menu
  [7] Exit
------------------------------------------------
```

### Menu [1] — Run Startup Sequence

**When to use:** Normal daily login — launches all configured apps in order.

**Workflow:**

1. Script waits `$InitialDelaySeconds` (default 10 s) for the desktop to stabilise.
2. **Bootstrap phase** — for every `Win32` entry, `Initialize-Shortcut` checks whether the `.lnk` exists at `ShortcutPath`:
   - If missing, checks the same folder for a misnumbered variant (e.g. `8 Teams.lnk` instead of `02 Teams.lnk`) and renames it.
   - If no variant found and `ExpectedArguments` is set, creates the shortcut pointing to `explorer.exe` with the stored arguments.
   - If no variant found and no arguments, prompts for the exact `.exe` path, validates allowlist and Authenticode signature, then creates the shortcut.
3. **Launch loop** — each app is launched in sequence:
   - `Test-AppAlreadyOpen` — skips the app if already running (uses `RequireWindow` for `Window`-mode apps).
   - `WshShell.Run` fires the `.lnk` for Win32; `Start-Process explorer.exe shell:appsFolder\<AUMID>` for Appx.
   - `Wait-ForAppReady` detects presence mode and waits up to `$LaunchTimeoutSeconds` (default 30 s).
   - On success: logs ready, pauses `$PostLaunchPauseSeconds` (default 2 s), moves to next app.
   - On timeout or exception: **inline failure menu** appears (see below).
4. At the end, lists any apps that failed to start.

**Inline failure menu** (appears per-app mid-sequence):

```
  [1] Add / fix shortcut for <AppName> and retry
  [2] Modify a different shortcut
  [3] Skip
```

- `[1]` — runs `Initialize-Shortcut` (or `Edit-Shortcut` on timeout), then retries launch. Up to 3 total attempts per app.
- `[2]` — opens `Show-AppPicker` so you can fix a different app's shortcut without stopping the sequence.
- `[3]` — logs the app as failed and continues to the next app.

---

### Menu [2] — Add Shortcut

**When to use:** Add a new app to the startup list, or re-create a missing shortcut for an existing entry.

**Workflow — re-initialise an existing entry:**

1. `Show-AppPicker` lists all configured apps with shortcut status (`exists` / `missing`). Select a number.
2. `Initialize-Shortcut` runs: renames a misnumbered `.lnk` if found, creates a fresh one otherwise.
3. No changes to `Win11startupapps.json` — the entry already exists.

**Workflow — add a brand-new entry (`[N]`):**

1. Prompted for: display name, launch type (`Win32` / `Appx`), shortcut number (1–2 digits), process name, expected exe.
2. Optional: expected publisher CN string.
3. **Win32 with arguments** — prompted for `ExpectedArguments`; shortcut created pointing to `explorer.exe` with those arguments.
4. **Win32 without arguments** — prompted for the exact `.exe` path; allowlist + signature validated before shortcut is created.
5. **Appx** — prompted for `StartAppName`, `KnownAumid`, `AppxName`; shortcut created pointing to `explorer.exe shell:appsFolder\<AUMID>`.
6. New entry appended to `$script:apps`; `Win11startupapps.json` saved automatically.

---

### Menu [3] — Delete Shortcut

**When to use:** Remove an app from the startup list and/or delete its `.lnk` file.

**Workflow:**

1. `Show-AppPicker` lists all apps. Select the one to remove.
2. **Shortcut exists** — confirms deletion of the `.lnk` file, then asks separately whether to also remove the entry from `Win11startupapps.json`.
3. **Shortcut missing** — skips file deletion, asks directly whether to remove the entry from `Win11startupapps.json`.
4. If confirmed, entry is removed and `Win11startupapps.json` is saved.

---

### Menu [4] — Modify Shortcut

**When to use:** Fix or update the target of an existing shortcut — e.g. after an app reinstalls to a new path, or when a packaged-app AUMID has gone stale.

**Workflow — entry has `ExpectedArguments` set (packaged app):**

1. If the `.lnk` is missing, `Initialize-Shortcut` creates it first.
2. `Repair-ShortcutArguments` scans `C:\Program Files\WindowsApps` for the matching package folder, reads `AppxManifest.xml`, reconstructs the AUMID, and rewrites the shortcut `Arguments` field.

**Workflow — standard Win32 entry:**

1. Prompts for the exact `.exe` path (up to 3 attempts).
2. Validates: file exists, filename matches `ExpectedExe`, path is under an allowed root (Program Files / Program Files (x86) / Windows), Authenticode signature is `Valid`, publisher CN matches `ExpectedPublisher` if set.
3. If the `.lnk` exists, updates `TargetPath`. If missing, creates a new shortcut.

---

### Menu [5] — List Startup Apps

**When to use:** Quick health check — see all configured apps, their type, shortcut status, and process name at a glance.

**Output:**

```
================================================
  Configured Startup Apps (8 total)
================================================
#    Name                   Type   Shortcut   Process
---  ---------------------  -----  --------   ---------------
1    Outlook                Win32  OK         OUTLOOK
2    Teams                  Win32  OK         ms-teams
3    OneDrive               Win32  OK         OneDrive
...
```

- `OK` printed in green; `MISSING` printed in yellow.
- No changes made to files or configuration.

---

### Menu [6] — Sync from Start Menu

**When to use:**
- First run on a new machine when `Win11startupapps.json` does not yet exist (auto-triggered).
- After manually adding or renaming numbered `.lnk` files in the Start Menu Programs folder.
- To rebuild `Win11startupapps.json` from scratch based on what is currently in Start Menu.

**What "numbered" means:** any `.lnk` whose filename starts with 1–2 digits followed by a space — e.g. `01 Outlook.lnk`, `7 Chrome.lnk`.

**Workflow:**

1. Scans `C:\ProgramData\Microsoft\Windows\Start Menu\Programs` for `.lnk` files matching the `^\d{1,2}\s` pattern, sorted by name.
2. For each shortcut:
   - Target is `explorer.exe` + `shell:appsFolder\*` arguments → classified as `Appx`; `KnownAumid`, `StartAppName`, `AppxName` populated automatically.
   - Target is any other `.exe` → classified as `Win32`; `ProcessName` and `ExpectedExe` derived from the filename; existing `Arguments` stored as `ExpectedArguments`.
   - Unexpected or blank target → classified as `Win32` with a warning to review manually.
3. Entries constructed via `New-AppEntry` (consistent field set, `PresenceMode = $null`).
4. `$script:apps` replaced with the new list; `Win11startupapps.json` written immediately.
5. Session continues without restarting — `$script:apps` is live.

> **Note:** `ProcessName` for Appx entries cannot be determined without running the app. Fill it in via Menu [4] — Modify after the sync.

---

### Menu [7] — Exit

Exits immediately. No changes made.

---

## App List (`Win11startupapps.json`)

### Win32 entry fields

| Field | Required | Description |
|-------|----------|-------------|
| `Name` | Yes | Display name used in log output |
| `LaunchType` | Yes | `"Win32"` |
| `ShortcutPath` | Yes | Full path to the `.lnk` shortcut file |
| `ProcessName` | Yes | Process name used to detect if running and confirm launch (no `.exe`) |
| `ExpectedExe` | Yes | Exact executable filename used during shortcut repair and user-prompt validation |
| `ExpectedPublisher` | No | Authenticode signer CN string; verified before any repaired or user-supplied exe is persisted |
| `ExpectedArguments` | No | Expected `Arguments` value in the `.lnk`; triggers argument self-healing when the shortcut arguments become stale (e.g. a `shell:appsFolder\<AUMID>` value for a packaged app launched via `explorer.exe`) |
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
| 06 | Phone Link | Win32 | `ExpectedArguments`: `shell:appsFolder\Microsoft.YourPhone_8wekyb3d8bbwe!App`; argument self-healing reconstructs AUMID from `AppxManifest.xml` after a package update |
| 07 | Microsoft Edge | Win32 | |
| 08 | Google Chrome | Win32 | |

---

## Win32 Launch Strategy

```
1. Test-AppAlreadyOpen (process running?)   -> skip if yes
2. Verify .lnk file exists                  -> inline failure menu if missing
3. Read shortcut TargetPath
4. Target valid?
   -> ExpectedArguments set? -> check shortcut Arguments
      Arguments valid?   -> invoke .lnk via WshShell.Run
      Arguments stale?   -> Repair-ShortcutArguments -> invoke repaired .lnk
   -> No ExpectedArguments -> invoke .lnk via WshShell.Run directly
5. Target broken? -> Repair-ShortcutTarget:
     a. Walk up to the grandparent of the broken target's folder
     b. Search all subfolders recursively for ExpectedExe
     c. Found? -> validate allowlist + signature -> update shortcut -> invoke repaired .lnk
     d. Not found? -> prompt user for exact exe path
     e. Valid input? -> update shortcut -> invoke repaired .lnk
     f. Skipped or blocked? -> return $null -> log failure
6. Wait-ForAppReady:
     Phase 1 (SettleSeconds): Get-AppPresenceMode polls MainWindowHandle
       -> Window mode: continue to Phase 2
       -> Tray mode:   confirm ready immediately
     Phase 2 (remaining timeout): wait for MainWindowHandle (Window mode only)
7. Log success or timeout failure; inline failure menu on timeout
   Up to 3 total attempts per app before giving up
```

---

## Shortcut Argument Self-Healing

Packaged UWP and MSIX apps install under `C:\Program Files\WindowsApps\`, which is ACL-locked to `TrustedInstaller`. Direct `.exe` invocation is blocked. These apps must be launched via `explorer.exe shell:appsFolder\<AUMID>`.

The script stores them as Win32 `.lnk` entries targeting `explorer.exe`, with `shell:appsFolder\<AUMID>` in both the shortcut `Arguments` field and `ExpectedArguments`. When Windows updates the package and the version token in the AUMID changes, `Repair-ShortcutArguments` rebuilds it automatically:

1. Extracts the `PackageFamilyName` fragment from `ExpectedArguments` using an anchored regex.
2. Scans `C:\Program Files\WindowsApps` for a matching folder (newest version preferred).
3. Reads `AppxManifest.xml` to confirm the current `AppId`.
4. Reconstructs `<PackageFamilyName>!<AppId>` and rewrites the shortcut `Arguments` field.

This applies to any entry in `Win11startupapps.json` whose `ExpectedArguments` starts with `shell:appsFolder\` — not only Phone Link.

---

## Appx AUMID Resolution

```
1. Get-StartApps filtered by StartAppName   (most reliable — reflects current install state)
2. Verify KnownAumid package family is still installed via Get-AppxPackage
3. Get-AppxPackage by AppxName + read AppId from AppxManifest.xml
If all three fail -> app skipped, added to failure list, warning written to startup-error.log
```

---

## Presence Mode Detection

`Get-AppPresenceMode` is called automatically after every launch. It polls `MainWindowHandle` for up to `$SettleSeconds` (default 5 s):

| Result | Meaning | Ready condition |
|--------|---------|----------------|
| `Window` | A visible window appeared within settle time | `MainWindowHandle != 0` within remaining timeout |
| `Tray` | Process running but no window after settle time | Process presence alone is sufficient |
| `$null` | Process never appeared during settle time | Continues polling up to full timeout |

No `WindowCheck`, `TrayApp`, or equivalent per-app flag is needed.

---

## Configuration

| Variable | Default | Description |
|----------|---------|-------------|
| `$startMenu` | `C:\ProgramData\Microsoft\Windows\Start Menu\Programs` | Base folder scanned by Sync and used for all shortcut paths |
| `$InitialDelaySeconds` | `10` | Wait time after login before starting the launch sequence |
| `$LaunchTimeoutSeconds` | `30` | Maximum seconds to wait for a process to become ready after launch |
| `$PostLaunchPauseSeconds` | `2` | Pause between apps after a successful launch |
| `$SettleSeconds` | `5` | Seconds polled for `MainWindowHandle` to classify an app as Window or Tray mode |

---

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
| TEST-02 | `Get-RelativeDepth` (test-file only) — all depth/boundary cases |
| TEST-03 | `Find-MisnumberedShortcut` — match, no-match, empty folder, missing folder |
| TEST-04a | `Test-ExePathAllowed` — allowed roots, denied paths |
| TEST-04b | `Test-ExeSignatureTrusted` — valid sig, correct/wrong publisher, unsigned file |
| TEST-08 | `Import-AppsConfig` — valid load, optional field normalisation, missing required field, missing file |
| TEST-09 | `Get-AncestorNLevelsUp` — exact 3-level climb, boundary at filesystem root, non-existent path |
| TEST-10 | `Show-AppPicker -AllowNew` — `__NEW__` sentinel on `N`, `$null` on `0` |
| TEST-11 | `Add-Shortcut` dispatch — `$null` cancel, real app object re-init, `__NEW__` new-entry |
| TEST-12 | `Wait-ForAppReady` phase-1 clamping — `TimeoutSeconds < SettleSeconds`, normal case |
| INT-01 | Integration harness setup/teardown |
| INT-02 | `Initialize-Shortcut` smoke test with real `.lnk` |

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

```powershell
.\Win11startup.ps1
```

To run automatically at login, add a shortcut to this script in the Windows Startup folder (`shell:startup`).

---

## Version History

| Version | Change |
|---------|--------|
| v1 | Direct shortcut launch; 30-second process wait |
| v2 | Added UWP fallback map, Start Menu search, broad exe search, retry prompt |
| v3 | Removed UWP; replaced broad search with self-healing shortcut repair (depth 3); validated user-prompt fallback |
| v4 | Added `LaunchType` per entry; Phone Link launched via packaged-app shell identity |
| v5 | Appx AUMID resolved dynamically (Get-StartApps → KnownAumid verification → AppxPackage manifest) |
| v6 | Added optional `Arguments` field; Sticky Notes launched with `/memoryWindow start` |
| v7 | Removed `Arguments` field; Win32 apps invoked via `WshShell.Run` on `.lnk` so baked-in arguments are preserved automatically |
| v8 | Removed entries 4 (ShareFile), 5 (Greenshot), 8 (SAP GUI), 9 (Notepad++) from Default Entries; remaining entries renumbered 01–08 |
| v9 | Replaced static `WindowCheck` flag with runtime presence-mode detection; added `$SettleSeconds`; Phone Link reclassified as Win32 with `ExpectedArguments` and argument self-healing |
| v10 | Added exe allowlist (`Test-ExePathAllowed`); restricted repair and user-prompt to Program Files, Program Files (x86), or Windows |
| v11 | Added Authenticode signature gate (`Test-ExeSignatureTrusted`); safer XML manifest loading via `[xml]::new() + Load()` |
| v12 | Added optional `ExpectedPublisher` per entry; publisher CN verified against signer certificate subject |
| v13 | Added `ExpectedExe` guard in `Test-AppAlreadyOpen`; prevents false skip by unrelated same-named processes |
| v14 | Anchored `Repair-ShortcutArguments` regex; accepts any valid PackageFamilyName prefix |
| v15 | All shared variables moved to `$script:` scope |
| v16 | Added Pester test file; `$env:PS_STARTUP_TESTMODE` guard; TEST-02 through TEST-04; INT-01/INT-02 |
| v17 | Externalised `$script:apps` to `Win11startupapps.json` (`Import-AppsConfig` / `Export-AppsConfig`) |
| v18 | FIX-01: Appx add support; FIX-02: `Show-AppPicker -AllowNew` with `__NEW__` sentinel; FIX-03: phase-1 clamping fix |
| v19 | TEST-08–12 added; `Win11startupapps.json` updated with Appx fields on all entries; README completed |
| v20 | Shortcut argument self-healing generalised to any `ExpectedArguments`-bearing entry |
| v21 | `Repair-ShortcutArguments` regex generalised to accept any valid PackageFamilyName prefix |
| v22 | Repair search: `Get-AncestorNLevelsUp` replaced `Get-NearestExistingParent`; downward search fully recursive; TEST-09 updated |
| v23 | SYNC-01: `Sync-AppsFromStartMenu` inlined; menu `[6]` added; `Sync-AppsJson.ps1` retired; config renamed to `Win11startupapps.json` throughout |
| v24 | AUD-01–07: `Get-RelativeDepth` removed from main script; `Repair-ShortcutTarget` explicit `return $null` after blocked repair; `Sync-AppsFromStartMenu` branches merged; `New-AppEntry` helper extracted; retry loop comment corrected; `Export-AppsConfig` message condensed |

---

## License

See [LICENSE](../LICENSE) in the repository root.
