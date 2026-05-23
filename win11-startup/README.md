# Win11 Startup Launcher

A curated PowerShell startup launcher for Windows 11 that sequentially starts a fixed personal app list at login. Win32 apps launch via `WshShell.Run` on their `.lnk` shortcut (preserving baked-in arguments). Packaged/Store apps launch via AUMID resolved dynamically at runtime. Presence mode (window vs. tray) is detected automatically after each launch — no per-app flags needed. All configuration lives in `Win11startupapps.json`.

---

## Quick Start

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\Win11startup.ps1
```

To run at login, add a shortcut to this script in your Windows Startup folder (`shell:startup`).

**Requirements:** Windows 10/11 · PowerShell 5.1+ · `Get-AppxPackage` and `Get-StartApps` (standard inbox)

---

## Main Menu

| Option | Action |
|--------|--------|
| `[1]` Run startup sequence | Launches all configured apps in order |
| `[2]` Add shortcut | Re-initialise an existing entry or add a new one |
| `[3]` Delete shortcut | Remove a `.lnk` file and/or its config entry |
| `[4]` Modify shortcut | Fix a broken target path or stale AUMID arguments |
| `[5]` List startup apps | Health-check table — shows shortcut status per app |
| `[6]` Sync from Start Menu | Rebuild `Win11startupapps.json` from numbered `.lnk` files |
| `[7]` Exit | |

**Sync** is auto-triggered on first run when `Win11startupapps.json` is missing. It picks up any `.lnk` in `Start Menu\Programs` whose filename starts with 1–2 digits (e.g. `01 Outlook.lnk`).

---

## App Configuration (`Win11startupapps.json`)

### Win32 fields

| Field | Required | Description |
|-------|----------|-------------|
| `Name` | Yes | Display name |
| `LaunchType` | Yes | `"Win32"` |
| `ShortcutPath` | Yes | Full path to the `.lnk` file |
| `ProcessName` | Yes | Process name without `.exe` (used for skip-if-running check) |
| `ExpectedExe` | Yes | Exact exe filename used in repair and user-prompt validation |
| `ExpectedPublisher` | No | Authenticode signer CN string; verified before any repaired exe is saved |
| `ExpectedArguments` | No | Expected `.lnk` Arguments value; triggers argument self-healing when stale (e.g. `shell:appsFolder\<AUMID>` after a package update) |

> Win32 launch arguments (e.g. `/memoryWindow start` for Sticky Notes) must be baked into the shortcut's Target field — the script fires the `.lnk` via `WshShell.Run` as-is.

### Appx fields

| Field | Required | Description |
|-------|----------|-------------|
| `Name` | Yes | Display name |
| `LaunchType` | Yes | `"Appx"` |
| `ShortcutPath` | Yes | Full path to the `.lnk` file (targets `explorer.exe`) |
| `ProcessName` | Yes | Process name without `.exe` |
| `ExpectedExe` | Yes | `"explorer.exe"` for all Appx entries |
| `KnownAumid` | Yes | Last-known `PackageFamilyName!AppId`; verified at runtime, not assumed permanent |
| `AppxName` | Yes | Partial package name for `Get-AppxPackage` fallback |
| `StartAppName` | Yes | Display name fragment for `Get-StartApps` (first resolution attempt) |

---

## Key Behaviours

- **Self-healing shortcut repair** — when a `.lnk` target is broken, walks up to the grandparent folder and searches all subfolders recursively for the expected exe. Falls back to a user prompt (3 attempts, allowlist + Authenticode validated).
- **Argument self-healing** — when `ExpectedArguments` is set and the `.lnk` Arguments field is stale, `Repair-ShortcutArguments` re-reads `AppxManifest.xml` and rewrites the AUMID.
- **Presence-mode detection** — after launch, `Get-AppPresenceMode` polls `MainWindowHandle` for `$SettleSeconds`. Apps with a visible window → `Window` mode; headless tray apps → `Tray` mode.
- **Inline failure menu** — on missing shortcut or launch timeout, per-app options: fix & retry / modify another / skip / delete entry.
- **Exe security gates** — `Test-ExeAcceptable` requires path under Program Files / Windows and a valid Authenticode signature before any exe is written to a shortcut.

---

## Configuration Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `$InitialDelaySeconds` | `10` | Wait after login before starting the sequence |
| `$LaunchTimeoutSeconds` | `30` | Max seconds to wait for an app to become ready |
| `$PostLaunchPauseSeconds` | `2` | Pause between apps after a successful launch |
| `$SettleSeconds` | `5` | Seconds polled to classify Window vs. Tray mode |

---

## Tests (Pester)

```powershell
# Unit tests only
Invoke-Pester .\Win11startup.Tests.ps1

# Unit + integration (requires live Windows environment)
$env:RUN_INTEGRATION = '1'; Invoke-Pester .\Win11startup.Tests.ps1
```

Test mode is activated via `$env:PS_STARTUP_TESTMODE = '1'` — dot-sources the script cleanly without triggering the menu or startup sequence.

---

## License

See [LICENSE](../LICENSE) in the repository root.
