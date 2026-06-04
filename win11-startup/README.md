# Win11 Startup Launcher (Simplified)

A lightweight PowerShell startup launcher for Windows 11 that sequentially launches numbered `.lnk` shortcuts found in a Start Menu folder. Launch order is driven by a numeric prefix on each shortcut filename (e.g., `01 Outlook.lnk`, `02 Teams.lnk`). Configuration is persisted automatically to a local JSON file. Windows Store / UWP apps (shortcuts targeting `explorer.exe shell:appsFolder\...`) are skipped.

---

## Quick Start

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\Win11startup.ps1
```

On first run the script prompts for:
1. A config JSON path (defaults to `Win11StartupConfig.json` next to the script).
2. The Start Menu folder that contains your numbered `.lnk` files.

To run at Windows login, add a shortcut to this script in your Startup folder (`shell:startup`).

**Requirements:** Windows 10/11 · PowerShell 5.1+ · `WScript.Shell` COM object (standard inbox)

---

## How It Works

| Step | What happens |
|------|--------------|
| 1 | Loads (or creates) `Win11StartupConfig.json` storing `StartMenuPath` and a `Shortcuts` array. |
| 2 | Reads all `.lnk` files whose base name begins with 1–2 digits (e.g., `01 Outlook.lnk`) from the configured Start Menu folder. |
| 3 | Sorts shortcuts by their numeric prefix. |
| 4 | For each shortcut: resolves the target executable, skips Store apps, checks whether the process is already running, then launches via `WshShell.Run`. |
| 5 | Waits up to `$ProcessStartTimeout` seconds for the process to appear. |
| 6 | If the process never starts, presents a Windows **Open File** dialog so you can select the correct `.exe` — the shortcut is then repaired and re-launched. |

---

## Configuration File (`Win11StartupConfig.json`)

The file is created automatically on first run and updated live as shortcuts are discovered.

```json
{
  "StartMenuPath": "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs",
  "Shortcuts": [
    {
      "Name": "Outlook",
      "ShortcutPath": "C:\\...\\01 Outlook.lnk",
      "ProcessName": "OUTLOOK"
    }
  ]
}
```

| Field | Description |
|-------|-------------|
| `StartMenuPath` | Full path to the folder containing numbered `.lnk` files. |
| `Shortcuts[].Name` | Display name derived from the shortcut filename (numeric prefix stripped). |
| `Shortcuts[].ShortcutPath` | Full path to the `.lnk` file. |
| `Shortcuts[].ProcessName` | Process name (without `.exe`) used to detect whether the app is already running. Auto-updated when a shortcut is repaired. |

---

## Shortcut Naming Convention

Shortcuts must start with a 1–2 digit number followed by a space:

```
01 Outlook.lnk
02 Microsoft Teams.lnk
03 Slack.lnk
```

Launch order is numeric (ascending). Shortcuts without a numeric prefix are ignored.

---

## Key Behaviours

- **Skip-if-running** — process already running? Shortcut is silently skipped.
- **Store app guard** — shortcuts targeting `explorer.exe` with `shell:appsFolder\` arguments are skipped (UWP launch via `WshShell.Run` is unsupported).
- **Self-healing shortcut repair** — if an app does not start within `$ProcessStartTimeout` seconds, a file-picker dialog lets you choose the correct `.exe`. The shortcut target and config are updated in place.
- **Config auto-sync** — every shortcut encountered updates or adds its entry in `Win11StartupConfig.json` automatically.

---

## Configuration Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `$ProcessStartTimeout` | `15` | Seconds to wait for a launched process to appear before offering repair. |

---

## Tests (Pester)

```powershell
# Unit tests only
Invoke-Pester .\Win11startup.Tests.ps1
```

Test mode is activated by setting `$env:PS_STARTUP_TESTMODE = '1'` before dot-sourcing the script, which suppresses the startup sequence.

> **Note:** Tests for functions not present in this simplified script (e.g., `Import-AppsConfig`, `Resolve-Aumid`, `Invoke-AppLaunch`) are not applicable to this version.

---

## License

See [LICENSE](../LICENSE) in the repository root.
