# Windows 11 Startup Manager

A modular, folder-first startup manager for Windows 11. It launches desktop (Win32) applications and Universal Windows Platform (UWP) apps in a strict sequential order with genuine readiness verification, automatic self-repair, and interactive shortcut management.

---

## Key Features

1. **Folder-First Runtime (Zero-Config Default)**:
   * By default, the startup sequence reads directly from the designated Start Menu folder (`01-99 *.lnk`).
   * No JSON configuration is queried or required to launch startup items.
2. **First-Run Wizard & Dedicated Settings Menu**:
   * Prompts new users on first launch to confirm recommended generic defaults (`%APPDATA%` and local JSON) or configure custom paths.
   * Dedicated submenu (**Option 8**) allows changing the active Start Menu folder or pointing the JSON manifest to a cloud-synced folder (e.g., OneDrive, Dropbox) for multi-PC synchronization.
3. **Environment Variable Path Expansion**:
   * Supports portable paths like `%APPDATA%`, `%ProgramData%`, and `%USERPROFILE%`, expanding them dynamically on the host system without hardcoding user identifiers.
4. **Automatic Shortcut Listing**:
   * After any shortcut modification (Add, Remove, or Modify), the updated sequence of managed shortcuts is immediately printed to the console (Option 5 logic).
5. **Readiness Detection**:
   * Uses a two-stage verification process: polls for a valid top-level window handle (`MainWindowHandle`) first before considering an application running in the system tray or background.
   * Prevents premature launches caused by background processes.
6. **Automated Source Self-Repair**:
   * If a shortcut target executable is missing (e.g., following an application update), the engine searches the parent directory recursively up to 3 levels deep.
   * Automatically updates shortcut target paths upon finding the new binary.
   * Prompts with an interactive file dialog if automated recovery fails.
7. **Native UWP App Activation**:
   * Resolves AppX packages and Application User Model IDs (AUMIDs) dynamically using Windows AppX cmdlets (`Get-AppxPackage`, `Get-AppxPackageManifest`).
   * Launches packaged applications via `shell:AppsFolder\<AUMID>`, bypassing NTFS permissions on `WindowsApps`.
8. **On-Demand Menu-Driven Synchronization**:
   * **Option 6 (Folder → JSON)**: Exports the current numbered folder shortcuts into `Win11startupapps.json` as a backup, pruning entries for shortcuts that no longer exist on disk (after confirmation).
   * **Option 7 (JSON → Folder)**: Rebuilds any shortcuts listed in `Win11startupapps.json` that are missing from the Start Menu folder, restoring the managed sequence.
---

## File Structure

```
win11-startup/
├── Win11startup.ps1           # Core executable launcher script
├── Win11startupapps.json       # Optional JSON manifest (used for export/reverse-sync)
├── Win11startup.Tests.ps1     # Pester unit tests for core script functions
├── TESTING.md                 # Test execution and verification instructions
├── tasks.md                   # Development backlog and changelog
└── README.md                  # Project documentation
```

---

## Numbering & Shortcut Conventions

Shortcuts in the Start Menu folder must follow a two-digit prefix matching the regex:
```regex
^(0[1-9]|[1-9][0-9])\s
```

* **Valid examples:**
  * `01 Windows Terminal.lnk`
  * `02 Google Chrome.lnk`
  * `03 Slack.lnk`
* **Ignored files:** Shortcuts lacking a two-digit prefix or standard files are skipped during startup.

---

## Menu System

### Main Menu

| Option | Action | Description |
| :--- | :--- | :--- |
| **1** | **Launch all startup shortcuts** | Runs through `01-99` shortcuts in the folder sequentially, waiting for readiness. |
| **2** | **Add a shortcut** | Selects an executable, calculates the next sequence number, creates the `.lnk`, and auto-lists updated shortcuts. |
| **3** | **Remove a shortcut** | Deletes a managed `.lnk` file from disk and auto-lists remaining shortcuts. |
| **4** | **Modify a shortcut** | Rename an application, update its target executable, or change its sequence number (01-99), then auto-lists. |
| **5** | **List shortcuts in folder** | Displays all active numbered shortcuts and their target binaries. |
| **6** | **Sync: Folder -> JSON** | Exports the current folder shortcuts into `Win11startupapps.json` (pruning missing items upon confirmation). |
| **7** | **Reverse Sync: JSON -> Folder** | Recreates any shortcuts found in `Win11startupapps.json` that are missing from disk. |
| **8** | **Settings & Path Configuration** | Opens the dedicated configuration submenu to manage active folder and JSON paths. |
| **9** | **Quit** | Exits the application. |

### Settings & Path Configuration Submenu (Option 8)
* **1. View current active paths**: Displays the currently loaded Start Menu directory and active JSON manifest location.
* **2. Configure Start Menu folder**: Choose between Current User (`%APPDATA%`), All Users (`%ProgramData%`), or browse to a custom folder.
* **3. Configure JSON location**: Select script directory default or link to a cloud-synced folder (e.g., OneDrive) for multi-device sync.
* **4. Reset to recommended defaults**: Reverts folder and JSON paths to generic standard defaults.
* **5. Return to Main Menu**.

---

## Testing

Automated tests are implemented with the **Pester** testing framework. See [TESTING.md](TESTING.md) for full details.

To execute tests:
```powershell
Invoke-Pester .\Win11startup.Tests.ps1 -Output Detailed
```

