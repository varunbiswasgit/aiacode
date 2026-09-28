# Windows 11 Startup Manager - Task Tracker & Changelog

## Completed Tasks

- [x] **Folder-First Architecture**: Decoupled default launch sequence (Option 1) from JSON configuration; Start Menu folder is now the primary runtime source of truth.
- [x] **First-Run Setup & Settings Submenu (Option 8)**: Implemented initial first-run setup wizard and dedicated settings submenu for path configuration (User vs. Machine folder, custom paths, and external/cloud JSON manifest location).
- [x] **Dynamic Environment Variable Expansion**: Added `Expand-PathString` to dynamically expand `%APPDATA%`, `%ProgramData%`, and `%USERPROFILE%` paths safely across different machines.
- [x] **Auto-Listing on Mutation**: Implemented automatic invocation of Option 5 display logic immediately following shortcut Add, Remove, or Modify actions.
- [x] **Readiness Polling Overhaul**: Fixed premature background/tray state triggers (`SessionId -gt 0`) by enforcing active window handle verification before evaluating background state.
- [x] **UWP / AppX Native Invocation**: Resolved WindowsApps NTFS ACL errors by integrating `Get-AppxPackage` and `Get-AppxPackageManifest` for AUMID extraction and launching via `shell:AppsFolder`.
- [x] **Bounded Self-Repair**: Added recursion depth limits and drive root protection to `Repair-BrokenShortcutSource` to prevent unbounded disk scans.
- [x] **User-Menu Driven Sync**:
  - Implemented Option 6: `Sync-FolderToJson` (folder-to-JSON export/backup).
  - Implemented Option 7: `Sync-JsonToFolder` (reverse sync: rebuilding disk shortcuts from JSON).
- [x] **Automated Test Suite**: Created Pester unit tests in `Win11startup.Tests.ps1` covering path expansion, pattern matching, process resolution, numbering, and config loading.
- [x] **Documentation Alignment**: Rewrote `README.md` and `TESTING.md` to reflect exact repository file names, folder-first operation, and settings management.

---

## Planned / Backlog

- [ ] Add toast notification support upon startup completion.
- [ ] Add support for delay intervals between sequential application launches.

