# Win11 Startup Manager
# Launches configured apps at startup via numbered .lnk shortcuts.
# Win32 apps  : launched via WshShell.Run; self-healing shortcut repair on broken target/args.
# Appx apps   : AUMID resolved at runtime (Get-StartApps -> KnownAumid -> AppxPackage manifest).
# Bootstrap   : ensures every Win32 .lnk exists before launch; renames misnumbered or creates fresh.
# Config      : Win11startupapps.json (same folder). Add/Delete/Modify via main menu.
# Test mode   : set $env:PS_STARTUP_TESTMODE = '1' to dot-source without running the menu/sequence.
