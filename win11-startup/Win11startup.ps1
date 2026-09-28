# ===========================================================================
# Windows 11 Startup Manager (Win11startup.ps1)
# ---------------------------------------------------------------------------
#   - First-run setup detection & dedicated Settings/Configuration submenu
#   - Default folder-first runtime execution (zero JSON dependency on launch)
#   - Sequential launch of numbered shortcuts (01-99) from the Start Menu
#   - Window & Tray readiness verification (prevents premature launches)
#   - Automated source self-repair for moved/broken executables
#   - Native UWP resolution and activation via shell:AppsFolder
#   - Auto-list shortcuts after Add, Modify, or Remove actions (Option 5 logic)
#   - User-menu-driven bidirectional sync:
#       * Option 6: Folder -> JSON (Snapshot / export folder to config)
#       * Option 7: Reverse Sync: JSON -> Folder (Rebuild shortcuts from config)
#   - Option 8: Settings & Path Configuration (Folder & JSON path management)
# ===========================================================================

[CmdletBinding()]
# ---------------------------------------------------------------------------
# Path Helper: Environment Variable Expansion
# ---------------------------------------------------------------------------
function Expand-PathString {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    return [System.Environment]::ExpandEnvironmentVariables($Path)
}

param()

# ---------------------------------------------------------------------------
# Default Paths & Patterns
# ---------------------------------------------------------------------------
$script:DefaultUserStartMenu = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Windows\Start Menu\Programs')
$script:DefaultMachineStartMenu = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs'
$script:DefaultStartMenuFolder = if (Test-Path -LiteralPath $script:DefaultUserStartMenu) {
    $script:DefaultUserStartMenu
    $script:DefaultMachineStartMenu
}

$ScriptRoot        = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$defaultConfigPath = Join-Path $ScriptRoot 'Win11startupapps.json'
$StartMenuFolder   = $script:DefaultStartMenuFolder
$ActiveConfigPath  = $defaultConfigPath
$WshShell          = New-Object -ComObject WScript.Shell

$ProcessStartTimeout = 15
$WindowReadyTimeout  = 20
$script:NumberedLnkPattern = '^(0[1-9]|[1-9][0-9])\s'

# ---------------------------------------------------------------------------
# Configuration Persistence & Validation (For On-Demand Sync)
# ---------------------------------------------------------------------------
function Save-Config {
    param(
        [Parameter(Mandatory = $true)] $Cfg,
        [Parameter(Mandatory = $true)] [string]$Path
    )
    try {
        $Cfg | ConvertTo-Json -Depth 5 | Out-File -FilePath $Path -Encoding UTF8 -Force
    } catch {
        Write-Warning "Failed to save configuration to '${Path}': $($_.Exception.Message)"
    }
}

function Load-ConfigSafe {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        try {
            $parsed = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
            if ($parsed -and ($parsed | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name) -contains 'StartMenuPath') {
                return [PSCustomObject]@{
                    StartMenuPath = Expand-PathString -Path ([string]$parsed.StartMenuPath)
                    Shortcuts     = if ($parsed.Shortcuts) { @($parsed.Shortcuts) } else { @() }
                }
            }
        } catch {
            Write-Warning "Config file at '$Path' is invalid or unreadable."
        }
    }
    return [PSCustomObject]@{
        StartMenuPath = $script:DefaultStartMenuFolder
        Shortcuts     = @()
    }
}

# ---------------------------------------------------------------------------
# Process & Window Identification Helpers
# ---------------------------------------------------------------------------
function Get-ProcName {
    param(
        [string]$TargetPath,
        [string]$DisplayName
    )
    $base = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    if ([string]::IsNullOrWhiteSpace($base) -or $base -ieq 'explorer') {
        $base = ($DisplayName -replace '[^\w]', '')
    }
    return $base
}

function Get-AppReadyState {
    param([string]$ProcessName)

    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if (-not $procs) { return 'NotReady' }

    $hasWindow = $procs | Where-Object { $_.MainWindowHandle -ne 0 -and -not [string]::IsNullOrEmpty($_.MainWindowTitle) }
    if ($hasWindow) { return 'Window' }

    return 'RunningNoWindow'
}

function Wait-ForAppReady {
    param(
        [string]$ProcessName,
        [int]$ProcessTimeout,
        [int]$WindowTimeout,
        [bool]$IsDisplayNameFallback = $false
    )

    if ($IsDisplayNameFallback -or [string]::IsNullOrWhiteSpace($ProcessName)) {
        return 'NotReady'
    }

    $found = $false
    for ($i = 0; $i -lt $ProcessTimeout; $i++) {
        if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) {
            $found = $true
            break
        }
        Start-Sleep -Seconds 1
    }

    if (-not $found) { return 'NotReady' }

    Write-Host "  Process '$ProcessName' detected. Waiting for window initialization..." -ForegroundColor Cyan
    for ($j = 0; $j -lt $WindowTimeout; $j++) {
        $state = Get-AppReadyState -ProcessName $ProcessName
        if ($state -eq 'Window') { return 'Window' }
        Start-Sleep -Seconds 1
    }

    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if ($procs) {
        return 'Tray'
    }

    return 'NotReady'
}

# ---------------------------------------------------------------------------
# Shortcut File Management
# ---------------------------------------------------------------------------
function Update-Shortcut {
    param(
        [Parameter(Mandatory = $true)] $Shortcut,
        [Parameter(Mandatory = $true)] [string]$ExePath
    )
    try {
        $Shortcut.TargetPath       = $ExePath
        $Shortcut.Arguments        = ''
        $Shortcut.WorkingDirectory = Split-Path $ExePath -Parent
        $Shortcut.Save()
        Write-Host "  Shortcut updated -> $ExePath" -ForegroundColor Green
    } catch {
        Write-Warning "Could not update shortcut: $($_.Exception.Message)"
    }
}

function Get-NextShortcutNumber {
    param([string]$StartMenuFolder)
    $existing = Get-ChildItem -LiteralPath $StartMenuFolder -Filter '*.lnk' -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match $script:NumberedLnkPattern }
    $maxNum = 0
    foreach ($f in $existing) {
        if ($f.BaseName -match '^(\d{2})\s') {
            $n = [int]$Matches[1]
            if ($n -gt $maxNum) { $maxNum = $n }
        }
    }
    $next = $maxNum + 1
    if ($next -gt 99) {
        Write-Warning "Numbers 01-99 saturated. Assigning 99."
        return 99
    }
    return $next
}

function Select-ExecutableManually {
    param([string]$AppDisplayName)

    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Select executable for '$AppDisplayName'"
    $dlg.Filter = 'Executable Files (*.exe)|*.exe|All Files (*.*)|*.*'
    $dlg.InitialDirectory = ${env:ProgramFiles}
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.FileName
    }
    return $null
}

function Select-FolderDialog {
    param([string]$Description)

    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $Description
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.SelectedPath
    }
    return $null
}

# ---------------------------------------------------------------------------
# Self-Repair Logic (Safe Bounded Directory Traversal)
# ---------------------------------------------------------------------------
function Repair-BrokenShortcutSource {
    param(
        [string]$TargetPath,
        $Shortcut,
        [string]$AppDisplayName
    )

    Write-Warning "Source target missing for '$AppDisplayName': $TargetPath"
    $fileName  = Split-Path -Leaf $TargetPath
    $sourceDir = Split-Path -Parent $TargetPath
    $foundPath = $null

    if (-not [string]::IsNullOrWhiteSpace($sourceDir)) {
        $searchRoot = Split-Path -Parent $sourceDir
        
        $isRoot = $false
        if ($searchRoot) {
            $rootPath = [System.IO.Path]::GetPathRoot($searchRoot)
            if ($searchRoot.TrimEnd('\') -ieq $rootPath.TrimEnd('\')) {
                $isRoot = $true
            }
        }

        if ($searchRoot -and -not $isRoot -and (Test-Path -LiteralPath $searchRoot -PathType Container)) {
            Write-Host "  Scanning parent directory '$searchRoot' for '$fileName'..." -ForegroundColor Cyan
            $match = Get-ChildItem -LiteralPath $searchRoot -Filter $fileName -File -Recurse -Depth 3 -ErrorAction SilentlyContinue |
                     Select-Object -First 1
            if ($match) {
                $foundPath = $match.FullName
                Write-Host "  Found target at: $foundPath" -ForegroundColor Green
            }
        }
    }

    if (-not $foundPath) {
        Write-Warning "  Automated recovery could not locate '$fileName'."
        $foundPath = Select-ExecutableManually -AppDisplayName $AppDisplayName
    }

    if ($foundPath -and (Test-Path -LiteralPath $foundPath)) {
        Update-Shortcut -Shortcut $Shortcut -ExePath $foundPath
        return $foundPath
    }

    Write-Warning "  No working executable could be linked for '$AppDisplayName'."
    return $null
}

# ---------------------------------------------------------------------------
# UWP Resolution (Via Native AppX Catalog)
# ---------------------------------------------------------------------------
function Resolve-UwpExe {
    param([string]$AppName)

    $normalized = ($AppName -replace '[^\w]', '')
    Write-Host "  Querying AppX catalog for '$AppName'..." -ForegroundColor DarkGray

    $packages = Get-AppxPackage -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match $normalized -or $_.PackageFamilyName -match $normalized
    }

    foreach ($pkg in $packages) {
        try {
            $manifest = Get-AppxPackageManifest -Package $pkg -ErrorAction Stop
            $app = $manifest.Package.Applications.Application | Select-Object -First 1
            if ($app) {
                $appId = $app.Id
                $aumid = "$($pkg.PackageFamilyName)!$appId"
                $exec  = $app.Executable
                $fullExePath = if ($exec) { Join-Path $pkg.InstallLocation $exec } else { '' }

                return @{
                    ExePath     = $fullExePath
                    Aumid       = $aumid
                    ProcessName = if ($exec) { [System.IO.Path]::GetFileNameWithoutExtension($exec) } else { $normalized }
                }
            }
        } catch {
            continue
        }
    }
    return $null
}

function Invoke-UwpFork {
    param(
        $AppDisplayName, $Shortcut, $File,
        [int]$ProcessTimeout, [int]$WindowTimeout
    )

    Write-Host "  Attempting UWP resolution for '$AppDisplayName'..." -ForegroundColor Yellow
    $resolved = Resolve-UwpExe -AppName $AppDisplayName

    if ($resolved -and -not [string]::IsNullOrWhiteSpace($resolved.Aumid)) {
        Write-Host "  Resolved AUMID: $($resolved.Aumid)" -ForegroundColor Green
        Write-Host "  Launching '$AppDisplayName' via Application Activation Manager..."
        Start-Process "explorer.exe" -ArgumentList "shell:AppsFolder\$($resolved.Aumid)" -ErrorAction SilentlyContinue

        $state = Wait-ForAppReady -ProcessName $resolved.ProcessName `
            -ProcessTimeout $ProcessTimeout -WindowTimeout $WindowTimeout
        switch ($state) {
            'Window'   { Write-Host "  Window ready. $AppDisplayName is active." -ForegroundColor Green }
            'Tray'     { Write-Host "  Running in background/tray. $AppDisplayName is active." -ForegroundColor Green }
            'NotReady' { Write-Warning "'$AppDisplayName' did not report ready within timeout." }
        }
    } else {
        Write-Warning "No automated UWP match found for '$AppDisplayName'. Please locate executable manually."
        $selectedExe = Select-ExecutableManually -AppDisplayName $AppDisplayName

        if ($selectedExe -and (Test-Path -LiteralPath $selectedExe)) {
            $repairedProc = [System.IO.Path]::GetFileNameWithoutExtension($selectedExe)
            Update-Shortcut -Shortcut $Shortcut -ExePath $selectedExe

            Write-Host "  Re-launching '$AppDisplayName'..."
            $WshShell.Run('"' + $File.FullName + '"', 1, $false)
            $state = Wait-ForAppReady -ProcessName $repairedProc `
                -ProcessTimeout $ProcessTimeout -WindowTimeout $WindowTimeout
            switch ($state) {
                'Window'   { Write-Host "  Window ready. $AppDisplayName is active." -ForegroundColor Green }
                'Tray'     { Write-Host "  Running in background/tray. $AppDisplayName is active." -ForegroundColor Green }
                'NotReady' { Write-Warning "'$AppDisplayName' did not report ready after manual update." }
            }
        } else {
            Write-Host "  No executable selected for '$AppDisplayName'. Skipping." -ForegroundColor Yellow
        }
    }
}

# ---------------------------------------------------------------------------
# Option 1: Folder-First Launch Execution (Zero JSON Dependencies)
# ---------------------------------------------------------------------------
function Invoke-LaunchAllShortcuts {
    param([string]$StartMenuFolder)

    if (-not (Test-Path -LiteralPath $StartMenuFolder -PathType Container)) {
        Write-Host "Start Menu folder not found: $StartMenuFolder" -ForegroundColor Red
        return
    }

    $lnkFiles = Get-ChildItem -LiteralPath $StartMenuFolder -Filter '*.lnk' -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match $script:NumberedLnkPattern } |
        Sort-Object { if ($_.BaseName -match '^(\d{2})\s') { [int]$Matches[1] } else { 0 } }

    if ($lnkFiles.Count -eq 0) {
        Write-Host "No numbered (01-99) .lnk shortcuts found in '$StartMenuFolder'." -ForegroundColor Yellow
        return
    }

    Write-Host "`n--- Launching Startup Apps from: $StartMenuFolder ---" -ForegroundColor Cyan

    foreach ($file in $lnkFiles) {
        $appDisplayName = $file.BaseName -replace '^\d+\s*', ''
        try {
            $shortcut = $WshShell.CreateShortcut($file.FullName)
        } catch {
            Write-Warning "Skipping damaged shortcut file: $($file.FullName)"
            continue
        }

        $targetPath = $shortcut.TargetPath

        # Source Self-Repair
        $isExplorerTarget = $targetPath -ieq "$env:SystemRoot\explorer.exe"
        if (-not $isExplorerTarget -and -not [string]::IsNullOrWhiteSpace($targetPath) -and -not (Test-Path -LiteralPath $targetPath)) {
            $repairedPath = Repair-BrokenShortcutSource -TargetPath $targetPath `
                -Shortcut $shortcut -AppDisplayName $appDisplayName
            if ($repairedPath) {
                $targetPath = $repairedPath
            } else {
                Write-Warning "Skipping '$appDisplayName': source unrecoverable."
                continue
            }
        }

        # Process Identification
        $procName = Get-ProcName -TargetPath $targetPath -DisplayName $appDisplayName
        $isDisplayNameFallback = $targetPath -ieq "$env:SystemRoot\explorer.exe"

        # Skip if already running with active window
        $activeProcs = Get-Process -Name $procName -ErrorAction SilentlyContinue
        if ($activeProcs) {
            $readyState = Get-AppReadyState -ProcessName $procName
            if ($readyState -eq 'Window') {
                Write-Host "Skipping $appDisplayName (already running with active window)." -ForegroundColor DarkGray
                continue
            }
        }

        Write-Host "Launching $appDisplayName..." -ForegroundColor Cyan
        $WshShell.Run('"' + $file.FullName + '"', 1, $false)

        $state = Wait-ForAppReady -ProcessName $procName `
            -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout `
            -IsDisplayNameFallback $isDisplayNameFallback

        switch ($state) {
            'Window' {
                Write-Host "  Window ready. $appDisplayName is operational." -ForegroundColor Green
            }
            'Tray' {
                Write-Host "  Process running in background/tray." -ForegroundColor Green
            }
            'NotReady' {
                Invoke-UwpFork -AppDisplayName $appDisplayName -Shortcut $shortcut -File $file `
                    -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout
            }
        }
    }

    Write-Host "`nAll startup shortcuts processed." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Direct Folder CRUD Operations (Options 2 - 5)
# Includes Auto-Listing (Option 5 logic) after mutations
# ---------------------------------------------------------------------------
function Show-FolderShortcuts {
    param([string]$StartMenuFolder)

    $lnkFiles = Get-ChildItem -LiteralPath $StartMenuFolder -Filter '*.lnk' -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match $script:NumberedLnkPattern } |
        Sort-Object { if ($_.BaseName -match '^(\d{2})\s') { [int]$Matches[1] } else { 0 } }

    if ($lnkFiles.Count -eq 0) {
        Write-Host "No numbered shortcuts found in folder." -ForegroundColor Yellow
        return
    }

    Write-Host "`n--- Managed Shortcuts in Start Menu Folder ---" -ForegroundColor Cyan
    for ($i = 0; $i -lt $lnkFiles.Count; $i++) {
        $f = $lnkFiles[$i]
        $sc = $WshShell.CreateShortcut($f.FullName)
        Write-Host ("  [{0:D2}] {1} -> {2}" -f ($i + 1), $f.Name, $sc.TargetPath)
    }
}

function Add-StartupShortcut {
    param([string]$StartMenuFolder)

    Write-Host "`n--- Add Startup Shortcut to Folder ---" -ForegroundColor Cyan
    $displayName = Read-Host "Enter display name"
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        Write-Warning "Display name cannot be empty."
        return
    }

    $targetPath = Read-Host "Enter executable path (leave blank to browse)"
    if ([string]::IsNullOrWhiteSpace($targetPath)) {
        $targetPath = Select-ExecutableManually -AppDisplayName $displayName
    }
    if (-not $targetPath -or -not (Test-Path -LiteralPath $targetPath)) {
        Write-Warning "Target path invalid or canceled."
        return
    }

    $number  = Get-NextShortcutNumber -StartMenuFolder $StartMenuFolder
    $numStr  = '{0:D2}' -f $number
    $lnkName = "$numStr $displayName.lnk"
    $lnkPath = Join-Path $StartMenuFolder $lnkName

    try {
        $sc = $WshShell.CreateShortcut($lnkPath)
        $sc.TargetPath       = $targetPath
        $sc.WorkingDirectory = Split-Path $targetPath -Parent
        $sc.Save()
        Write-Host "Shortcut created: '$lnkName'" -ForegroundColor Green

        Show-FolderShortcuts -StartMenuFolder $StartMenuFolder
    } catch {
        Write-Warning "Failed to create shortcut: $($_.Exception.Message)"
    }
}

function Remove-StartupShortcut {
    param([string]$StartMenuFolder)

    $lnkFiles = Get-ChildItem -LiteralPath $StartMenuFolder -Filter '*.lnk' -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match $script:NumberedLnkPattern } |
        Sort-Object { if ($_.BaseName -match '^(\d{2})\s') { [int]$Matches[1] } else { 0 } }

    if ($lnkFiles.Count -eq 0) {
        Write-Host "No shortcuts in folder to remove." -ForegroundColor Yellow
        return
    }

    Write-Host "`n--- Remove Shortcut from Folder ---" -ForegroundColor Cyan
    for ($i = 0; $i -lt $lnkFiles.Count; $i++) {
        Write-Host "  [$($i + 1)] $($lnkFiles[$i].Name)"
    }

    $sel = Read-Host "Select shortcut number to delete (blank to cancel)"
    if ([string]::IsNullOrWhiteSpace($sel)) { return }
    if (-not ($sel -as [int]) -or [int]$sel -lt 1 -or [int]$sel -gt $lnkFiles.Count) {
        Write-Warning "Invalid choice."
        return
    }

    $targetFile = $lnkFiles[[int]$sel - 1]
    $confirm = Read-Host "Delete '$($targetFile.Name)' from disk? (y/N)"
    if ($confirm -match '^[Yy]') {
        Remove-Item -LiteralPath $targetFile.FullName -Force -ErrorAction SilentlyContinue
        Write-Host "Deleted '$($targetFile.Name)'." -ForegroundColor Green

        Show-FolderShortcuts -StartMenuFolder $StartMenuFolder
    }
}

function Set-StartupShortcut {
    param([string]$StartMenuFolder)

    $lnkFiles = Get-ChildItem -LiteralPath $StartMenuFolder -Filter '*.lnk' -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match $script:NumberedLnkPattern } |
        Sort-Object { if ($_.BaseName -match '^(\d{2})\s') { [int]$Matches[1] } else { 0 } }

    if ($lnkFiles.Count -eq 0) {
        Write-Host "No shortcuts in folder to modify." -ForegroundColor Yellow
        return
    }

    Write-Host "`n--- Modify Shortcut in Folder ---" -ForegroundColor Cyan
    for ($i = 0; $i -lt $lnkFiles.Count; $i++) {
        Write-Host "  [$($i + 1)] $($lnkFiles[$i].Name)"
    }

    $sel = Read-Host "Select shortcut number (blank to cancel)"
    if ([string]::IsNullOrWhiteSpace($sel)) { return }
    if (-not ($sel -as [int]) -or [int]$sel -lt 1 -or [int]$sel -gt $lnkFiles.Count) {
        Write-Warning "Invalid choice."
        return
    }

    $targetFile = $lnkFiles[[int]$sel - 1]
    Write-Host "  1. Rename display name"
    Write-Host "  2. Change executable target"
    Write-Host "  3. Change launch sequence number (01-99)"
    Write-Host "  4. Cancel"
    $action = Read-Host "Choose action"

    $modified = $false
    switch ($action) {
        '1' {
            $newName = Read-Host "Enter new display name"
            if ([string]::IsNullOrWhiteSpace($newName)) { return }
            $prefix = '01'
            if ($targetFile.BaseName -match '^(\d{2})\s') { $prefix = $Matches[1] }
            $newLnkName = "$prefix $newName.lnk"
            $newPath    = Join-Path $StartMenuFolder $newLnkName
            try {
                Rename-Item -LiteralPath $targetFile.FullName -NewName $newLnkName -ErrorAction Stop
                Write-Host "Renamed to '$newLnkName'." -ForegroundColor Green
                $modified = $true
            } catch {
                Write-Warning "Rename error: $($_.Exception.Message)"
            }
        }
        '2' {
            $newTarget = Read-Host "Enter executable path (blank to browse)"
            if ([string]::IsNullOrWhiteSpace($newTarget)) {
                $newTarget = Select-ExecutableManually -AppDisplayName $targetFile.BaseName
            }
            if (-not $newTarget -or -not (Test-Path -LiteralPath $newTarget)) {
                Write-Warning "Invalid path."
                return
            }
            try {
                $sc = $WshShell.CreateShortcut($targetFile.FullName)
                Update-Shortcut -Shortcut $sc -ExePath $newTarget
                Write-Host "Target updated successfully." -ForegroundColor Green
                $modified = $true
            } catch {
                Write-Warning "Update error: $($_.Exception.Message)"
            }
        }
        '3' {
            $newNum = Read-Host "Enter new order number (01-99)"
            if (-not ($newNum -as [int]) -or [int]$newNum -lt 1 -or [int]$newNum -gt 99) {
                Write-Warning "Must be between 1 and 99."
                return
            }
            $numStr     = '{0:D2}' -f [int]$newNum
            $baseName   = $targetFile.BaseName -replace '^\d{2}\s*', ''
            $newLnkName = "$numStr $baseName.lnk"
            try {
                Rename-Item -LiteralPath $targetFile.FullName -NewName $newLnkName -ErrorAction Stop
                Write-Host "Order updated to $numStr." -ForegroundColor Green
                $modified = $true
            } catch {
                Write-Warning "Reorder error: $($_.Exception.Message)"
            }
        }
        default { Write-Host "Canceled." }
    }

    if ($modified) {
        Show-FolderShortcuts -StartMenuFolder $StartMenuFolder
}

# ---------------------------------------------------------------------------
# Menu Option 6: Folder -> JSON (Snapshot / Export to JSON)
# ---------------------------------------------------------------------------
function Sync-FolderToJson {
    param([string]$ConfigPath, [string]$StartMenuFolder)

    if (-not (Test-Path -LiteralPath $StartMenuFolder -PathType Container)) {
        Write-Host "Start Menu folder not found: $StartMenuFolder" -ForegroundColor Red
        return
    }

    Write-Host "`n--- Sync: Folder -> JSON (Export Folder to Config) ---" -ForegroundColor Cyan
    Write-Host "Scanning: $StartMenuFolder`n" -ForegroundColor DarkGray

    $lnkFiles = Get-ChildItem -LiteralPath $StartMenuFolder -Filter '*.lnk' -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match $script:NumberedLnkPattern } |
        Sort-Object { if ($_.BaseName -match '^(\d{2})\s') { [int]$Matches[1] } else { 0 } }

    $config = Load-ConfigSafe -Path $ConfigPath
    $shortcuts = @($config.Shortcuts)
    $foundNames = @()
    $addedCount = 0
    $updatedCount = 0

    foreach ($file in $lnkFiles) {
        $displayName = $file.BaseName -replace '^\d{2}\s*', ''
        if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $file.BaseName }
        $foundNames += $displayName

        $existing = $shortcuts | Where-Object { $_.Name -eq $displayName } | Select-Object -First 1

        if ($existing) {
            if ($existing.ShortcutPath -ne $file.FullName) {
                $existing.ShortcutPath = $file.FullName
                $updatedCount++
                Write-Host "  Updated path for '$displayName' in JSON" -ForegroundColor Green
            }
        } else {
            try {
                $sc         = $WshShell.CreateShortcut($file.FullName)
                $targetPath = $sc.TargetPath
            } catch {
                Write-Warning "  Could not inspect '$($file.Name)'. Skipping."
                continue
            }
            $procName = Get-ProcName -TargetPath $targetPath -DisplayName $displayName
            $isExplorer = $targetPath -ieq "$env:SystemRoot\explorer.exe"

            $shortcuts += [PSCustomObject]@{
                Name         = $displayName
                ShortcutPath = $file.FullName
                ProcessName  = $procName
                LaunchType   = if ($isExplorer) { 'UWP' } else { 'Win32' }
                ExePath      = $targetPath
                Aumid        = ''
            }
            $addedCount++
            Write-Host "  Exported to JSON: '$displayName' ($($file.Name))" -ForegroundColor Green
        }
    }

    $missingFromFolder = @($shortcuts | Where-Object { $foundNames -notcontains $_.Name })
    $prunedCount = 0
    if ($missingFromFolder.Count -gt 0) {
        Write-Host "`nShortcuts registered in JSON but missing from disk folder:" -ForegroundColor Yellow
        foreach ($m in $missingFromFolder) {
            Write-Host "  - $($m.Name) ($($m.ShortcutPath))"
        }
        $remove = Read-Host "`nPrune (remove) these missing entries from the JSON config? (y/N)"
        if ($remove -match '^[Yy]') {
            $shortcuts = @($shortcuts | Where-Object { $foundNames -contains $_.Name })
            $prunedCount = $missingFromFolder.Count
            Write-Host "  Pruned $prunedCount entry/entries from JSON." -ForegroundColor Yellow
        } else {
            Write-Host "  Retained entries in JSON as-is." -ForegroundColor DarkGray
        }
    }

    $config.Shortcuts = $shortcuts
    $config.StartMenuPath = '%APPDATA%\Microsoft\Windows\Start Menu\Programs'
    Save-Config -Cfg $config -Path $ConfigPath

    Write-Host "`nExport complete: $addedCount added, $updatedCount updated, $prunedCount pruned. Saved to $ConfigPath." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Menu Option 7: JSON -> Folder (Reverse Sync: Rebuild on Disk)
# ---------------------------------------------------------------------------
function Sync-JsonToFolder {
    param([string]$ConfigPath, [string]$StartMenuFolder)

    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        Write-Host "JSON configuration not found at: $ConfigPath" -ForegroundColor Red
        return
    }

    $config = Load-ConfigSafe -Path $ConfigPath
    $shortcuts = @($config.Shortcuts)
    if ($shortcuts.Count -eq 0) {
        Write-Host "No shortcuts configured in JSON to rebuild." -ForegroundColor Yellow
        return
    }

    Write-Host "`n--- Reverse Sync: JSON -> Folder (Rebuild Shortcuts on Disk) ---" -ForegroundColor Cyan
    Write-Host "Source Config : $ConfigPath" -ForegroundColor DarkGray
    Write-Host "Target Folder : $StartMenuFolder`n" -ForegroundColor DarkGray

    $missingShortcuts = @($shortcuts | Where-Object { -not (Test-Path -LiteralPath $_.ShortcutPath) })

    if ($missingShortcuts.Count -eq 0) {
        Write-Host "All $($shortcuts.Count) shortcut(s) in JSON already exist in the folder. Nothing to restore." -ForegroundColor Green
        return
    }

    Write-Host "Found $($missingShortcuts.Count) shortcut(s) in JSON missing from the folder:" -ForegroundColor Yellow
    foreach ($m in $missingShortcuts) {
        Write-Host "  - $($m.Name) [Type: $($m.LaunchType), Target: $($m.ExePath)]"
    }

    $confirm = Read-Host "`nRebuild these missing shortcuts in '$StartMenuFolder'? (y/N)"
    if ($confirm -notmatch '^[Yy]') {
        Write-Host "Reverse sync canceled." -ForegroundColor DarkGray
        return
    }

    $rebuiltCount = 0
    foreach ($item in $missingShortcuts) {
        if ($item.LaunchType -eq 'Win32' -and -not (Test-Path -LiteralPath $item.ExePath)) {
            Write-Warning "Cannot rebuild '$($item.Name)': Target executable missing at '$($item.ExePath)'."
            continue
        }

        $destPath = $item.ShortcutPath
        if ([string]::IsNullOrWhiteSpace($destPath) -or -not ($destPath.StartsWith($StartMenuFolder, [System.StringComparison]::OrdinalIgnoreCase))) {
            $num      = Get-NextShortcutNumber -StartMenuFolder $StartMenuFolder
            $destPath = Join-Path $StartMenuFolder (('{0:D2} {1}.lnk' -f $num, $item.Name))
            $item.ShortcutPath = $destPath
        }

        try {
            $sc = $WshShell.CreateShortcut($destPath)
            if ($item.LaunchType -eq 'UWP' -and -not [string]::IsNullOrEmpty($item.Aumid)) {
                $sc.TargetPath       = "$env:SystemRoot\explorer.exe"
                $sc.Arguments        = "shell:AppsFolder\$($item.Aumid)"
                $sc.WorkingDirectory = $env:SystemRoot
            } else {
                $sc.TargetPath       = $item.ExePath
                $sc.WorkingDirectory = Split-Path $item.ExePath -Parent
            }
            $sc.Save()
            $rebuiltCount++
            Write-Host "  Created shortcut: '$($item.Name)' -> $destPath" -ForegroundColor Green
        } catch {
            Write-Warning "Failed to create shortcut '$($item.Name)': $($_.Exception.Message)"
        }
    }

    Save-Config -Cfg $config -Path $ConfigPath
    Write-Host "`nReverse sync complete: $rebuiltCount shortcut(s) restored on disk." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Menu Option 8: Dedicated Settings & Path Configuration Submenu
# ---------------------------------------------------------------------------
function Show-SettingsMenu {
    param(
        [ref]$FolderRef,
        [ref]$ConfigPathRef
    )

    $subQuit = $false
    while (-not $subQuit) {
        Write-Host "`n---------------------------------------------------" -ForegroundColor Yellow
        Write-Host " Settings & Path Configuration" -ForegroundColor White
        Write-Host "---------------------------------------------------" -ForegroundColor Yellow
        Write-Host " 1. View current active paths"
        Write-Host " 2. Configure Start Menu folder"
        Write-Host " 3. Configure JSON configuration location"
        Write-Host " 4. Reset to recommended defaults"
        Write-Host " 5. Return to Main Menu"
        Write-Host "---------------------------------------------------" -ForegroundColor Yellow

        $subChoice = Read-Host "Select setting option (1-5)"
        switch ($subChoice) {
            '1' {
                Write-Host "`n[Current Paths]" -ForegroundColor Cyan
                Write-Host "  Start Menu Folder : $($FolderRef.Value)"
                Write-Host "  Manifest JSON Path: $($ConfigPathRef.Value)"
            }
            '2' {
                Write-Host "`nSelect Start Menu Folder:"
                Write-Host "  [1] Current User: %APPDATA%\Microsoft\Windows\Start Menu\Programs (Recommended)"
                Write-Host "  [2] All Users   : %ProgramData%\Microsoft\Windows\Start Menu\Programs (Admin required)"
                Write-Host "  [3] Custom Folder (Browse / Manual Input)"
                $opt = Read-Host "Choose option (1-3)"
                switch ($opt) {
                    '1' {
                        $FolderRef.Value = $script:DefaultUserStartMenu
                        Write-Host "Start Menu folder set to User programs folder." -ForegroundColor Green
                    }
                    '2' {
                        $FolderRef.Value = $script:DefaultMachineStartMenu
                        Write-Host "Start Menu folder set to All Users programs folder." -ForegroundColor Green
                    }
                    '3' {
                        $browse = Select-FolderDialog -Description "Select Start Menu Folder for Startup Shortcuts"
                        if ($browse -and (Test-Path -LiteralPath $browse -PathType Container)) {
                            $FolderRef.Value = $browse
                            Write-Host "Start Menu folder set to: $browse" -ForegroundColor Green
                        } else {
                            $manual = Read-Host "Enter folder path manually"
                            $expanded = Expand-PathString -Path $manual
                            if (Test-Path -LiteralPath $expanded -PathType Container) {
                                $FolderRef.Value = $expanded
                                Write-Host "Start Menu folder set to: $expanded" -ForegroundColor Green
                            } else {
                                Write-Warning "Folder does not exist. Path unchanged."
                            }
                        }
                    }
                    default { Write-Host "Canceled." }
                }
            }
            '3' {
                Write-Host "`nSelect JSON Configuration Location:"
                Write-Host "  [1] Script directory (Default: .\Win11startupapps.json)"
                Write-Host "  [2] Custom Path (e.g., OneDrive / Cloud sync directory)"
                $jsonOpt = Read-Host "Choose option (1-2)"
                switch ($jsonOpt) {
                    '1' {
                        $ConfigPathRef.Value = $defaultConfigPath
                        Write-Host "JSON path set to default: $defaultConfigPath" -ForegroundColor Green
                    }
                    '2' {
                        $inputJson = Read-Host "Enter JSON file path (e.g., C:\Users\<Username>\OneDrive\Win11startupapps.json)"
                        $expandedJson = Expand-PathString -Path $inputJson
                        if ($expandedJson -notmatch '\.json$') { $expandedJson += '.json' }
                        $ConfigPathRef.Value = $expandedJson
                        Write-Host "JSON configuration path set to: $expandedJson" -ForegroundColor Green
                    }
                    default { Write-Host "Canceled." }
                }
            }
            '4' {
                $FolderRef.Value = $script:DefaultStartMenuFolder
                $ConfigPathRef.Value = $defaultConfigPath
                Write-Host "Reset paths to recommended system defaults." -ForegroundColor Green
            }
            '5' { $subQuit = $true }
            default { Write-Warning "Invalid choice." }
        }
    }
}

# ===========================================================================
# First-Run Initialization Check
# ===========================================================================
if (-not (Test-Path -LiteralPath $defaultConfigPath)) {
    Write-Host "`n===================================================\n Welcome to Windows 11 Startup Manager!\n===================================================" -ForegroundColor Cyan
    Write-Host "Initial Setup Defaults:"
    Write-Host "  - Start Menu Folder: %APPDATA%\Microsoft\Windows\Start Menu\Programs"
    Write-Host "  - Manifest JSON    : .\Win11startupapps.json`n"
    $acceptDefault = Read-Host "Use recommended defaults? (Y/n)"
    if ($acceptDefault -match '^[Nn]') {
        Show-SettingsMenu -FolderRef ([ref]$StartMenuFolder) -ConfigPathRef ([ref]$ActiveConfigPath)
    } else {
        # Initialize generic template
        $initialConfig = [PSCustomObject]@{
            StartMenuPath = '%APPDATA%\Microsoft\Windows\Start Menu\Programs'
            Shortcuts     = @()
        }
        Save-Config -Cfg $initialConfig -Path $defaultConfigPath
        Write-Host "Configured with recommended defaults." -ForegroundColor Green
    }
}

# ===========================================================================
# Master Menu Loop (Default: Direct Folder Access)
# ===========================================================================
$quit = $false
while (-not $quit) {
    if (-not (Test-Path -LiteralPath $StartMenuFolder -PathType Container)) {
        Write-Host "`nStart Menu folder missing: $StartMenuFolder" -ForegroundColor Red
        Write-Host "Launching Settings Menu to configure a valid path..."
        Show-SettingsMenu -FolderRef ([ref]$StartMenuFolder) -ConfigPathRef ([ref]$ActiveConfigPath)
    }

    Write-Host "`n===================================================" -ForegroundColor DarkCyan
    Write-Host " Windows 11 Startup Manager (Win11startup.ps1)" -ForegroundColor White
    Write-Host " Active Folder: $StartMenuFolder" -ForegroundColor DarkGray
    Write-Host " Config JSON  : $ActiveConfigPath" -ForegroundColor DarkGray
    Write-Host "===================================================" -ForegroundColor DarkCyan
    Write-Host " 1. Launch all startup shortcuts"
    Write-Host " 2. Add a shortcut"
    Write-Host " 3. Remove a shortcut"
    Write-Host " 4. Modify a shortcut"
    Write-Host " 5. List shortcuts in folder"
    Write-Host " 6. Sync: Folder -> JSON (Export/backup folder to JSON)"
    Write-Host " 7. Reverse Sync: JSON -> Folder (Restore/rebuild shortcuts from JSON)"
    Write-Host " 8. Settings & Path Configuration"
    Write-Host " 9. Quit"
    Write-Host "===================================================" -ForegroundColor DarkCyan

    $choice = Read-Host "Select option (1-9)"
    switch ($choice) {
        '1' { Invoke-LaunchAllShortcuts   -StartMenuFolder $StartMenuFolder }
        '2' { Add-StartupShortcut         -StartMenuFolder $StartMenuFolder }
        '3' { Remove-StartupShortcut      -StartMenuFolder $StartMenuFolder }
        '4' { Set-StartupShortcut         -StartMenuFolder $StartMenuFolder }
        '5' { Show-FolderShortcuts        -StartMenuFolder $StartMenuFolder }
        '6' { Sync-FolderToJson           -ConfigPath $ActiveConfigPath -StartMenuFolder $StartMenuFolder }
        '7' { Sync-JsonToFolder           -ConfigPath $ActiveConfigPath -StartMenuFolder $StartMenuFolder }
        '8' { Show-SettingsMenu           -FolderRef ([ref]$StartMenuFolder) -ConfigPathRef ([ref]$ActiveConfigPath) }
        '9' { $quit = $true }
        'q' { $quit = $true }
        default { Write-Warning "Invalid choice." }
    }

    if (-not $quit) {
        Write-Host ""
        Read-Host "Press Enter to return to menu" | Out-Null
    }
}

Write-Host "Exiting Startup Manager." -ForegroundColor DarkGray

