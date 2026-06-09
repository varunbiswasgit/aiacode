# Windows 11 Startup Manager (Simplified Version)
cls
# Prepare script path for default config file
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$defaultConfigPath = Join-Path $ScriptRoot 'Win11StartupConfig.json'

# Initialize WScript.Shell for launching shortcuts
$WshShell = New-Object -ComObject WScript.Shell

# Helper function to save configuration to JSON
function Save-Config($Cfg, $Path) {
    try {
        $Cfg | ConvertTo-Json -Depth 5 | Out-File -FilePath $Path -Encoding UTF8 -Force
    } catch {
        Write-Warning "Failed to save configuration file ${Path}: $($_.Exception.Message)"
    }
}

# Helper: test whether a parsed JSON object is a valid config (has StartMenuPath property)
function Test-ValidConfig($Obj) {
    if ($null -eq $Obj) { return $false }
    if ($Obj -is [System.Array]) { return $false }
    $props = $Obj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    return ($props -contains 'StartMenuPath')
}

# Two-phase ready check: Phase 1 waits for process, Phase 2 waits for visible window.
# Returns: 'Ready' | 'ProcessOnly' | 'Failed'
function Wait-ForAppReady($ProcessName, $ProcessTimeout, $WindowTimeout) {
    # Phase 1: wait for process to appear
    $proc = $null
    for ($i = 0; $i -lt $ProcessTimeout; $i++) {
        $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($proc) { break }
        Start-Sleep -Seconds 1
    }
    if (-not $proc) { return 'Failed' }

    Write-Host "  Process started. Waiting for window..."

    # Phase 2: wait for a visible main window (MainWindowHandle != 0)
    for ($j = 0; $j -lt $WindowTimeout; $j++) {
        $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
        if ($proc) { return 'Ready' }
        Start-Sleep -Seconds 1
    }

    # Process running but no window appeared -- tray/background app
    return 'ProcessOnly'
}

# Load or create configuration
$configPath = $defaultConfigPath
$config = $null

if (Test-Path -LiteralPath $configPath -PathType Leaf) {
    try {
        $parsed = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
        if (Test-ValidConfig $parsed) {
            $config = [PSCustomObject]@{
                StartMenuPath = [string]$parsed.StartMenuPath
                Shortcuts     = if ($parsed.Shortcuts) { $parsed.Shortcuts } else { @() }
            }
        } else {
            Write-Warning "Config file exists but is not a valid startup config (may be legacy format). It will be reinitialized."
            $config = $null
        }
    } catch {
        Write-Warning "Config file exists but could not be loaded (invalid JSON). It will be reinitialized."
        $config = $null
    }
}

if (-not $config) {
    $jsonInput = Read-Host "Enter configuration JSON file path (or press Enter for default '$defaultConfigPath')"
    if (-not [string]::IsNullOrWhiteSpace($jsonInput)) {
        if ($jsonInput -notmatch '\.json$') { $jsonInput += '.json' }
        $configPath = $jsonInput
    }
    if (-not (Test-Path -LiteralPath $configPath)) {
        $startMenuPath = Read-Host "Enter the Start Menu folder path for startup .lnk files (e.g., 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs')"
        $config = [PSCustomObject]@{
            StartMenuPath = $startMenuPath
            Shortcuts     = @()
        }
        Save-Config -Cfg $config -Path $configPath
        Write-Host "Configuration saved to $configPath."
    } else {
        $parsed = $null
        try {
            $parsed = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
        } catch {
            Write-Warning "Selected config file invalid JSON. Creating new config."
        }
        if (Test-ValidConfig $parsed) {
            $config = [PSCustomObject]@{
                StartMenuPath = [string]$parsed.StartMenuPath
                Shortcuts     = if ($parsed.Shortcuts) { $parsed.Shortcuts } else { @() }
            }
        } else {
            Write-Warning "Selected config file is not a valid startup config (may be legacy format). Creating new config."
            $startMenuPath = Read-Host "Enter the Start Menu folder path for startup .lnk files"
            $config = [PSCustomObject]@{
                StartMenuPath = $startMenuPath
                Shortcuts     = @()
            }
            Save-Config -Cfg $config -Path $configPath
        }
        if (-not $config.StartMenuPath) {
            $startMenuPath = Read-Host "Enter the Start Menu folder path for startup .lnk files"
            $config.StartMenuPath = $startMenuPath
        }
        Save-Config -Cfg $config -Path $configPath
    }
} else {
    if (-not $config.StartMenuPath) {
        $startMenuPath = Read-Host "Enter the Start Menu folder path for startup .lnk files"
        $config.StartMenuPath = $startMenuPath
    }
    if (-not $config.Shortcuts) { $config.Shortcuts = @() }
    Save-Config -Cfg $config -Path $configPath
}

# Verify Start Menu folder
$startMenuFolder = $config.StartMenuPath
if (-not (Test-Path -LiteralPath $startMenuFolder -PathType Container)) {
    Write-Host "Start Menu folder not found: $startMenuFolder. Exiting." -ForegroundColor Red
    return
}

# Find all numbered .lnk files in Start Menu folder, sorted by numeric prefix
$lnkFiles = Get-ChildItem -LiteralPath $startMenuFolder -Filter '*.lnk' -ErrorAction SilentlyContinue |
    Where-Object { $_.BaseName -match '^\d{1,2}\s' } |
    Sort-Object { if ($_.BaseName -match '^(\d{1,2})\s') { [int]$Matches[1] } else { 0 } }

if ($lnkFiles.Count -eq 0) {
    Write-Host "No numbered .lnk shortcuts found in '$startMenuFolder'."
    return
}

# Timeouts (seconds)
$ProcessStartTimeout = 15   # max wait for process to appear
$WindowReadyTimeout  = 20   # max wait for visible window after process starts

# Launch each shortcut in numeric order, skipping already running apps
foreach ($file in $lnkFiles) {
    $appDisplayName = $file.BaseName -replace '^\d+\s*', ''
    try {
        $shortcut = $WshShell.CreateShortcut($file.FullName)
    } catch {
        Write-Warning "Skipping unreadable shortcut: $($file.FullName)"
        continue
    }
    $targetPath = $shortcut.TargetPath
    $targetArgs = $shortcut.Arguments

    # Skip Windows Store / UWP apps
    if ($targetPath -ieq "$env:SystemRoot\explorer.exe" -and $targetArgs -match '^shell:appsFolder\\') {
        Write-Host "Skipping (unsupported store app): $appDisplayName"
        continue
    }

    # Resolve process name from target exe
    $expectedProc = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
    if ([string]::IsNullOrEmpty($expectedProc)) {
        Write-Warning "Skipping '$appDisplayName' (invalid or missing target path)."
        continue
    }

    # Sync config entry
    $existingApp = $config.Shortcuts | Where-Object { $_.ShortcutPath -eq $file.FullName } | Select-Object -First 1
    if ($existingApp) {
        if ($existingApp.ProcessName -ne $expectedProc) {
            $existingApp.ProcessName = $expectedProc
            Save-Config -Cfg $config -Path $configPath
        }
    } else {
        $config.Shortcuts += [PSCustomObject]@{
            Name         = $appDisplayName
            ShortcutPath = $file.FullName
            ProcessName  = $expectedProc
        }
        Save-Config -Cfg $config -Path $configPath
    }

    # Skip if process already running
    $running = Get-Process -Name $expectedProc -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($running) {
        if ($running.MainWindowHandle -ne 0) {
            Write-Host "Skipping $appDisplayName (already running with active window)."
        } else {
            Write-Host "Skipping $appDisplayName (already running in background/tray)."
        }
        continue
    }

    # Launch the shortcut
    Write-Host "Launching $appDisplayName..."
    $WshShell.Run('"' + $file.FullName + '"', 1, $false)

    # Two-phase ready check
    $readyState = Wait-ForAppReady -ProcessName $expectedProc -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout

    switch ($readyState) {
        'Ready'       { Write-Host "  Window ready. $appDisplayName is up." }
        'ProcessOnly' { Write-Host "  Running (tray/background mode). $appDisplayName is up." }
        'Failed' {
            Write-Warning "'$appDisplayName' did not start within $ProcessStartTimeout seconds."

            # Offer file-picker repair
            Add-Type -AssemblyName System.Windows.Forms
            $openDlg                 = New-Object System.Windows.Forms.OpenFileDialog
            $openDlg.Title           = "Select the correct executable for $appDisplayName"
            $openDlg.Filter          = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
            $openDlg.InitialDirectory = $env:ProgramFiles
            $selectedExe = ''
            if ($openDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedExe = $openDlg.FileName
            }

            if ($selectedExe -and (Test-Path -LiteralPath $selectedExe)) {
                Write-Host "Repairing $appDisplayName shortcut to target $selectedExe"
                try {
                    $shortcut.TargetPath       = $selectedExe
                    $shortcut.Arguments        = ''
                    $shortcut.WorkingDirectory = Split-Path -Path $selectedExe -Parent
                    $shortcut.Save()
                } catch {
                    Write-Warning "Could not update shortcut $($file.Name). Skipping."
                    continue
                }
                $expectedProc = [System.IO.Path]::GetFileNameWithoutExtension($selectedExe)
                if ($existingApp) {
                    $existingApp.ProcessName = $expectedProc
                } else {
                    $config.Shortcuts += [PSCustomObject]@{
                        Name         = $appDisplayName
                        ShortcutPath = $file.FullName
                        ProcessName  = $expectedProc
                    }
                }
                Save-Config -Cfg $config -Path $configPath
                Write-Host "Re-launching $appDisplayName..."
                $WshShell.Run('"' + $file.FullName + '"', 1, $false)

                $retryState = Wait-ForAppReady -ProcessName $expectedProc -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout
                switch ($retryState) {
                    'Ready'       { Write-Host "  Window ready. $appDisplayName is up." }
                    'ProcessOnly' { Write-Host "  Running (tray/background mode). $appDisplayName is up." }
                    'Failed'      { Write-Warning "'$appDisplayName' still did not start after repair. Moving on." }
                }
            } else {
                Write-Host "No executable selected for '$appDisplayName'. Skipping."
            }
        }
    }
}
