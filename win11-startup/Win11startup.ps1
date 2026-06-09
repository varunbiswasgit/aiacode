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
    # Must be a single object (not an array) with a StartMenuPath property
    if ($Obj -is [System.Array]) { return $false }
    $props = $Obj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    return ($props -contains 'StartMenuPath')
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
    # No valid config loaded -> prompt user for config file path and Start Menu path
    $jsonInput = Read-Host "Enter configuration JSON file path (or press Enter for default '$defaultConfigPath')"
    if (-not [string]::IsNullOrWhiteSpace($jsonInput)) {
        if ($jsonInput -notmatch '\.json$') { $jsonInput += '.json' }
        $configPath = $jsonInput
    }
    if (-not (Test-Path -LiteralPath $configPath)) {
        # Config file not found -> create new
        $startMenuPath = Read-Host "Enter the Start Menu folder path for startup .lnk files (e.g., 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs')"
        $config = [PSCustomObject]@{
            StartMenuPath = $startMenuPath
            Shortcuts     = @()
        }
        Save-Config -Cfg $config -Path $configPath
        Write-Host "Configuration saved to $configPath."
    } else {
        # Path exists -> try to load; reinit if not valid
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
    # Valid config loaded but may lack StartMenuPath value
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

# Set process start wait timeout (seconds)
$ProcessStartTimeout = 15

# Launch each shortcut in numeric order, skipping already running apps
foreach ($file in $lnkFiles) {
    $appDisplayName = $file.BaseName -replace '^\d+\s*', ''   # remove numeric prefix for display
    # Create WshShortcut object to get target details
    try {
        $shortcut = $WshShell.CreateShortcut($file.FullName)
    } catch {
        Write-Warning "Skipping unreadable shortcut: $($file.FullName)"
        continue
    }
    $targetPath = $shortcut.TargetPath
    $targetArgs = $shortcut.Arguments
    # Skip Windows Store apps (shell:appsFolder shortcuts)
    if ($targetPath -ieq "$env:SystemRoot\explorer.exe" -and $targetArgs -match '^shell:appsFolder\\') {
        Write-Host "Skipping (unsupported store app): $appDisplayName"
        continue
    }
    # Determine expected process name from target executable
    $expectedProc = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
    if ([string]::IsNullOrEmpty($expectedProc)) {
        Write-Warning "Skipping '$appDisplayName' (invalid or missing target path)."
        continue
    }
    # Sync config: update or add shortcut entry with process name
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
    if (Get-Process -Name $expectedProc -ErrorAction SilentlyContinue) {
        Write-Host "Skipping $appDisplayName (already running)."
        continue
    }
    # Launch the shortcut
    Write-Host "Launching $appDisplayName..."
    $WshShell.Run('"' + $file.FullName + '"', 1, $false)
    # Wait for process to start
    $processStarted = $false
    for ($i = 0; $i -lt $ProcessStartTimeout; $i++) {
        if (Get-Process -Name $expectedProc -ErrorAction SilentlyContinue) {
            $processStarted = $true; break
        }
        Start-Sleep -Seconds 1
    }
    if (-not $processStarted) {
        Write-Warning "'$appDisplayName' did not start within $ProcessStartTimeout seconds."
        # Prompt user to manually select correct executable
        Add-Type -AssemblyName System.Windows.Forms
        $openDlg = New-Object System.Windows.Forms.OpenFileDialog
        $openDlg.Title = "Select the correct executable for $appDisplayName"
        $openDlg.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
        $openDlg.InitialDirectory = $env:ProgramFiles
        $selectedExe = ''
        if ($openDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedExe = $openDlg.FileName
        }
        if ($selectedExe -and (Test-Path -LiteralPath $selectedExe)) {
            Write-Host "Repairing $appDisplayName shortcut to target $selectedExe"
            try {
                $shortcut.TargetPath = $selectedExe
                $shortcut.Arguments = ''
                $shortcut.WorkingDirectory = Split-Path -Path $selectedExe -Parent
                $shortcut.Save()
            } catch {
                Write-Warning "Could not update shortcut $($file.Name). Skipping."
                continue
            }
            # Update expected process and config for the new target
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
            # Wait again for process to start
            $processStarted = $false
            for ($j = 0; $j -lt $ProcessStartTimeout; $j++) {
                if (Get-Process -Name $expectedProc -ErrorAction SilentlyContinue) {
                    $processStarted = $true; break
                }
                Start-Sleep -Seconds 1
            }
            if (-not $processStarted) {
                Write-Warning "'$appDisplayName' still did not start after repair. Moving on."
            }
        } else {
            Write-Host "No executable selected for '$appDisplayName'. Skipping."
        }
    }
}
