# Windows 11 Startup Manager (Simplified Version)
cls
# Prepare script path for default config file
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$defaultConfigPath = Join-Path $ScriptRoot 'Win11StartupConfig.json'

# Initialize WScript.Shell for launching shortcuts
$WshShell = New-Object -ComObject WScript.Shell

# ---------------------------------------------------------------------------
# Helper: save configuration to JSON
# ---------------------------------------------------------------------------
function Save-Config($Cfg, $Path) {
    try {
        $Cfg | ConvertTo-Json -Depth 5 | Out-File -FilePath $Path -Encoding UTF8 -Force
    } catch {
        Write-Warning "Failed to save configuration file ${Path}: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Helper: validate a parsed JSON object is a proper startup config
# ---------------------------------------------------------------------------
function Test-ValidConfig($Obj) {
    if ($null -eq $Obj) { return $false }
    if ($Obj -is [System.Array]) { return $false }
    $props = $Obj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    return ($props -contains 'StartMenuPath')
}

# ---------------------------------------------------------------------------
# Two-phase ready check: Phase 1 = process exists, Phase 2 = window visible.
# Returns: 'Ready' | 'ProcessOnly' | 'Failed'
# ---------------------------------------------------------------------------
function Wait-ForAppReady($ProcessName, $ProcessTimeout, $WindowTimeout) {
    $proc = $null
    for ($i = 0; $i -lt $ProcessTimeout; $i++) {
        $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($proc) { break }
        Start-Sleep -Seconds 1
    }
    if (-not $proc) { return 'Failed' }

    Write-Host "  Process started. Waiting for window..."

    for ($j = 0; $j -lt $WindowTimeout; $j++) {
        $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
        if ($proc) { return 'Ready' }
        Start-Sleep -Seconds 1
    }
    return 'ProcessOnly'
}

# ---------------------------------------------------------------------------
# Helper: search WindowsApps for a package matching $AppName.
# Returns a hashtable: @{ ExePath = '...'; Aumid = '...' } or $null.
# ---------------------------------------------------------------------------
function Resolve-UwpExe($AppName) {
    $windowsApps = 'C:\Program Files\WindowsApps'
    if (-not (Test-Path -LiteralPath $windowsApps -PathType Container)) {
        Write-Warning "WindowsApps folder not found. Cannot resolve UWP exe."
        return $null
    }

    # Fuzzy-match package folders by app name (case-insensitive, partial match)
    $candidates = Get-ChildItem -LiteralPath $windowsApps -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match [regex]::Escape($AppName.Replace(' ','')) -or
                       $_.Name -match ($AppName -replace '\s+', '.*') }

    foreach ($pkg in $candidates) {
        $manifest = Join-Path $pkg.FullName 'AppxManifest.xml'
        if (-not (Test-Path -LiteralPath $manifest)) { continue }
        try {
            [xml]$xml = Get-Content -LiteralPath $manifest -Raw -ErrorAction Stop
            $ns  = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace('x', 'http://schemas.microsoft.com/appx/manifest/foundation/windows10')
            $ns.AddNamespace('u', 'http://schemas.microsoft.com/appx/2010/manifest')

            # Try foundation/windows10 namespace first, fall back to 2010
            $identity = $xml.SelectSingleNode('//x:Identity', $ns)
            if (-not $identity) { $identity = $xml.SelectSingleNode('//u:Identity', $ns) }
            $app = $xml.SelectSingleNode('//x:Application', $ns)
            if (-not $app) { $app = $xml.SelectSingleNode('//u:Application', $ns) }

            if (-not $identity -or -not $app) { continue }

            $pfn   = $identity.GetAttribute('Name') + '_' + $identity.GetAttribute('ProcessorArchitecture') + '__' +
                     (Get-AppxPackage -ErrorAction SilentlyContinue |
                      Where-Object { $_.Name -eq $identity.GetAttribute('Name') } |
                      Select-Object -First 1).PublisherId
            $appId = $app.GetAttribute('Id')
            $exe   = $app.GetAttribute('Executable')

            if ([string]::IsNullOrEmpty($exe)) { continue }
            $exePath = Join-Path $pkg.FullName $exe
            if (-not (Test-Path -LiteralPath $exePath)) { continue }

            # Build AUMID from Get-AppxPackage (most reliable source)
            $appxPkg = Get-AppxPackage -ErrorAction SilentlyContinue |
                       Where-Object { $_.InstallLocation -eq $pkg.FullName } |
                       Select-Object -First 1
            if ($appxPkg) {
                $aumid = $appxPkg.PackageFamilyName + '!' + $appId
            } else {
                $aumid = ''
            }

            return @{ ExePath = $exePath; Aumid = $aumid; ProcessName = [System.IO.Path]::GetFileNameWithoutExtension($exePath) }
        } catch {
            continue
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Config load / create
# ---------------------------------------------------------------------------
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
        try { $parsed = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json } catch {
            Write-Warning "Selected config file invalid JSON. Creating new config."
        }
        if (Test-ValidConfig $parsed) {
            $config = [PSCustomObject]@{
                StartMenuPath = [string]$parsed.StartMenuPath
                Shortcuts     = if ($parsed.Shortcuts) { $parsed.Shortcuts } else { @() }
            }
        } else {
            Write-Warning "Selected config file is not a valid startup config. Creating new config."
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

# ---------------------------------------------------------------------------
# Verify Start Menu folder
# ---------------------------------------------------------------------------
$startMenuFolder = $config.StartMenuPath
if (-not (Test-Path -LiteralPath $startMenuFolder -PathType Container)) {
    Write-Host "Start Menu folder not found: $startMenuFolder. Exiting." -ForegroundColor Red
    return
}

# Find all numbered .lnk files, sorted by numeric prefix
$lnkFiles = Get-ChildItem -LiteralPath $startMenuFolder -Filter '*.lnk' -ErrorAction SilentlyContinue |
    Where-Object { $_.BaseName -match '^\d{1,2}\s' } |
    Sort-Object { if ($_.BaseName -match '^(\d{1,2})\s') { [int]$Matches[1] } else { 0 } }

if ($lnkFiles.Count -eq 0) {
    Write-Host "No numbered .lnk shortcuts found in '$startMenuFolder'."
    return
}

# Timeouts (seconds)
$ProcessStartTimeout = 15
$WindowReadyTimeout  = 20

# ---------------------------------------------------------------------------
# Main launch loop
# ---------------------------------------------------------------------------
foreach ($file in $lnkFiles) {
    $appDisplayName = $file.BaseName -replace '^\d+\s*', ''

    try {
        $shortcut = $WshShell.CreateShortcut($file.FullName)
    } catch {
        Write-Warning "Skipping unreadable shortcut: $($file.FullName)"
        continue
    }
    $targetPath = $shortcut.TargetPath

    # Look up existing config entry by Name (not by shortcut number)
    $existingApp = $config.Shortcuts | Where-Object { $_.Name -eq $appDisplayName } | Select-Object -First 1

    # Keep ShortcutPath in sync if the shortcut was renumbered
    if ($existingApp -and $existingApp.ShortcutPath -ne $file.FullName) {
        $existingApp.ShortcutPath = $file.FullName
        Save-Config -Cfg $config -Path $configPath
    }

    # -----------------------------------------------------------------------
    # KNOWN app: LaunchType already persisted in config
    # -----------------------------------------------------------------------
    if ($existingApp -and $existingApp.LaunchType) {

        $expectedProc = $existingApp.ProcessName

        # Skip if already running
        $running = Get-Process -Name $expectedProc -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($running) {
            if ($running.MainWindowHandle -ne 0) {
                Write-Host "Skipping $appDisplayName (already running with active window)."
            } else {
                Write-Host "Skipping $appDisplayName (already running in background/tray)."
            }
            continue
        }

        Write-Host "Launching $appDisplayName [$($existingApp.LaunchType)]..."

        if ($existingApp.LaunchType -eq 'UWP') {
            # Verify stored exe path still valid; re-resolve if not
            $uwpExePath = $existingApp.ExePath
            if ([string]::IsNullOrEmpty($uwpExePath) -or -not (Test-Path -LiteralPath $uwpExePath)) {
                Write-Host "  UWP exe path stale. Re-resolving from WindowsApps..."
                $resolved = Resolve-UwpExe -AppName $appDisplayName
                if ($resolved) {
                    $existingApp.ExePath     = $resolved.ExePath
                    $existingApp.Aumid       = $resolved.Aumid
                    $existingApp.ProcessName = $resolved.ProcessName
                    $expectedProc            = $resolved.ProcessName
                    Save-Config -Cfg $config -Path $configPath
                    $uwpExePath = $resolved.ExePath
                } else {
                    Write-Warning "Could not re-resolve UWP exe for '$appDisplayName'. Skipping."
                    continue
                }
            }
            Start-Process -FilePath $uwpExePath -ErrorAction SilentlyContinue
        } else {
            # Win32
            $WshShell.Run('"' + $file.FullName + '"', 1, $false)
        }

        $readyState = Wait-ForAppReady -ProcessName $expectedProc -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout
        switch ($readyState) {
            'Ready'       { Write-Host "  Window ready. $appDisplayName is up." }
            'ProcessOnly' { Write-Host "  Running (tray/background mode). $appDisplayName is up." }
            'Failed'      { Write-Warning "'$appDisplayName' did not start. Check the shortcut or reinstall the app." }
        }
        continue
    }

    # -----------------------------------------------------------------------
    # NEW app: no LaunchType yet -- attempt Win32 first
    # -----------------------------------------------------------------------
    $expectedProc = [System.IO.Path]::GetFileNameWithoutExtension($targetPath)
    if ([string]::IsNullOrEmpty($expectedProc)) {
        Write-Warning "Skipping '$appDisplayName' (invalid or missing target path)."
        continue
    }

    # Skip if already running
    $running = Get-Process -Name $expectedProc -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($running) {
        if ($running.MainWindowHandle -ne 0) {
            Write-Host "Skipping $appDisplayName (already running with active window)."
        } else {
            Write-Host "Skipping $appDisplayName (already running in background/tray)."
        }
        continue
    }

    Write-Host "Launching $appDisplayName [detecting...]..."
    $WshShell.Run('"' + $file.FullName + '"', 1, $false)

    $readyState = Wait-ForAppReady -ProcessName $expectedProc -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout

    if ($readyState -eq 'Ready' -or $readyState -eq 'ProcessOnly') {
        # Win32 confirmed
        $label = if ($readyState -eq 'Ready') { 'Window ready' } else { 'Running (tray/background mode)' }
        Write-Host "  $label. $appDisplayName is up."
        $entry = [PSCustomObject]@{
            Name         = $appDisplayName
            ShortcutPath = $file.FullName
            ProcessName  = $expectedProc
            LaunchType   = 'Win32'
            ExePath      = $targetPath
            Aumid        = ''
        }
        if ($existingApp) {
            $existingApp.ProcessName = $expectedProc
            $existingApp.LaunchType  = 'Win32'
            $existingApp.ExePath     = $targetPath
            $existingApp.Aumid       = ''
        } else {
            $config.Shortcuts += $entry
        }
        Save-Config -Cfg $config -Path $configPath

    } else {
        # Win32 failed -- attempt UWP resolution via WindowsApps
        Write-Host "  Win32 launch failed. Searching WindowsApps for UWP match..."
        $resolved = Resolve-UwpExe -AppName $appDisplayName

        if ($resolved) {
            Write-Host "  Found UWP exe: $($resolved.ExePath). Updating shortcut and config..."

            # Update the .lnk to point directly to the real exe
            try {
                $shortcut.TargetPath       = $resolved.ExePath
                $shortcut.Arguments        = ''
                $shortcut.WorkingDirectory = Split-Path $resolved.ExePath -Parent
                $shortcut.Save()
            } catch {
                Write-Warning "Could not update shortcut for '$appDisplayName': $($_.Exception.Message)"
            }

            # Persist UWP entry
            $entry = [PSCustomObject]@{
                Name         = $appDisplayName
                ShortcutPath = $file.FullName
                ProcessName  = $resolved.ProcessName
                LaunchType   = 'UWP'
                ExePath      = $resolved.ExePath
                Aumid        = $resolved.Aumid
            }
            if ($existingApp) {
                $existingApp.ProcessName = $resolved.ProcessName
                $existingApp.LaunchType  = 'UWP'
                $existingApp.ExePath     = $resolved.ExePath
                $existingApp.Aumid       = $resolved.Aumid
            } else {
                $config.Shortcuts += $entry
            }
            Save-Config -Cfg $config -Path $configPath

            Write-Host "  Re-launching $appDisplayName as UWP..."
            Start-Process -FilePath $resolved.ExePath -ErrorAction SilentlyContinue

            $retryState = Wait-ForAppReady -ProcessName $resolved.ProcessName -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout
            switch ($retryState) {
                'Ready'       { Write-Host "  Window ready. $appDisplayName is up." }
                'ProcessOnly' { Write-Host "  Running (tray/background mode). $appDisplayName is up." }
                'Failed'      { Write-Warning "'$appDisplayName' still did not start after UWP repair. Moving on." }
            }

        } else {
            # Not a UWP app -- offer Win32 file-picker repair
            Write-Warning "'$appDisplayName' did not start and no UWP match found."
            Add-Type -AssemblyName System.Windows.Forms
            $openDlg                  = New-Object System.Windows.Forms.OpenFileDialog
            $openDlg.Title            = "Select the correct executable for $appDisplayName"
            $openDlg.Filter           = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
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
                $repairedProc = [System.IO.Path]::GetFileNameWithoutExtension($selectedExe)
                $entry = [PSCustomObject]@{
                    Name         = $appDisplayName
                    ShortcutPath = $file.FullName
                    ProcessName  = $repairedProc
                    LaunchType   = 'Win32'
                    ExePath      = $selectedExe
                    Aumid        = ''
                }
                if ($existingApp) {
                    $existingApp.ProcessName = $repairedProc
                    $existingApp.LaunchType  = 'Win32'
                    $existingApp.ExePath     = $selectedExe
                    $existingApp.Aumid       = ''
                } else {
                    $config.Shortcuts += $entry
                }
                Save-Config -Cfg $config -Path $configPath

                Write-Host "Re-launching $appDisplayName..."
                $WshShell.Run('"' + $file.FullName + '"', 1, $false)

                $retryState = Wait-ForAppReady -ProcessName $repairedProc -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout
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
