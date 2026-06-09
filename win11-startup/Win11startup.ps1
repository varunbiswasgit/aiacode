# Windows 11 Startup Manager (Simplified Version)
cls
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$defaultConfigPath = Join-Path $ScriptRoot 'Win11StartupConfig.json'
$WshShell = New-Object -ComObject WScript.Shell

# ---------------------------------------------------------------------------
# Save config to JSON
# ---------------------------------------------------------------------------
function Save-Config($Cfg, $Path) {
    try {
        $Cfg | ConvertTo-Json -Depth 5 | Out-File -FilePath $Path -Encoding UTF8 -Force
    } catch {
        Write-Warning "Failed to save config ${Path}: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Validate parsed JSON is a proper startup config
# ---------------------------------------------------------------------------
function Test-ValidConfig($Obj) {
    if ($null -eq $Obj) { return $false }
    if ($Obj -is [System.Array]) { return $false }
    $props = $Obj | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    return ($props -contains 'StartMenuPath')
}

# ---------------------------------------------------------------------------
# Derive a meaningful process name -- never returns 'explorer'.
# Falls back to display name (spaces stripped) if target is explorer.exe.
# ---------------------------------------------------------------------------
function Get-ProcName($TargetPath, $DisplayName) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    if ([string]::IsNullOrEmpty($base) -or $base -ieq 'explorer') {
        $base = $DisplayName -replace '\s+', ''
    }
    return $base
}

# ---------------------------------------------------------------------------
# Check genuine app readiness: process exists AND window or tray present.
# Returns: 'Window' | 'Tray' | 'NotReady'
# ---------------------------------------------------------------------------
function Get-AppReadyState($ProcessName) {
    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if (-not $procs) { return 'NotReady' }
    if ($procs | Where-Object { $_.MainWindowHandle -ne 0 }) { return 'Window' }
    if ($procs | Where-Object { $_.SessionId -gt 0 })        { return 'Tray' }
    return 'NotReady'
}

# ---------------------------------------------------------------------------
# Poll until app is ready or timeouts expire.
# If ProcessName looks like a display-name fallback (no file extension chars,
# derived from an explorer.exe target), returns 'NotReady' immediately so the
# caller falls through to UWP fork without waiting.
# Returns: 'Window' | 'Tray' | 'NotReady'
# ---------------------------------------------------------------------------
function Wait-ForAppReady($ProcessName, $ProcessTimeout, $WindowTimeout, $IsDisplayNameFallback = $false) {
    # If the proc name is unreliable (explorer target fallback), skip wait entirely
    if ($IsDisplayNameFallback) { return 'NotReady' }

    $found = $false
    for ($i = 0; $i -lt $ProcessTimeout; $i++) {
        if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) { $found = $true; break }
        Start-Sleep -Seconds 1
    }
    if (-not $found) { return 'NotReady' }
    Write-Host "  Process started. Waiting for window or tray..."
    for ($j = 0; $j -lt $WindowTimeout; $j++) {
        $state = Get-AppReadyState -ProcessName $ProcessName
        if ($state -ne 'NotReady') { return $state }
        Start-Sleep -Seconds 1
    }
    return 'NotReady'
}

# ---------------------------------------------------------------------------
# Update shortcut .lnk target to a new exe path.
# Only called explicitly on UWP resolution and file-picker repair.
# ---------------------------------------------------------------------------
function Update-Shortcut($Shortcut, $ExePath) {
    try {
        $Shortcut.TargetPath       = $ExePath
        $Shortcut.Arguments        = ''
        $Shortcut.WorkingDirectory = Split-Path $ExePath -Parent
        $Shortcut.Save()
        Write-Host "  Shortcut updated -> $ExePath"
    } catch {
        Write-Warning "Could not update shortcut: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Persist app classification entry to JSON config only.
# Does NOT touch the shortcut file.
# ---------------------------------------------------------------------------
function Update-AppEntry {
    param(
        $ExistingApp, $Config, $ConfigPath,
        $File, $AppDisplayName,
        $ProcessName, $LaunchType, $ExePath, $Aumid
    )
    if ($ExistingApp) {
        $ExistingApp.ProcessName = $ProcessName
        $ExistingApp.LaunchType  = $LaunchType
        $ExistingApp.ExePath     = $ExePath
        $ExistingApp.Aumid       = $Aumid
    } else {
        $Config.Shortcuts += [PSCustomObject]@{
            Name         = $AppDisplayName
            ShortcutPath = $File.FullName
            ProcessName  = $ProcessName
            LaunchType   = $LaunchType
            ExePath      = $ExePath
            Aumid        = $Aumid
        }
    }
    Save-Config -Cfg $Config -Path $ConfigPath
}

# ---------------------------------------------------------------------------
# Search WindowsApps for a UWP package matching $AppName.
# Returns @{ ExePath; Aumid; ProcessName } or $null.
# ---------------------------------------------------------------------------
function Resolve-UwpExe($AppName) {
    $windowsApps = 'C:\Program Files\WindowsApps'
    if (-not (Test-Path -LiteralPath $windowsApps -PathType Container)) {
        Write-Warning "WindowsApps folder not found."
        return $null
    }
    $candidates = Get-ChildItem -LiteralPath $windowsApps -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match [regex]::Escape($AppName.Replace(' ','')) -or
                       $_.Name -match ($AppName -replace '\s+', '.*') }
    foreach ($pkg in $candidates) {
        $manifest = Join-Path $pkg.FullName 'AppxManifest.xml'
        if (-not (Test-Path -LiteralPath $manifest)) { continue }
        try {
            [xml]$xml = Get-Content -LiteralPath $manifest -Raw -ErrorAction Stop
            $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace('x', 'http://schemas.microsoft.com/appx/manifest/foundation/windows10')
            $ns.AddNamespace('u', 'http://schemas.microsoft.com/appx/2010/manifest')
            $identity = $xml.SelectSingleNode('//x:Identity', $ns)
            if (-not $identity) { $identity = $xml.SelectSingleNode('//u:Identity', $ns) }
            $app = $xml.SelectSingleNode('//x:Application', $ns)
            if (-not $app) { $app = $xml.SelectSingleNode('//u:Application', $ns) }
            if (-not $identity -or -not $app) { continue }
            $exe = $app.GetAttribute('Executable')
            if ([string]::IsNullOrEmpty($exe)) { continue }
            $exePath = Join-Path $pkg.FullName $exe
            if (-not (Test-Path -LiteralPath $exePath)) { continue }
            $appId   = $app.GetAttribute('Id')
            $appxPkg = Get-AppxPackage -ErrorAction SilentlyContinue |
                       Where-Object { $_.InstallLocation -eq $pkg.FullName } | Select-Object -First 1
            $aumid   = if ($appxPkg) { $appxPkg.PackageFamilyName + '!' + $appId } else { '' }
            return @{
                ExePath     = $exePath
                Aumid       = $aumid
                ProcessName = [System.IO.Path]::GetFileNameWithoutExtension($exePath)
            }
        } catch { continue }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Shared UWP fork: resolve, update shortcut (explicit), update JSON, relaunch.
# Falls back to file-picker if no UWP match found.
# ---------------------------------------------------------------------------
function Invoke-UwpFork {
    param(
        $AppDisplayName, $Shortcut, $File,
        $ExistingApp, $Config, $ConfigPath,
        $ProcessTimeout, $WindowTimeout
    )
    Write-Host "  Attempting UWP resolution for '$AppDisplayName'..."
    $resolved = Resolve-UwpExe -AppName $AppDisplayName

    if ($resolved) {
        Write-Host "  Found UWP exe: $($resolved.ExePath)"
        Update-Shortcut -Shortcut $Shortcut -ExePath $resolved.ExePath
        Update-AppEntry -ExistingApp $ExistingApp -Config $Config -ConfigPath $ConfigPath `
            -File $File -AppDisplayName $AppDisplayName `
            -ProcessName $resolved.ProcessName -LaunchType 'UWP' `
            -ExePath $resolved.ExePath -Aumid $resolved.Aumid

        Write-Host "  Re-launching '$AppDisplayName' as UWP..."
        Start-Process -FilePath $resolved.ExePath -ErrorAction SilentlyContinue
        $state = Wait-ForAppReady -ProcessName $resolved.ProcessName `
                     -ProcessTimeout $ProcessTimeout -WindowTimeout $WindowTimeout
        switch ($state) {
            'Window'   { Write-Host "  Window ready. $AppDisplayName is up." }
            'Tray'     { Write-Host "  Running in tray. $AppDisplayName is up." }
            'NotReady' { Write-Warning "'$AppDisplayName' still did not start after UWP resolution." }
        }
    } else {
        Write-Warning "'$AppDisplayName' did not start and no UWP match found. Select executable manually."
        Add-Type -AssemblyName System.Windows.Forms
        $dlg                  = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Title            = "Select executable for $AppDisplayName"
        $dlg.Filter           = 'Executable Files (*.exe)|*.exe|All Files (*.*)|*.*'
        $dlg.InitialDirectory = $env:ProgramFiles
        $selectedExe = ''
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $selectedExe = $dlg.FileName }

        if ($selectedExe -and (Test-Path -LiteralPath $selectedExe)) {
            $repairedProc = [System.IO.Path]::GetFileNameWithoutExtension($selectedExe)
            Update-Shortcut -Shortcut $Shortcut -ExePath $selectedExe
            Update-AppEntry -ExistingApp $ExistingApp -Config $Config -ConfigPath $ConfigPath `
                -File $File -AppDisplayName $AppDisplayName `
                -ProcessName $repairedProc -LaunchType 'Win32' `
                -ExePath $selectedExe -Aumid ''

            Write-Host "  Re-launching '$AppDisplayName'..."
            $WshShell.Run('"' + $File.FullName + '"', 1, $false)
            $state = Wait-ForAppReady -ProcessName $repairedProc `
                         -ProcessTimeout $ProcessTimeout -WindowTimeout $WindowTimeout
            switch ($state) {
                'Window'   { Write-Host "  Window ready. $AppDisplayName is up." }
                'Tray'     { Write-Host "  Running in tray. $AppDisplayName is up." }
                'NotReady' { Write-Warning "'$AppDisplayName' still did not start after repair." }
            }
        } else {
            Write-Host "  No executable selected for '$AppDisplayName'. Skipping."
        }
    }
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
            Write-Warning "Config is not a valid startup config. Reinitializing."
        }
    } catch {
        Write-Warning "Config could not be loaded (invalid JSON). Reinitializing."
    }
}

if (-not $config) {
    $jsonInput = Read-Host "Enter configuration JSON file path (or press Enter for default '$defaultConfigPath')"
    if (-not [string]::IsNullOrWhiteSpace($jsonInput)) {
        if ($jsonInput -notmatch '\.json$') { $jsonInput += '.json' }
        $configPath = $jsonInput
    }
    if (-not (Test-Path -LiteralPath $configPath)) {
        $startMenuPath = Read-Host "Enter the Start Menu folder path for startup .lnk files"
        $config = [PSCustomObject]@{ StartMenuPath = $startMenuPath; Shortcuts = @() }
        Save-Config -Cfg $config -Path $configPath
        Write-Host "Configuration saved to $configPath."
    } else {
        $parsed = $null
        try { $parsed = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json } catch {
            Write-Warning "Selected config invalid JSON. Creating new config."
        }
        if (Test-ValidConfig $parsed) {
            $config = [PSCustomObject]@{
                StartMenuPath = [string]$parsed.StartMenuPath
                Shortcuts     = if ($parsed.Shortcuts) { $parsed.Shortcuts } else { @() }
            }
        } else {
            Write-Warning "Selected config is not valid. Creating new config."
            $startMenuPath = Read-Host "Enter the Start Menu folder path for startup .lnk files"
            $config = [PSCustomObject]@{ StartMenuPath = $startMenuPath; Shortcuts = @() }
            Save-Config -Cfg $config -Path $configPath
        }
        if (-not $config.StartMenuPath) {
            $config.StartMenuPath = Read-Host "Enter the Start Menu folder path for startup .lnk files"
        }
        Save-Config -Cfg $config -Path $configPath
    }
} else {
    if (-not $config.StartMenuPath) {
        $config.StartMenuPath = Read-Host "Enter the Start Menu folder path for startup .lnk files"
    }
    if (-not $config.Shortcuts) { $config.Shortcuts = @() }
    Save-Config -Cfg $config -Path $configPath
}

$startMenuFolder = $config.StartMenuPath
if (-not (Test-Path -LiteralPath $startMenuFolder -PathType Container)) {
    Write-Host "Start Menu folder not found: $startMenuFolder. Exiting." -ForegroundColor Red
    return
}

$lnkFiles = Get-ChildItem -LiteralPath $startMenuFolder -Filter '*.lnk' -ErrorAction SilentlyContinue |
    Where-Object { $_.BaseName -match '^\d{1,2}\s' } |
    Sort-Object { if ($_.BaseName -match '^(\d{1,2})\s') { [int]$Matches[1] } else { 0 } }

if ($lnkFiles.Count -eq 0) {
    Write-Host "No numbered .lnk shortcuts found in '$startMenuFolder'."
    return
}

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

    # Match config entry by Name (renumber-safe)
    $existingApp = $config.Shortcuts | Where-Object { $_.Name -eq $appDisplayName } | Select-Object -First 1

    # Keep ShortcutPath in sync if shortcut was renumbered
    if ($existingApp -and $existingApp.ShortcutPath -ne $file.FullName) {
        $existingApp.ShortcutPath = $file.FullName
        Save-Config -Cfg $config -Path $configPath
    }

    # ------------------------------------------------------------------
    # Stale entry detection: wipe classification so re-detection runs
    # ------------------------------------------------------------------
    $isStale = $existingApp -and (
        [string]::IsNullOrEmpty($existingApp.ProcessName) -or
        $existingApp.ProcessName -ieq 'explorer' -or
        ($existingApp.LaunchType -eq 'Win32' -and
            $existingApp.ExePath -ieq "$env:SystemRoot\explorer.exe") -or
        ($existingApp.LaunchType -eq 'UWP' -and (
            [string]::IsNullOrEmpty($existingApp.ExePath) -or
            -not (Test-Path -LiteralPath $existingApp.ExePath)))
    )
    if ($isStale) {
        Write-Host "Re-detecting $appDisplayName (stale or misclassified entry)..."
        $existingApp.LaunchType  = $null
        $existingApp.ProcessName = $null
        $existingApp.ExePath     = $null
        $existingApp.Aumid       = $null
    }

    # Derive real process name -- never use 'explorer'
    # Track whether it is a display-name fallback (explorer.exe target)
    $isDisplayNameFallback = $targetPath -ieq "$env:SystemRoot\explorer.exe"
    $procName = if ($existingApp -and $existingApp.ProcessName -and
                    $existingApp.ProcessName -ine 'explorer') {
                    $isDisplayNameFallback = $false   # known good name from JSON
                    $existingApp.ProcessName
                } else {
                    Get-ProcName -TargetPath $targetPath -DisplayName $appDisplayName
                }

    # ------------------------------------------------------------------
    # Skip-if-running: only skip if window OR tray confirmed
    # ------------------------------------------------------------------
    $currentState = Get-AppReadyState -ProcessName $procName
    if ($currentState -eq 'Window') {
        Write-Host "Skipping $appDisplayName (already running with active window)."
        continue
    }
    if ($currentState -eq 'Tray') {
        Write-Host "Skipping $appDisplayName (already running in tray)."
        continue
    }

    # ------------------------------------------------------------------
    # PATH 1: Known UWP with valid ExePath -> Start-Process directly
    # ------------------------------------------------------------------
    if ($existingApp -and $existingApp.LaunchType -eq 'UWP' -and
        -not [string]::IsNullOrEmpty($existingApp.ExePath) -and
        (Test-Path -LiteralPath $existingApp.ExePath)) {

        Write-Host "Launching $appDisplayName [UWP]..."
        Start-Process -FilePath $existingApp.ExePath -ErrorAction SilentlyContinue
        $state = Wait-ForAppReady -ProcessName $procName `
                     -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout
        switch ($state) {
            'Window'   { Write-Host "  Window ready. $appDisplayName is up." }
            'Tray'     { Write-Host "  Running in tray. $appDisplayName is up." }
            'NotReady' {
                Write-Warning "'$appDisplayName' [UWP] did not start. Re-resolving..."
                Invoke-UwpFork -AppDisplayName $appDisplayName -Shortcut $shortcut -File $file `
                    -ExistingApp $existingApp -Config $config -ConfigPath $configPath `
                    -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout
            }
        }
        continue
    }

    # ------------------------------------------------------------------
    # PATH 2 & 3: Known Win32 OR unknown/stale -> WshShell.Run .lnk
    # Shortcut NOT updated here -- target is already correct for Win32.
    # ------------------------------------------------------------------
    if ($existingApp -and $existingApp.LaunchType -eq 'Win32') {
        Write-Host "Launching $appDisplayName [Win32]..."
    } else {
        Write-Host "Launching $appDisplayName [detecting...]..."
    }

    $WshShell.Run('"' + $file.FullName + '"', 1, $false)

    # Wait-ForAppReady returns 'NotReady' immediately if IsDisplayNameFallback
    # (explorer.exe target) -- no inline branching needed here
    $state = Wait-ForAppReady -ProcessName $procName `
                 -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout `
                 -IsDisplayNameFallback $isDisplayNameFallback

    switch ($state) {
        { $_ -in 'Window', 'Tray' } {
            $label = if ($state -eq 'Window') { 'Window ready' } else { 'Running in tray' }
            Write-Host "  $label. $appDisplayName is up."
            # Shortcut already correct -- only update JSON
            Update-AppEntry -ExistingApp $existingApp -Config $config -ConfigPath $configPath `
                -File $file -AppDisplayName $appDisplayName `
                -ProcessName $procName -LaunchType 'Win32' -ExePath $targetPath -Aumid ''
        }
        'NotReady' {
            Invoke-UwpFork -AppDisplayName $appDisplayName -Shortcut $shortcut -File $file `
                -ExistingApp $existingApp -Config $config -ConfigPath $configPath `
                -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout
        }
    }
}
