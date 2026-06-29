# Windows 11 Startup Manager (Menu + Startup Launcher)
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
# Returns: 'Window' | 'Tray' | 'NotReady'
# ---------------------------------------------------------------------------
function Wait-ForAppReady($ProcessName, $ProcessTimeout, $WindowTimeout, $IsDisplayNameFallback = $false) {
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
        $ExistingApp.ShortcutPath = $File.FullName
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
# Search WindowsApps for a package matching $AppName.
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
# Shared UWP fork: resolve, update shortcut, update JSON, relaunch.
# Falls back to file-picker if no package match found.
# ---------------------------------------------------------------------------
function Invoke-UwpFork {
    param(
        $AppDisplayName, $Shortcut, $File,
        $ExistingApp, $Config, $ConfigPath,
        $ProcessTimeout, $WindowTimeout
    )
    Write-Host "  Attempting package resolution for '$AppDisplayName'..."
    $resolved = Resolve-UwpExe -AppName $AppDisplayName

    if ($resolved) {
        Write-Host "  Found packaged app executable: $($resolved.ExePath)"
        Update-Shortcut -Shortcut $Shortcut -ExePath $resolved.ExePath
        Update-AppEntry -ExistingApp $ExistingApp -Config $Config -ConfigPath $ConfigPath `
            -File $File -AppDisplayName $AppDisplayName `
            -ProcessName $resolved.ProcessName -LaunchType 'UWP' `
            -ExePath $resolved.ExePath -Aumid $resolved.Aumid

        Write-Host "  Re-launching '$AppDisplayName' as packaged app..."
        Start-Process -FilePath $resolved.ExePath -ErrorAction SilentlyContinue
        $state = Wait-ForAppReady -ProcessName $resolved.ProcessName `
                     -ProcessTimeout $ProcessTimeout -WindowTimeout $WindowTimeout
        switch ($state) {
            'Window'   { Write-Host "  Window ready. $AppDisplayName is up." }
            'Tray'     { Write-Host "  Running in tray. $AppDisplayName is up." }
            'NotReady' { Write-Warning "'$AppDisplayName' still did not start after packaged app resolution." }
        }
    } else {
        Write-Warning "'$AppDisplayName' did not start and no packaged app match was found. Select executable manually."
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
# File/path helper modules
# ---------------------------------------------------------------------------
function Read-ProgramPath($PromptText = 'Enter full program path') {
    do {
        $programPath = Read-Host $PromptText
        if ([string]::IsNullOrWhiteSpace($programPath)) {
            Write-Host 'Path cannot be empty.' -ForegroundColor Yellow
            continue
        }
        if (-not (Test-Path -LiteralPath $programPath -PathType Leaf)) {
            Write-Host 'File not found. Enter a valid executable or program path.' -ForegroundColor Yellow
            $programPath = $null
            continue
        }
        return $programPath
    } while ($true)
}

# Strict 01-99 only. Examples matched: "01 App", "99 App". Not matched: "1 App", "001 App", "100 App".
function Test-StrictShortcutSequence($BaseName) {
    return ($BaseName -match '^(0[1-9]|[1-9][0-9])\s+')
}

function Get-ShortcutDisplayName($BaseName) {
    return (($BaseName -replace '^(0[1-9]|[1-9][0-9])\s+', '').Trim())
}

function Get-ShortcutSequenceNumber($BaseName) {
    if ($BaseName -match '^(0[1-9]|[1-9][0-9])\s+') {
        return [int]$Matches[1]
    }
    return 9999
}

function Get-OrderedShortcutFiles($FolderPath) {
    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) { return @() }

    $files = Get-ChildItem -LiteralPath $FolderPath -Filter '*.lnk' -ErrorAction SilentlyContinue |
        Where-Object { Test-StrictShortcutSequence $_.BaseName }

    if (-not $files) { return @() }

    return $files | Sort-Object `
        @{ Expression = { Get-ShortcutSequenceNumber $_.BaseName }; Ascending = $true },
        @{ Expression = { Get-ShortcutDisplayName $_.BaseName }; Ascending = $true }
}

function Restore-OrRemoveTempShortcut {
    param(
        [Parameter(Mandatory = $true)] $TempFile,
        [Parameter(Mandatory = $true)] $FolderPath
    )

    $base = $TempFile.BaseName
    $restoreBase = $null

    if ($base -match '^__TMP_RESEQ__(.+)$') {
        $restoreBase = $Matches[1]
    } elseif ($base -match '^__TMP__(.+)$') {
        $restoreBase = $Matches[1]
    }

    if ([string]::IsNullOrWhiteSpace($restoreBase)) {
        Remove-Item -LiteralPath $TempFile.FullName -Force
        return
    }

    $restoreName = $restoreBase + $TempFile.Extension
    $restorePath = Join-Path $FolderPath $restoreName

    if (-not (Test-Path -LiteralPath $restorePath)) {
        Rename-Item -LiteralPath $TempFile.FullName -NewName $restoreName -Force
    } else {
        $recoveredName = ('_Recovered_{0}_{1}' -f ([guid]::NewGuid().ToString('N')), $restoreName)
        Rename-Item -LiteralPath $TempFile.FullName -NewName $recoveredName -Force
        Write-Warning "Temp shortcut restored as '$recoveredName' because '$restoreName' already exists."
    }
}

function Cleanup-OrphanTempShortcuts($FolderPath) {
    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) { return }

    Get-ChildItem -LiteralPath $FolderPath -Filter '__TMP*.lnk' -ErrorAction SilentlyContinue |
        ForEach-Object {
            try {
                Restore-OrRemoveTempShortcut -TempFile $_ -FolderPath $FolderPath
            } catch {
                Write-Warning "Failed to cleanup temp shortcut '$($_.FullName)': $($_.Exception.Message)"
            }
        }
}

function Sync-ConfigShortcutPaths($Config, $ConfigPath, $FolderPath) {
    $files = Get-OrderedShortcutFiles -FolderPath $FolderPath
    foreach ($entry in @($Config.Shortcuts)) {
        $match = $files | Where-Object { (Get-ShortcutDisplayName $_.BaseName) -eq $entry.Name } | Select-Object -First 1
        if ($match) {
            $entry.ShortcutPath = $match.FullName
        } elseif ($entry.ShortcutPath -and -not (Test-Path -LiteralPath $entry.ShortcutPath)) {
            $Config.Shortcuts = @($Config.Shortcuts | Where-Object { $_.Name -ne $entry.Name })
        }
    }
    Save-Config -Cfg $Config -Path $ConfigPath
}

function Resequence-Shortcuts {
    param(
        [Parameter(Mandatory = $true)] $FolderPath,
        [Parameter(Mandatory = $true)] $Config,
        [Parameter(Mandatory = $true)] $ConfigPath
    )

    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) { return }

    Cleanup-OrphanTempShortcuts -FolderPath $FolderPath

    $files = Get-OrderedShortcutFiles -FolderPath $FolderPath
    if (-not $files -or $files.Count -eq 0) {
        $Config.Shortcuts = @()
        Save-Config -Cfg $Config -Path $ConfigPath
        return
    }

    $tempMap = @()
    $index = 1

    try {
        foreach ($file in $files) {
            $displayName = Get-ShortcutDisplayName $file.BaseName
            $tempName = ('__TMP_RESEQ__{0:00} {1}{2}' -f $index, $displayName, $file.Extension)
            $tempPath = Join-Path $FolderPath $tempName

            Rename-Item -LiteralPath $file.FullName -NewName $tempName -Force

            $finalName = ('{0:00} {1}{2}' -f $index, $displayName, $file.Extension)
            $finalPath = Join-Path $FolderPath $finalName

            $tempMap += [PSCustomObject]@{
                DisplayName = $displayName
                TempPath    = $tempPath
                TempName    = $tempName
                FinalName   = $finalName
                FinalPath   = $finalPath
            }
            $index++
        }

        foreach ($move in $tempMap) {
            if (Test-Path -LiteralPath $move.FinalPath) {
                throw "Target already exists during resequence: $($move.FinalPath)"
            }

            Rename-Item -LiteralPath $move.TempPath -NewName $move.FinalName -Force

            $entry = $Config.Shortcuts | Where-Object { $_.Name -eq $move.DisplayName } | Select-Object -First 1
            if ($entry) {
                $entry.ShortcutPath = $move.FinalPath
            }
        }

        $existingNames = Get-OrderedShortcutFiles -FolderPath $FolderPath |
            ForEach-Object { Get-ShortcutDisplayName $_.BaseName }

        $Config.Shortcuts = @($Config.Shortcuts | Where-Object { $existingNames -contains $_.Name })
        Save-Config -Cfg $Config -Path $ConfigPath
    }
    catch {
        Write-Warning "Resequence failed: $($_.Exception.Message)"
        Write-Warning "Attempting to restore temp shortcuts."

        foreach ($move in $tempMap) {
            if (Test-Path -LiteralPath $move.TempPath) {
                try {
                    $restoreName = ($move.TempName -replace '^__TMP_RESEQ__', '')
                    $restorePath = Join-Path $FolderPath $restoreName
                    if (-not (Test-Path -LiteralPath $restorePath)) {
                        Rename-Item -LiteralPath $move.TempPath -NewName $restoreName -Force
                    } else {
                        Restore-OrRemoveTempShortcut -TempFile (Get-Item -LiteralPath $move.TempPath) -FolderPath $FolderPath
                    }
                } catch {
                    Write-Warning "Could not restore '$($move.TempPath)': $($_.Exception.Message)"
                }
            }
        }

        Cleanup-OrphanTempShortcuts -FolderPath $FolderPath
    }
}

function Show-ShortcutList($FolderPath) {
    $files = Get-OrderedShortcutFiles -FolderPath $FolderPath
    if (-not $files -or $files.Count -eq 0) {
        Write-Host 'No 01-99 numbered shortcuts found.' -ForegroundColor Yellow
        return @()
    }

    Write-Host ''
    Write-Host 'Current shortcuts:' -ForegroundColor Cyan
    $i = 1
    foreach ($file in $files) {
        Write-Host ('[{0}] {1}' -f $i, $file.Name)
        $i++
    }
    Write-Host ''
    return $files
}

function Add-ShortcutModule {
    param($FolderPath, $Config, $ConfigPath)

    Write-Host ''
    Write-Host '=== Add Shortcut ===' -ForegroundColor Cyan
    $displayName = Read-Host 'Enter shortcut display name'
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        Write-Host 'Display name cannot be empty.' -ForegroundColor Yellow
        return
    }

    $programPath = Read-ProgramPath -PromptText 'Enter full program path for the shortcut'
    $procName = [System.IO.Path]::GetFileNameWithoutExtension($programPath)

    $existingFiles = Get-OrderedShortcutFiles -FolderPath $FolderPath
    if ($existingFiles.Count -ge 99) {
        Write-Host 'Cannot add shortcut. 01-99 limit reached.' -ForegroundColor Yellow
        return
    }

    $nextNumber = if ($existingFiles.Count -gt 0) { $existingFiles.Count + 1 } else { 1 }
    $newName = ('{0:00} {1}.lnk' -f $nextNumber, $displayName)
    $newPath = Join-Path $FolderPath $newName

    if (Test-Path -LiteralPath $newPath) {
        Write-Host "Shortcut already exists: $newName" -ForegroundColor Yellow
        return
    }

    try {
        $sc = $WshShell.CreateShortcut($newPath)
        $sc.TargetPath = $programPath
        $sc.WorkingDirectory = Split-Path $programPath -Parent
        $sc.Arguments = ''
        $sc.Save()

        $Config.Shortcuts += [PSCustomObject]@{
            Name         = $displayName
            ShortcutPath = $newPath
            ProcessName  = $procName
            LaunchType   = 'Win32'
            ExePath      = $programPath
            Aumid        = ''
        }
        Save-Config -Cfg $Config -Path $ConfigPath
        Resequence-Shortcuts -FolderPath $FolderPath -Config $Config -ConfigPath $ConfigPath
        Write-Host "Added shortcut: $newName" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to add shortcut: $($_.Exception.Message)"
    }
}

function Delete-ShortcutModule {
    param($FolderPath, $Config, $ConfigPath)

    Write-Host ''
    Write-Host '=== Delete Shortcut ===' -ForegroundColor Cyan
    $files = Show-ShortcutList -FolderPath $FolderPath
    if (-not $files -or $files.Count -eq 0) { return }

    $selection = Read-Host 'Enter the number of the shortcut to delete'
    if (-not ($selection -as [int])) {
        Write-Host 'Invalid selection.' -ForegroundColor Yellow
        return
    }
    $index = [int]$selection
    if ($index -lt 1 -or $index -gt $files.Count) {
        Write-Host 'Selection out of range.' -ForegroundColor Yellow
        return
    }

    $file = $files[$index - 1]
    $displayName = Get-ShortcutDisplayName $file.BaseName

    try {
        Remove-Item -LiteralPath $file.FullName -Force
        $Config.Shortcuts = @($Config.Shortcuts | Where-Object { $_.Name -ne $displayName })
        Save-Config -Cfg $Config -Path $ConfigPath
        Resequence-Shortcuts -FolderPath $FolderPath -Config $Config -ConfigPath $ConfigPath
        Write-Host "Deleted shortcut: $($file.Name)" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to delete shortcut: $($_.Exception.Message)"
    }
}

function Modify-ShortcutModule {
    param($FolderPath, $Config, $ConfigPath)

    Write-Host ''
    Write-Host '=== Modify Shortcut ===' -ForegroundColor Cyan
    $files = Show-ShortcutList -FolderPath $FolderPath
    if (-not $files -or $files.Count -eq 0) { return }

    $selection = Read-Host 'Enter the number of the shortcut to modify'
    if (-not ($selection -as [int])) {
        Write-Host 'Invalid selection.' -ForegroundColor Yellow
        return
    }
    $index = [int]$selection
    if ($index -lt 1 -or $index -gt $files.Count) {
        Write-Host 'Selection out of range.' -ForegroundColor Yellow
        return
    }

    $file = $files[$index - 1]
    $displayName = Get-ShortcutDisplayName $file.BaseName

    # --- New display name (optional -- press Enter to keep current) ---
    $newDisplayName = Read-Host ("Enter new display name for '{0}' (or press Enter to keep)" -f $displayName)
    if ([string]::IsNullOrWhiteSpace($newDisplayName)) {
        $newDisplayName = $displayName
    }

    # --- New program path (optional -- press Enter to keep current) ---
    $shortcut = $WshShell.CreateShortcut($file.FullName)
    $currentExePath = $shortcut.TargetPath

    $newProgramPath = Read-Host ("Enter new full program path (or press Enter to keep '{0}')" -f $currentExePath)

    if ([string]::IsNullOrWhiteSpace($newProgramPath)) {
        # User skipped path change -- keep existing
        $newProgramPath = $currentExePath
        Write-Host "  Keeping existing path: $newProgramPath" -ForegroundColor DarkGray
    } elseif (-not (Test-Path -LiteralPath $newProgramPath -PathType Leaf)) {
        Write-Host 'File not found. Keeping existing path.' -ForegroundColor Yellow
        $newProgramPath = $currentExePath
    }

    $newProcName = [System.IO.Path]::GetFileNameWithoutExtension($newProgramPath)

    try {
        # Update shortcut file target if path changed
        if ($newProgramPath -ne $currentExePath) {
            Update-Shortcut -Shortcut $shortcut -ExePath $newProgramPath
        }

        # Rename shortcut file if display name changed
        $entry = $Config.Shortcuts | Where-Object { $_.Name -eq $displayName } | Select-Object -First 1
        if ($newDisplayName -ne $displayName) {
            $seqNum = Get-ShortcutSequenceNumber $file.BaseName
            $newFileName = ('{0:00} {1}{2}' -f $seqNum, $newDisplayName, $file.Extension)
            $newFilePath = Join-Path $FolderPath $newFileName
            Rename-Item -LiteralPath $file.FullName -NewName $newFileName -Force
            $file = Get-Item -LiteralPath $newFilePath

            # Update config name and remove old entry if it exists under old name
            if ($entry) {
                $entry.Name        = $newDisplayName
                $entry.ShortcutPath = $file.FullName
            } else {
                # Orphan entry -- add fresh
                $Config.Shortcuts = @($Config.Shortcuts | Where-Object { $_.Name -ne $displayName })
            }
        }

        if ($entry) {
            $entry.ProcessName  = $newProcName
            $entry.LaunchType   = 'Win32'
            $entry.ExePath      = $newProgramPath
            $entry.Aumid        = ''
            $entry.ShortcutPath = $file.FullName
        } else {
            $Config.Shortcuts += [PSCustomObject]@{
                Name         = $newDisplayName
                ShortcutPath = $file.FullName
                ProcessName  = $newProcName
                LaunchType   = 'Win32'
                ExePath      = $newProgramPath
                Aumid        = ''
            }
        }

        Save-Config -Cfg $Config -Path $ConfigPath
        Resequence-Shortcuts -FolderPath $FolderPath -Config $Config -ConfigPath $ConfigPath
        Write-Host "Modified shortcut: $($file.Name)" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to modify shortcut: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Startup launcher
# ---------------------------------------------------------------------------
function Start-StartupApps {
    param($Config, $ConfigPath)

    $startMenuFolder = $Config.StartMenuPath
    if (-not (Test-Path -LiteralPath $startMenuFolder -PathType Container)) {
        Write-Host "Start Menu folder not found: $startMenuFolder. Exiting." -ForegroundColor Red
        return
    }

    Cleanup-OrphanTempShortcuts -FolderPath $startMenuFolder

    $lnkFiles = Get-OrderedShortcutFiles -FolderPath $startMenuFolder
    if (-not $lnkFiles -or $lnkFiles.Count -eq 0) {
        Write-Host "No 01-99 numbered .lnk shortcuts found in '$startMenuFolder'."
        return
    }

    $ProcessStartTimeout = 15
    $WindowReadyTimeout  = 20

    foreach ($file in $lnkFiles) {
        $appDisplayName = Get-ShortcutDisplayName $file.BaseName
        try {
            $shortcut = $WshShell.CreateShortcut($file.FullName)
        } catch {
            Write-Warning "Skipping unreadable shortcut: $($file.FullName)"
            continue
        }
        $targetPath = $shortcut.TargetPath

        $existingApp = $Config.Shortcuts | Where-Object { $_.Name -eq $appDisplayName } | Select-Object -First 1

        if ($existingApp -and $existingApp.ShortcutPath -ne $file.FullName) {
            $existingApp.ShortcutPath = $file.FullName
            Save-Config -Cfg $Config -Path $ConfigPath
        }

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

        $isDisplayNameFallback = $targetPath -ieq "$env:SystemRoot\explorer.exe"
        $procName = if ($existingApp -and $existingApp.ProcessName -and
                        $existingApp.ProcessName -ine 'explorer') {
                        $isDisplayNameFallback = $false
                        $existingApp.ProcessName
                    } else {
                        Get-ProcName -TargetPath $targetPath -DisplayName $appDisplayName
                    }

        $currentState = Get-AppReadyState -ProcessName $procName
        if ($currentState -eq 'Window') {
            Write-Host "Skipping $appDisplayName (already running with active window)."
            continue
        }
        if ($currentState -eq 'Tray') {
            Write-Host "Skipping $appDisplayName (already running in tray)."
            continue
        }

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
                        -ExistingApp $existingApp -Config $Config -ConfigPath $ConfigPath `
                        -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout
                }
            }
            continue
        }

        if ($existingApp -and $existingApp.LaunchType -eq 'Win32') {
            Write-Host "Launching $appDisplayName [Win32]..."
        } else {
            Write-Host "Launching $appDisplayName [detecting...]..."
        }

        $WshShell.Run('"' + $file.FullName + '"', 1, $false)

        $state = Wait-ForAppReady -ProcessName $procName `
                     -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout `
                     -IsDisplayNameFallback $isDisplayNameFallback

        switch ($state) {
            { $_ -in 'Window', 'Tray' } {
                $label = if ($state -eq 'Window') { 'Window ready' } else { 'Running in tray' }
                Write-Host "  $label. $appDisplayName is up."
                Update-AppEntry -ExistingApp $existingApp -Config $Config -ConfigPath $ConfigPath `
                    -File $file -AppDisplayName $appDisplayName `
                    -ProcessName $procName -LaunchType 'Win32' -ExePath $targetPath -Aumid ''
            }
            'NotReady' {
                Invoke-UwpFork -AppDisplayName $appDisplayName -Shortcut $shortcut -File $file `
                    -ExistingApp $existingApp -Config $Config -ConfigPath $ConfigPath `
                    -ProcessTimeout $ProcessStartTimeout -WindowTimeout $WindowReadyTimeout
            }
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
                Shortcuts     = if ($parsed.Shortcuts) { @($parsed.Shortcuts) } else { @() }
            }
        } else {
            Write-Warning 'Config is not a valid startup config. Reinitializing.'
        }
    } catch {
        Write-Warning 'Config could not be loaded (invalid JSON). Reinitializing.'
    }
}

if (-not $config) {
    $jsonInput = Read-Host "Enter configuration JSON file path (or press Enter for default '$defaultConfigPath')"
    if (-not [string]::IsNullOrWhiteSpace($jsonInput)) {
        if ($jsonInput -notmatch '\.json$') { $jsonInput += '.json' }
        $configPath = $jsonInput
    }
    if (-not (Test-Path -LiteralPath $configPath)) {
        $startMenuPath = Read-Host 'Enter the Start Menu folder path for startup .lnk files'
        $config = [PSCustomObject]@{ StartMenuPath = $startMenuPath; Shortcuts = @() }
        Save-Config -Cfg $config -Path $configPath
        Write-Host "Configuration saved to $configPath."
    } else {
        $parsed = $null
        try { $parsed = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json } catch {
            Write-Warning 'Selected config invalid JSON. Creating new config.'
        }
        if (Test-ValidConfig $parsed) {
            $config = [PSCustomObject]@{
                StartMenuPath = [string]$parsed.StartMenuPath
                Shortcuts     = if ($parsed.Shortcuts) { @($parsed.Shortcuts) } else { @() }
            }
        } else {
            Write-Warning 'Selected config is not valid. Creating new config.'
            $startMenuPath = Read-Host 'Enter the Start Menu folder path for startup .lnk files'
            $config = [PSCustomObject]@{ StartMenuPath = $startMenuPath; Shortcuts = @() }
            Save-Config -Cfg $config -Path $configPath
        }
        if (-not $config.StartMenuPath) {
            $config.StartMenuPath = Read-Host 'Enter the Start Menu folder path for startup .lnk files'
        }
        Save-Config -Cfg $config -Path $configPath
    }
} else {
    if (-not $config.StartMenuPath) {
        $config.StartMenuPath = Read-Host 'Enter the Start Menu folder path for startup .lnk files'
    }
    if (-not $config.Shortcuts) { $config.Shortcuts = @() }
    Save-Config -Cfg $config -Path $configPath
}

if (-not (Test-Path -LiteralPath $config.StartMenuPath -PathType Container)) {
    Write-Host "Start Menu folder not found: $($config.StartMenuPath). Exiting." -ForegroundColor Red
    return
}

Cleanup-OrphanTempShortcuts -FolderPath $config.StartMenuPath
Sync-ConfigShortcutPaths -Config $config -ConfigPath $configPath -FolderPath $config.StartMenuPath
Resequence-Shortcuts -FolderPath $config.StartMenuPath -Config $config -ConfigPath $configPath

# ---------------------------------------------------------------------------
# Main menu
# ---------------------------------------------------------------------------
function Show-MainMenu {
    Write-Host ''
    Write-Host '=== Windows 11 Startup Manager ===' -ForegroundColor Cyan
    Write-Host '1. Run startup apps'
    Write-Host '2. Add shortcut'
    Write-Host '3. Modify shortcut'
    Write-Host '4. Delete shortcut'
    Write-Host '5. List shortcuts'
    Write-Host '6. Exit'
    Write-Host ''
}

do {
    Show-MainMenu
    $choice = Read-Host 'Select an option'
    switch ($choice) {
        '1' {
            Start-StartupApps -Config $config -ConfigPath $configPath
        }
        '2' {
            Add-ShortcutModule -FolderPath $config.StartMenuPath -Config $config -ConfigPath $configPath
        }
        '3' {
            Modify-ShortcutModule -FolderPath $config.StartMenuPath -Config $config -ConfigPath $configPath
        }
        '4' {
            Delete-ShortcutModule -FolderPath $config.StartMenuPath -Config $config -ConfigPath $configPath
        }
       '5' {
            Show-ShortcutList -FolderPath $config.StartMenuPath
        }
       '6' {
            Write-Host 'Exiting program.' -ForegroundColor Cyan
        }
        default {
            Write-Host 'Invalid option. Please select 1-6.' -ForegroundColor Yellow
        }
    }
} while ($choice -ne '6')
