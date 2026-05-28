# Win11 Startup Manager
# Launches configured apps at startup via numbered .lnk shortcuts.
# Win32 apps  : launched via WshShell.Run; self-healing shortcut repair on broken target/args.
# Appx apps   : AUMID resolved at runtime (Get-StartApps -> KnownAumid -> AppxPackage manifest).
# Bootstrap   : ensures every Win32 .lnk exists before launch; renames misnumbered or creates fresh.
# Config      : Win11startupapps.json (same folder). Add/Delete/Modify via main menu.
# Test mode   : set $env:PS_STARTUP_TESTMODE = '1' to dot-source without running the menu/sequence.
cls
# ---------------------------------------------------------------------------
# Error log + trap (defined before trap so cold errors can call Write-ErrorLog)
# ---------------------------------------------------------------------------
$script:ErrorLogPath = Join-Path $PSScriptRoot "startup-error.log"

function Write-ErrorLog {
    param([string]$Message, [System.Management.Automation.ErrorRecord]$ErrorRecord = $null)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $lines = @("[$timestamp] $Message")
    if ($ErrorRecord) {
        $lines += "  Exception : $($ErrorRecord.Exception.Message)"
        $lines += "  Category  : $($ErrorRecord.CategoryInfo)"
        $lines += "  ScriptLine: $($ErrorRecord.InvocationInfo.ScriptLineNumber) - $($ErrorRecord.InvocationInfo.Line.Trim())"
        $lines += "  StackTrace: $($ErrorRecord.ScriptStackTrace)"
    }
    $lines | Add-Content -LiteralPath $script:ErrorLogPath -Encoding UTF8
}

trap {
    Write-ErrorLog -Message "UNHANDLED TERMINATING ERROR" -ErrorRecord $_
    Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full details written to: $script:ErrorLogPath" -ForegroundColor Yellow
    break
}

# ---------------------------------------------------------------------------
# Script-scope configuration
# ---------------------------------------------------------------------------
$script:startMenu               = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
$script:InitialDelaySeconds     = 10
$script:LaunchTimeoutSeconds    = 30
$script:PostLaunchPauseSeconds  = 2
$script:SettleSeconds           = 5
$script:AppsConfigPath          = Join-Path $PSScriptRoot "Win11startupapps.json"
$script:AppsConfigSchemaVersion = 1

$script:AllowedExeRoots = @(
    $env:ProgramFiles,
    ${env:ProgramFiles(x86)},
    $env:SystemRoot
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

# ---------------------------------------------------------------------------
# Helper: detect if a shortcut entry should be Appx-typed
# Returns $true when the exe is explorer.exe and args start with shell:appsFolder
# Used in Import-AppsConfig, Sync-AppsFromStartMenu, and Initialize-Shortcut
# to guarantee consistent classification at every entry point.
# ---------------------------------------------------------------------------
function Test-IsAppxShortcut {
    param([string]$ExpectedExe, [string]$ExpectedArguments)
    return ($ExpectedExe -ieq 'explorer.exe' -and $ExpectedArguments -like 'shell:appsFolder\*')
}

# ---------------------------------------------------------------------------
# Config loader / exporter
# ---------------------------------------------------------------------------
function Import-AppsConfig {
    param([string]$Path = $script:AppsConfigPath)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Win11startupapps.json not found at '$Path'. Cannot continue without app configuration."
    }
    $raw    = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $parsed = $raw | ConvertFrom-Json

    if ($parsed -is [System.Array]) {
        Write-Warning "Win11startupapps.json has no schemaVersion wrapper. Expected schemaVersion $script:AppsConfigSchemaVersion. Processing as legacy format."
        $entries = $parsed
    } else {
        $sv = $parsed.schemaVersion
        if ($null -eq $sv) {
            Write-Warning "Win11startupapps.json schemaVersion field missing. Expected $script:AppsConfigSchemaVersion."
        } elseif ([int]$sv -ne $script:AppsConfigSchemaVersion) {
            Write-Warning "Win11startupapps.json schemaVersion is '$sv'; expected '$script:AppsConfigSchemaVersion'. Proceeding with caution."
        }
        $entries = $parsed.apps
    }

    $requiredAlways = @('Name','LaunchType','ShortcutPath','ExpectedExe')
    $validated = @()
    foreach ($entry in $entries) {
        foreach ($field in $requiredAlways) {
            if ([string]::IsNullOrWhiteSpace($entry.$field)) {
                throw "Win11startupapps.json entry missing required field '$field': $(ConvertTo-Json $entry -Compress)"
            }
        }
        if ([string]::IsNullOrWhiteSpace($entry.ProcessName)) {
            if (-not [string]::IsNullOrWhiteSpace($entry.ExpectedArguments)) {
                Write-Warning "$($entry.Name): ProcessName is empty. Will be auto-detected on first run."
            } else {
                throw "Win11startupapps.json entry missing required field 'ProcessName': $(ConvertTo-Json $entry -Compress)"
            }
        }
        if ($null -eq $entry.ExpectedPublisher) { $entry | Add-Member -NotePropertyName ExpectedPublisher -NotePropertyValue '' -Force }
        if ($null -eq $entry.ExpectedArguments) { $entry | Add-Member -NotePropertyName ExpectedArguments -NotePropertyValue '' -Force }
        if ($null -eq $entry.StartAppName)       { $entry | Add-Member -NotePropertyName StartAppName       -NotePropertyValue '' -Force }
        if ($null -eq $entry.KnownAumid)         { $entry | Add-Member -NotePropertyName KnownAumid         -NotePropertyValue '' -Force }
        if ($null -eq $entry.AppxName)           { $entry | Add-Member -NotePropertyName AppxName           -NotePropertyValue '' -Force }
        # FIX 1/3: default PresenceMode to 'Window' when absent from JSON.
        # Tray apps (OneDrive, ShareFile, Greenshot) store 'Tray' here so
        # Test-AppAlreadyOpen can skip the window-title check for them.
        if ($null -eq $entry.PresenceMode -or [string]::IsNullOrWhiteSpace($entry.PresenceMode)) {
            $entry | Add-Member -NotePropertyName PresenceMode -NotePropertyValue 'Window' -Force
        }

        # Auto-correct LaunchType: if the shortcut pattern is Appx but was saved as Win32, fix it here
        # so every downstream code path routes it correctly without requiring a re-sync.
        if ($entry.LaunchType -ne 'Appx' -and
            (Test-IsAppxShortcut -ExpectedExe $entry.ExpectedExe -ExpectedArguments $entry.ExpectedArguments)) {
            Write-Warning "$($entry.Name): LaunchType was '$($entry.LaunchType)' but shortcut uses shell:appsFolder. Auto-correcting to 'Appx'."
            $entry.LaunchType = 'Appx'
        }

        $validated += $entry
    }
    return $validated
}

function Export-AppsConfig {
    param([string]$Path = $script:AppsConfigPath)
    try {
        $wrapper = [PSCustomObject]@{ schemaVersion = $script:AppsConfigSchemaVersion; apps = $script:apps }
        $wrapper | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
        Write-Host "Win11startupapps.json saved ($($script:apps.Count) entries)."
    } catch {
        Write-ErrorLog -Message "Export-AppsConfig failed to write '$Path'" -ErrorRecord $_
        Write-Warning "Win11startupapps.json could not be saved: $($_.Exception.Message)"
        Write-Host "Error details written to: $script:ErrorLogPath" -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
# Sync from Start Menu
# FIX: $launchType is now set via Test-IsAppxShortcut inside the detection
# block, so explorer.exe + shell:appsFolder shortcuts are always written as
# LaunchType='Appx'. Previously the variable stayed 'Win32' causing
# Phone Link and Sticky Notes to be miscategorised on every sync.
# ---------------------------------------------------------------------------
function Sync-AppsFromStartMenu {
    param([string]$StartMenuPath = $script:startMenu)
    Write-Host "`n--- Sync from Start Menu ---"
    $lnkFiles = Get-ChildItem -LiteralPath $StartMenuPath -Filter '*.lnk' -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match '^\d{1,2}\s' } |
        Sort-Object Name

    if ($lnkFiles.Count -eq 0) {
        Write-Warning "No numbered .lnk files found in '$StartMenuPath'. Nothing to sync."
        return $false
    }

    Write-Host "Found $($lnkFiles.Count) numbered shortcut(s). Scanning..."

    # Load existing config so we can preserve back-filled ProcessName values.
    $existingApps = @()
    if (Test-Path -LiteralPath $script:AppsConfigPath -PathType Leaf) {
        try { $existingApps = @(Import-AppsConfig -Path $script:AppsConfigPath) } catch {}
    }

    $entries = @()

    foreach ($file in $lnkFiles) {
        $sc       = $script:WshShell.CreateShortcut($file.FullName)
        $target   = $sc.TargetPath
        $scArgs   = $sc.Arguments
        $appName  = ($file.BaseName -replace '^\d{1,2}\s+', '')
        $leafName = if ($target) { [System.IO.Path]::GetFileName($target) } else { '' }

        $launchType        = 'Win32'
        $expectedExe       = $leafName
        $processName       = [System.IO.Path]::GetFileNameWithoutExtension($leafName)
        $expectedArguments = if (-not [string]::IsNullOrWhiteSpace($scArgs)) { $scArgs.Trim() } else { '' }
        $startAppName      = ''
        $knownAumid        = ''
        $appxName          = ''

        if (Test-IsAppxShortcut -ExpectedExe $leafName -ExpectedArguments $scArgs) {
            # FIX: set LaunchType to Appx for all shell:appsFolder shortcuts.
            # Previously this block ran but never updated $launchType, leaving
            # every Appx app written as Win32 in the JSON after a sync.
            $launchType        = 'Appx'
            $expectedExe       = 'explorer.exe'
            $expectedArguments = $scArgs.Trim()
            $knownAumid        = ($scArgs -replace '^shell:appsFolder\\', '').Trim()
            $startAppName      = $appName
            $appxName          = ($knownAumid -split '_')[0]

            # Preserve ProcessName if it was already back-filled by a previous run sequence.
            $known = $existingApps | Where-Object { $_.ShortcutPath -eq $file.FullName } | Select-Object -First 1
            if ($known -and -not [string]::IsNullOrWhiteSpace($known.ProcessName)) {
                $processName = $known.ProcessName
            } else {
                $processName = ''
                Write-Warning "${appName}: ProcessName unknown after sync. Will be auto-detected on first run."
            }
        } elseif ($leafName -notlike '*.exe') {
            Write-Warning "'$($file.Name)': unexpected target '$target'. Review and fill in fields manually."
        }

        if ([string]::IsNullOrWhiteSpace($leafName)) {
            Write-Warning "'$($file.Name)': shortcut target is empty. Skipping entry."
            continue
        }

        $entries += New-AppEntry -Name $appName -LaunchType $launchType -ShortcutPath $file.FullName `
            -ProcessName $processName -ExpectedExe $expectedExe -ExpectedArguments $expectedArguments `
            -StartAppName $startAppName -KnownAumid $knownAumid -AppxName $appxName

        Write-Host ("  [{0}] {1,-28} {2}" -f $launchType, $appName, $file.Name)
    }

    $script:apps = $entries
    Export-AppsConfig
    Write-Host "--- Sync complete: $($entries.Count) entries written to Win11startupapps.json ---`n"
    return $true
}

# ---------------------------------------------------------------------------
# App entry constructor
# FIX 2/3: PresenceMode added to New-AppEntry output object.
# Without this, entries created programmatically (Add-Shortcut, Sync) were
# missing the field, so $App.PresenceMode was $null when Invoke-LaunchAttempt
# passed it to Test-AppAlreadyOpen -- defaulting to 'Window' silently only
# after a JSON round-trip via Import-AppsConfig.  Now the field is present
# from the moment the entry is constructed in memory.
# ---------------------------------------------------------------------------
function New-AppEntry {
    param(
        [string]$Name,
        [string]$LaunchType,
        [string]$ShortcutPath,
        [string]$ProcessName       = '',
        [string]$ExpectedExe       = '',
        [string]$ExpectedPublisher = '',
        [string]$ExpectedArguments = '',
        [string]$StartAppName      = '',
        [string]$KnownAumid        = '',
        [string]$AppxName          = '',
        [string]$PresenceMode      = 'Window'   # FIX 2/3: emit field so in-memory entries behave identically to JSON-loaded ones
    )
    return [PSCustomObject]@{
        Name              = $Name
        LaunchType        = $LaunchType
        ShortcutPath      = $ShortcutPath
        ProcessName       = $ProcessName
        ExpectedExe       = $ExpectedExe
        ExpectedPublisher = $ExpectedPublisher
        ExpectedArguments = $ExpectedArguments
        StartAppName      = $StartAppName
        KnownAumid        = $KnownAumid
        AppxName          = $AppxName
        PresenceMode      = $PresenceMode       # FIX 2/3: always present; 'Window' default matches Import-AppsConfig
    }
}

# ---------------------------------------------------------------------------
# Boot: resolve config path, then load
# ---------------------------------------------------------------------------
$script:WshShell = New-Object -ComObject WScript.Shell

function Resolve-ConfigPath {
    param([string]$Path = $script:AppsConfigPath)
    if (Test-Path -LiteralPath $Path -PathType Leaf) { return $true }

    Write-Host "`n[CONFIG NOT FOUND] Win11startupapps.json not found at:" -ForegroundColor Yellow
    Write-Host "  $Path" -ForegroundColor Yellow
    Write-Host "Press Enter to auto-sync from Start Menu into the default location,"
    Write-Host "or enter a full path to an existing or new .json file."
    $userPath = (Read-Host "JSON path (or Enter to auto-sync)").Trim().Trim('"')

    if ([string]::IsNullOrWhiteSpace($userPath)) {
        Write-Host "[FIRST RUN] Running sync from Start Menu..." -ForegroundColor Yellow
        if (-not (Sync-AppsFromStartMenu)) {
            Write-ErrorLog -Message "FATAL: config missing and sync found no numbered shortcuts."
            exit 1
        }
        return $true
    }

    if (-not ($userPath -imatch '\.json$')) {
        Write-Warning "Path must end in .json. Got: $userPath"
        return $false
    }

    if (Test-Path -LiteralPath $userPath -PathType Leaf) {
        $script:AppsConfigPath = $userPath
        Write-Host "Using existing config: $script:AppsConfigPath" -ForegroundColor Green
        return $true
    }

    $confirm = Read-Host "File not found. Create new empty config at '$userPath'? (Y/N)"
    if ($confirm -ine 'Y') { Write-Host "Cancelled."; return $false }

    $dir = Split-Path $userPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir -PathType Container)) {
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        catch { Write-Warning "Cannot create directory '$dir': $_"; return $false }
    }

    $script:AppsConfigPath = $userPath
    $script:apps = @()
    Export-AppsConfig
    Write-Host "Empty config created. Running sync from Start Menu to populate it..." -ForegroundColor Yellow
    if (-not (Sync-AppsFromStartMenu)) {
        Write-Warning "Sync produced no entries. Add apps manually via menu [2]."
    }
    return $true
}

$configReady = Resolve-ConfigPath
if (-not $configReady) {
    Write-ErrorLog -Message "FATAL: could not resolve a valid config path."
    exit 1
}

if (-not $script:apps) {
    try {
        $script:apps = Import-AppsConfig -Path $script:AppsConfigPath
    } catch {
        Write-ErrorLog -Message "FATAL: Import-AppsConfig failed" -ErrorRecord $_
        Write-Error $_
        Write-Host "Error details written to: $script:ErrorLogPath" -ForegroundColor Yellow
        exit 1
    }
}

# ---------------------------------------------------------------------------
# Presence / ready detection
# ---------------------------------------------------------------------------
function Get-AppPresence {
    param([string]$ProcessName, [int]$SettleSecs = $script:SettleSeconds, [switch]$Normalise)
    if ([string]::IsNullOrWhiteSpace($ProcessName)) { return $null }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $SettleSecs) {
        $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) {
            if ($Normalise) { return 'WindowVisible' } else { return 'Window' }
        }
        Start-Sleep -Seconds 1
    }
    return $null
}

# ---------------------------------------------------------------------------
# Test-AppAlreadyOpen
# DESIGN RULE: an app is only considered "already open" when it has a visible
# window (MainWindowHandle != Zero) whose title matches the app name.
# A background host process (e.g. PhoneExperienceHost) running permanently in
# Task Manager must never trigger a false "already open" skip -- it has no
# matching window title even if MainWindowHandle is briefly non-zero.
#
# FIX 2/3: Add PresenceMode parameter.
# Tray apps (OneDrive, ShareFile, Greenshot) never open a window during startup
# so the window-title check would always return $false, causing them to be
# relaunched on every startup run. When PresenceMode='Tray', return $true as
# soon as any matching process is found -- no window or title check needed.
# ---------------------------------------------------------------------------
function Test-AppAlreadyOpen {
    param(
        [string]$ProcessName,
        [string]$ExpectedExe  = '',
        [string]$AppName      = '',
        [string]$StartAppName = '',
        [string]$PresenceMode = 'Window',  # 'Window' (default) or 'Tray'
        [switch]$RequireWindow             # kept for call-site compatibility
    )
    if ([string]::IsNullOrWhiteSpace($ProcessName)) { return $false }
    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if (-not $procs) { return $false }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedExe)) {
        $procs = $procs | Where-Object {
            try   { [System.IO.Path]::GetFileName($_.MainModule.FileName) -ieq $ExpectedExe }
            catch { $false }
        }
        if (-not $procs) { return $false }
    }
    # FIX 2/3: tray apps live in the system tray and never raise a window.
    # Treat process presence alone as sufficient proof they are already running.
    if ($PresenceMode -ieq 'Tray') {
        Write-Host "$AppName: tray process already running. Skipping."
        return $true
    }
    # Window mode: require a visible window whose title matches the app name.
    # Prevents background host processes from being mistaken for the real app UI
    # (e.g. PhoneExperienceHost running permanently vs Phone Link UI window).
    $windowProcs = $procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }
    if (-not $windowProcs) { return $false }
    $terms = @($AppName, $StartAppName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if ($terms.Count -eq 0) { return $true }
    return [bool]($windowProcs | Where-Object {
        $t = $_.MainWindowTitle
        $terms | Where-Object { $t -like "*$_*" }
    })
}

function Wait-ForProcessCondition {
    param([scriptblock]$Condition, [int]$Remaining)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $Remaining) {
        if (& $Condition) { return $true }
        Start-Sleep -Seconds 1
    }
    return $false
}

function Wait-ForAppReady {
    param([string]$ProcessName, [int]$TimeoutSeconds = $script:LaunchTimeoutSeconds)
    if ([string]::IsNullOrWhiteSpace($ProcessName)) { return $true }
    $phase1Secs = [Math]::Min($script:SettleSeconds, $TimeoutSeconds)
    $mode       = Get-AppPresence -ProcessName $ProcessName -SettleSecs $phase1Secs
    $remaining  = $TimeoutSeconds - $phase1Secs
    if ($null -eq $mode) {
        return Wait-ForProcessCondition -Remaining $remaining -Condition {
            [bool](Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })
        }
    }
    Write-Host "  (presence mode: $mode)"
    return Wait-ForProcessCondition -Remaining $remaining -Condition {
        [bool](Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })
    }
}

function Wait-ForWindowByTitle {
    param(
        $App = $null,
        [string]$TitleFragment = '',
        [int]$WaitSecs = $script:SettleSeconds,
        [int]$TimeoutSeconds = -1
    )
    if ($TimeoutSeconds -ge 0) { $WaitSecs = $TimeoutSeconds }

    if (-not [string]::IsNullOrWhiteSpace($TitleFragment)) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while ($sw.Elapsed.TotalSeconds -lt $WaitSecs) {
            $match = Get-Process -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero -and $_.MainWindowTitle -like "*$TitleFragment*" } |
                Select-Object -First 1
            if ($match) { return $true }
            Start-Sleep -Seconds 1
        }
        return $false
    }

    $searchTerms = @($App.Name, $App.StartAppName) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $WaitSecs) {
        $match = Get-Process -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } |
            Where-Object {
                $title = $_.MainWindowTitle
                $searchTerms | Where-Object { $title -like "*$_*" }
            } |
            Select-Object -First 1
        if ($match) { return $match }
        Start-Sleep -Seconds 1
    }
    return $null
}

# ---------------------------------------------------------------------------
# App picker / list display
# ---------------------------------------------------------------------------
function Show-AppPicker {
    param([string]$Prompt, [switch]$AllowNew)
    $statuses = $script:apps | ForEach-Object {
        if (Test-Path -LiteralPath $_.ShortcutPath -PathType Leaf) { 'exists' } else { 'missing' }
    }
    Write-Host "`n$Prompt"
    for ($i = 0; $i -lt $script:apps.Count; $i++) {
        Write-Host ("  [{0}] {1,-20}  {2}  ({3})" -f ($i+1), $script:apps[$i].Name, $statuses[$i], $script:apps[$i].ShortcutPath)
    }
    if ($AllowNew) { Write-Host "  [N] Add a brand-new app entry" }
    Write-Host "  [0] Cancel"
    while ($true) {
        $choice = Read-Host "Select"
        if ($AllowNew -and $choice -imatch '^n$') { return '__NEW__' }
        if ($choice -match '^\d+$') {
            $idx = [int]$choice
            if ($idx -eq 0)                                   { return $null }
            if ($idx -ge 1 -and $idx -le $script:apps.Count) { return $script:apps[$idx - 1] }
        }
        $hint = if ($AllowNew) { "0-$($script:apps.Count) or N" } else { "0-$($script:apps.Count)" }
        Write-Warning "Invalid selection. Enter $hint."
    }
}

function Show-AppList {
    $sep0 = '---'
    $sep1 = '---------------------'
    $sep2 = '-----'
    $sep3 = '--------'
    $sep4 = '---------------'
    Write-Host "`n================================================"
    Write-Host "  Configured Startup Apps ($($script:apps.Count) total)"
    Write-Host "================================================"
    Write-Host ("{0,-5} {1,-22} {2,-6} {3,-10} {4}" -f '#No', 'Name', 'Type', 'Shortcut', 'Process')
    Write-Host ("{0,-5} {1,-22} {2,-6} {3,-10} {4}" -f $sep0, $sep1, $sep2, $sep3, $sep4)
    for ($i = 0; $i -lt $script:apps.Count; $i++) {
        $app      = $script:apps[$i]
        $status   = if (Test-Path -LiteralPath $app.ShortcutPath -PathType Leaf) { 'OK' } else { 'MISSING' }
        $color    = if ($status -eq 'OK') { 'Green' } else { 'Yellow' }
        # Extract the leading number from the shortcut filename (e.g. "06" from "06 Phone Link.lnk")
        $lnkBase  = [System.IO.Path]::GetFileNameWithoutExtension($app.ShortcutPath)
        $fileNum  = if ($lnkBase -match '^(\d{1,2})') { $Matches[1] } else { '-' }
        Write-Host ("{0,-5} {1,-22} {2,-6} {3,-10} {4}" -f $fileNum, $app.Name, $app.LaunchType, $status, $app.ProcessName) -ForegroundColor $color
    }
    Write-Host ""
}

# ---------------------------------------------------------------------------
# Shortcut writer
# ---------------------------------------------------------------------------
function New-AppShortcut {
    param(
        [string]$Path = '',
        [string]$TargetPath = '',
        [string]$Arguments = '',
        [string]$WorkingDirectory = '',
        [PSCustomObject]$App = $null
    )
    if ($null -ne $App) {
        $Path             = $App.ShortcutPath
        $TargetPath       = if (-not [string]::IsNullOrWhiteSpace($App.ExpectedExe)) { $App.ExpectedExe } else { "$env:SystemRoot\explorer.exe" }
        $Arguments        = if (-not [string]::IsNullOrWhiteSpace($App.ExpectedArguments)) { $App.ExpectedArguments } else { '' }
        $WorkingDirectory = if (Test-Path -LiteralPath $TargetPath -PathType Leaf -ErrorAction SilentlyContinue) { Split-Path $TargetPath -Parent } else { '' }
    }
    $sc = $script:WshShell.CreateShortcut($Path)
    $sc.TargetPath = $TargetPath
    if (-not [string]::IsNullOrWhiteSpace($Arguments))        { $sc.Arguments        = $Arguments }
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) { $sc.WorkingDirectory = $WorkingDirectory }
    $sc.Save()
}

# ---------------------------------------------------------------------------
# Exe validation
# ---------------------------------------------------------------------------
function Test-ExeAcceptable {
    param([string]$ExePath, [string]$ExpectedPublisher = "")
    $full = [System.IO.Path]::GetFullPath($ExePath)
    $allowed = $false
    foreach ($root in $script:AllowedExeRoots) {
        if ($full.StartsWith([System.IO.Path]::GetFullPath($root.TrimEnd('\')) + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
            $allowed = $true; break
        }
    }
    if (-not $allowed) { Write-Warning "Path outside allowed roots: $ExePath"; return $false }
    try {
        $sig = Get-AuthenticodeSignature -FilePath $ExePath -ErrorAction Stop
        if ($sig.Status -ne 'Valid') { Write-Warning "Signature '$($sig.Status)' for '$ExePath'."; return $false }
        if (-not [string]::IsNullOrWhiteSpace($ExpectedPublisher)) {
            if ($sig.SignerCertificate.Subject -notlike "*$ExpectedPublisher*") {
                Write-Warning "Publisher mismatch for '$ExePath'. Expected '$ExpectedPublisher'."; return $false
            }
        }
        return $true
    } catch { Write-Warning "Cannot verify signature for '$ExePath'. $_"; return $false }
}

# ---------------------------------------------------------------------------
# Shortcut health check
# ---------------------------------------------------------------------------
function Test-ShortcutHealthy {
    param([string]$ShortcutPath, [string]$ExpectedPublisher = '')
    if (-not (Test-Path -LiteralPath $ShortcutPath -PathType Leaf)) { return $false }
    try {
        $sc = $script:WshShell.CreateShortcut($ShortcutPath)
        $target = $sc.TargetPath
        if ([string]::IsNullOrWhiteSpace($target) -or -not (Test-Path -LiteralPath $target -PathType Leaf)) { return $false }
        return Test-ExeAcceptable -ExePath $target -ExpectedPublisher $ExpectedPublisher
    } catch { return $false }
}

# ---------------------------------------------------------------------------
# Shared manual-AUMID prompt helper
# Called by Edit-Shortcut and Invoke-LaunchAttempt when auto-repair fails.
# Returns the repaired argument string, or $null if the user skips.
# FIX: clears ProcessName when AUMID is corrected so stale process names
# from the old broken AUMID cannot cause a false "already open" result.
# ---------------------------------------------------------------------------
function Invoke-ManualAumidPrompt {
    param($App)
    Write-Host ""
    Write-Host "$($App.Name): automatic repair could not resolve the AUMID." -ForegroundColor Yellow
    Write-Host "You can enter the correct argument string manually." -ForegroundColor Yellow
    Write-Host "Expected format: shell:appsFolder\<PackageFamilyName>!<AppId>" -ForegroundColor Cyan
    Write-Host "Tip: run  Get-StartApps | Where-Object { `$_.Name -like '*$($App.Name)*' }  to find it." -ForegroundColor DarkGray
    Write-Host "Current stored value: $($App.ExpectedArguments)" -ForegroundColor Gray
    Write-Host "  [1] Enter argument string manually"
    Write-Host "  [2] Skip (leave unchanged)"
    $manualChoice = Read-Host "Select"
    if ($manualChoice -ne '1') { Write-Host "$($App.Name): skipped."; return $null }

    $manualArgs = (Read-Host "Enter argument string").Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($manualArgs)) { Write-Host "$($App.Name): no input provided. Skipping."; return $null }

    $sc = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
    $sc.Arguments = $manualArgs
    $sc.Save()
    $App.ExpectedArguments = $manualArgs
    $App.KnownAumid        = ($manualArgs -replace '^shell:appsFolder\\', '')
    # Clear ProcessName so it is re-detected via open window on the next run.
    # This prevents a stale ProcessName from the old AUMID causing a false
    # "already open" skip when the corrected app is launched.
    $App.ProcessName = ''
    Export-AppsConfig
    Write-Host "$($App.Name): AUMID updated, ProcessName cleared for re-detection on next run." -ForegroundColor Green
    return $manualArgs
}

# ---------------------------------------------------------------------------
# Shortcut management
# ---------------------------------------------------------------------------
function Add-Shortcut {
    param($App)
    if ($null -eq $App)     { Write-Host "Add cancelled."; return }
    if ($App -ne '__NEW__') { Initialize-Shortcut -App $App; return }

    Write-Host "`n--- Add new app entry ---"
    $name = Read-Host "App name (display label)"
    if ([string]::IsNullOrWhiteSpace($name)) { Write-Host "Cancelled."; return }

    $launchType = ''
    while ($launchType -notin @('Win32','Appx')) { $launchType = Read-Host "Launch type [Win32 / Appx]" }

    $number = ''
    while ($number -notmatch '^\d{1,2}$') {
        $number = Read-Host "Shortcut number (1-2 digits, e.g. 09)"
        if ($number -notmatch '^\d{1,2}$') { Write-Warning "Shortcut number must be 1-2 digits. Try again." }
    }
    $shortcutPath      = Join-Path $script:startMenu "$number $name.lnk"
    $processName       = Read-Host "Process name (without .exe, e.g. chrome)"
    $expectedExe       = if ($launchType -eq 'Appx') { 'explorer.exe' } else { Read-Host "Expected exe filename (e.g. chrome.exe)" }
    $expectedPublisher = Read-Host "Expected publisher CN string (optional, press Enter to skip)"
    $expectedArguments = ''
    $startAppName = ''; $knownAumid = ''; $appxName = ''

    if ($launchType -eq 'Appx') {
        Write-Host "(Appx fields - used by AUMID resolver at launch time)"
        $startAppName = Read-Host "Start menu display name fragment"
        $knownAumid   = Read-Host "Known AUMID (press Enter to skip)"
        $appxName     = Read-Host "Appx package name fragment (press Enter to skip)"
        $expectedArguments = if (-not [string]::IsNullOrWhiteSpace($knownAumid)) { "shell:appsFolder\$knownAumid" } else { '' }
    } else {
        $expectedArguments = Read-Host "Expected arguments (optional, press Enter to skip)"
    }

    # Safety net: if user entered Win32 but the args reveal it is an Appx app, correct it.
    if ($launchType -ne 'Appx' -and
        (Test-IsAppxShortcut -ExpectedExe $expectedExe -ExpectedArguments $expectedArguments)) {
        Write-Warning "$($name): arguments indicate an Appx app. Correcting LaunchType to 'Appx'."
        $launchType = 'Appx'
    }

    $newEntry = New-AppEntry -Name $name -LaunchType $launchType -ShortcutPath $shortcutPath `
        -ProcessName $processName -ExpectedExe $expectedExe -ExpectedPublisher $expectedPublisher `
        -ExpectedArguments $expectedArguments -StartAppName $startAppName `
        -KnownAumid $knownAumid -AppxName $appxName

    if ($launchType -eq 'Win32' -and [string]::IsNullOrWhiteSpace($expectedArguments)) {
        $exePath = Prompt-ForExactExePath -AppName $name -ExpectedExe $expectedExe -ExpectedPublisher $expectedPublisher
        if ($exePath) {
            $newEntry.ExpectedExe = $exePath
            Initialize-Shortcut -App $newEntry
            Write-Host "$($name): shortcut created at '$shortcutPath'."
        } else {
            Write-Warning "$($name): no exe path supplied. Entry added without shortcut creation."
        }
    } else {
        Initialize-Shortcut -App $newEntry
    }

    $script:apps += $newEntry
    Export-AppsConfig
    Write-Host "'$name' added to Win11startupapps.json."
}

function Remove-Shortcut {
    param($App)
    if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
        $confirm = Read-Host "$($App.Name): shortcut not found at '$($App.ShortcutPath)'. Remove from Win11startupapps.json anyway? (Y/N)"
        if ($confirm -ine 'Y') { Write-Host "$($App.Name): cancelled."; return }
        $script:apps = $script:apps | Where-Object { $_ -ne $App }
        Export-AppsConfig; Write-Host "$($App.Name): removed from Win11startupapps.json."
        return
    }
    $confirm = Read-Host "Delete '$($App.ShortcutPath)'? (Y/N)"
    if ($confirm -ine 'Y') { Write-Host "$($App.Name): deletion cancelled."; return }
    Remove-Item -LiteralPath $App.ShortcutPath -Force
    Write-Host "$($App.Name): shortcut deleted."
    $confirm2 = Read-Host "Also remove '$($App.Name)' from Win11startupapps.json? (Y/N)"
    if ($confirm2 -ieq 'Y') {
        $script:apps = $script:apps | Where-Object { $_ -ne $App }
        Export-AppsConfig; Write-Host "$($App.Name): removed from Win11startupapps.json."
    }
}

function Edit-Shortcut {
    param($App)
    if (-not [string]::IsNullOrWhiteSpace($App.ExpectedArguments)) {
        if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
            Write-Warning "$($App.Name): shortcut does not exist. Creating it first."
            Initialize-Shortcut -App $App; return
        }
        Write-Host "$($App.Name): re-running argument repair..."
        $repaired = Repair-ShortcutArguments -App $App
        if ($repaired) {
            Write-Host "$($App.Name): argument repair complete." -ForegroundColor Green
        } else {
            # Automatic repair failed -- offer manual AUMID entry instead of silently returning.
            $manual = Invoke-ManualAumidPrompt -App $App
            if (-not $manual) {
                Write-Warning "$($App.Name): repair not resolved. Shortcut argument unchanged."
            }
        }
        return
    }
    $exePath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe -ExpectedPublisher $App.ExpectedPublisher
    if ($exePath) {
        New-AppShortcut -Path $App.ShortcutPath -TargetPath $exePath -WorkingDirectory (Split-Path $exePath -Parent)
        $verb = if (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf) { 'updated' } else { 'created' }
        Write-Host "$($App.Name): shortcut $verb to '$exePath'."
    } else { Write-Host "$($App.Name): modify cancelled." }
}

# ---------------------------------------------------------------------------
# Core helpers
# ---------------------------------------------------------------------------
function Resolve-Aumid {
    param($App)
    $startApp = Get-StartApps | Where-Object { $_.Name -like "*$($App.StartAppName)*" } | Select-Object -First 1
    if ($startApp) { Write-Host "$($App.Name): AUMID resolved via Get-StartApps: $($startApp.AppID)"; return $startApp.AppID }
    $pkgs = @(Get-AppxPackage)
    if (-not [string]::IsNullOrWhiteSpace($App.KnownAumid)) {
        $knownPfn  = ($App.KnownAumid -split '!')[0]
        $installed = $pkgs | Where-Object { $_.PackageFamilyName -eq $knownPfn } | Select-Object -First 1
        if ($installed) { Write-Host "$($App.Name): KnownAumid verified: $($App.KnownAumid)"; return $App.KnownAumid }
        Write-Warning "$($App.Name): KnownAumid PFN '$knownPfn' not found."
    }
    $pkg = $pkgs | Where-Object { $_.Name -like "*$($App.AppxName)*" } | Select-Object -First 1
    if ($pkg) {
        try {
            $appIds = (Get-AppxPackageManifest $pkg).Package.Applications.Application.Id
            $appId  = if ($appIds -contains 'App') { 'App' } else { $appIds | Select-Object -First 1 }
            if ($appId) { $aumid = "$($pkg.PackageFamilyName)!$appId"; Write-Host "$($App.Name): AUMID via manifest: $aumid"; return $aumid }
        } catch { Write-Warning "$($App.Name): could not read AppxPackage manifest. $_" }
    }
    $msg = "$($App.Name): AUMID could not be resolved (all paths exhausted)."
    Write-Warning $msg; Write-ErrorLog -Message $msg
    return $null
}

function Get-ShortcutObject {
    param([Alias('LnkPath')][string]$ShortcutPath)
    if (-not (Test-Path -LiteralPath $ShortcutPath)) { throw "Shortcut not found: $ShortcutPath" }
    return $script:WshShell.CreateShortcut($ShortcutPath)
}

function Get-ParentFolder {
    param([string]$BrokenTargetPath)
    if ([string]::IsNullOrWhiteSpace($BrokenTargetPath)) { return $null }
    $targetFolder = Split-Path -Path $BrokenTargetPath -Parent
    $grandparent  = Split-Path -Path $targetFolder -Parent
    if (-not [string]::IsNullOrWhiteSpace($grandparent) -and (Test-Path -LiteralPath $grandparent -PathType Container)) { return $grandparent }
    $parent = $targetFolder
    if ([string]::IsNullOrWhiteSpace($parent) -or -not (Test-Path -LiteralPath $parent -PathType Container)) { return $null }
    return $parent
}

function Find-ExeWithinDepth {
    param(
        [Alias('SearchRoot')][string]$RootFolder,
        [Alias('ExeName')][string]$ExpectedExe,
        [int]$MaxDepth = [int]::MaxValue
    )
    if (-not (Test-Path -LiteralPath $RootFolder -PathType Container)) { return $null }
    $items = Get-ChildItem -LiteralPath $RootFolder -Filter $ExpectedExe -File -Recurse -ErrorAction SilentlyContinue
    if ($MaxDepth -ne [int]::MaxValue) {
        $baseDepth = ($RootFolder.TrimEnd('\') -split '\\').Count
        $items = $items | Where-Object {
            ($_.FullName -split '\\').Count - $baseDepth - 1 -le $MaxDepth
        }
    }
    return $items | Sort-Object FullName | Select-Object -First 1
}

function Prompt-ForExactExePath {
    param([string]$AppName, [string]$ExpectedExe, [string]$ExpectedPublisher = "", [int]$MaxAttempts = 3)
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $inputPath = Read-Host "Enter full path for $AppName ($ExpectedExe), or Enter to skip (attempt $attempt of $MaxAttempts)"
        if ([string]::IsNullOrWhiteSpace($inputPath)) { return $null }
        $trimmed = $inputPath.Trim('"').Trim()
        if (-not (Test-Path -LiteralPath $trimmed -PathType Leaf))                          { Write-Warning "Path does not exist: $trimmed"; continue }
        if ([System.IO.Path]::GetFileName($trimmed) -ine $ExpectedExe)                     { Write-Warning "File name must be exactly $ExpectedExe"; continue }
        if (-not (Test-ExeAcceptable -ExePath $trimmed -ExpectedPublisher $ExpectedPublisher)) { continue }
        return $trimmed
    }
    Write-Warning "${AppName}: max path attempts reached. Skipping."
    return $null
}

function Find-MisnumberedShortcut {
    param([string]$ExpectedPath, [string]$AppName)
    $expectedFile = [System.IO.Path]::GetFileName($ExpectedPath)
    $folder       = [System.IO.Path]::GetDirectoryName($ExpectedPath)
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { return $null }
    return Get-ChildItem -LiteralPath $folder -Filter "*.lnk" |
        Where-Object { $_.Name -ne $expectedFile -and $_.BaseName -like "*$AppName*" } |
        Select-Object -First 1
}

# ---------------------------------------------------------------------------
# Initialize-Shortcut
# FIX: calls Test-IsAppxShortcut to auto-correct LaunchType before any
# shortcut work, so Phone Link and other Appx apps miscategorised as Win32
# are silently fixed and saved the first time this function runs for them.
# ---------------------------------------------------------------------------
function Initialize-Shortcut {
    param($App)

    # Auto-correct miscategorised Appx entries before any shortcut work.
    if ($App.LaunchType -ne 'Appx' -and
        (Test-IsAppxShortcut -ExpectedExe $App.ExpectedExe -ExpectedArguments $App.ExpectedArguments)) {
        Write-Warning "$($App.Name): LaunchType was '$($App.LaunchType)' but uses shell:appsFolder. Auto-correcting to 'Appx' and saving."
        $App.LaunchType = 'Appx'
        Export-AppsConfig
    }

    if (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf) { return }
    $misnumbered = Find-MisnumberedShortcut -ExpectedPath $App.ShortcutPath -AppName $App.Name
    if ($misnumbered) {
        Write-Host "$($App.Name): renaming misnumbered shortcut '$($misnumbered.Name)'."
        Rename-Item -LiteralPath $misnumbered.FullName -NewName ([System.IO.Path]::GetFileName($App.ShortcutPath)); return
    }
    Write-Warning "$($App.Name): no shortcut at '$($App.ShortcutPath)'. Creating..."
    if (-not [string]::IsNullOrWhiteSpace($App.ExpectedArguments)) {
        New-AppShortcut -Path $App.ShortcutPath -TargetPath "$env:SystemRoot\explorer.exe" -Arguments $App.ExpectedArguments -WorkingDirectory $env:SystemRoot
        Write-Host "$($App.Name): shortcut created with Arguments: $($App.ExpectedArguments)"; return
    }
    $exePath = if (Test-Path -LiteralPath $App.ExpectedExe -PathType Leaf -ErrorAction SilentlyContinue) {
        $App.ExpectedExe
    } else {
        Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe -ExpectedPublisher $App.ExpectedPublisher
    }
    if ($exePath) {
        New-AppShortcut -Path $App.ShortcutPath -TargetPath $exePath -WorkingDirectory (Split-Path $exePath -Parent)
        Write-Host "$($App.Name): shortcut created at '$($App.ShortcutPath)'."
    } else {
        Write-Warning "$($App.Name): shortcut creation skipped (no valid exe path)."
    }
}

# ---------------------------------------------------------------------------
# Shortcut repair
# ---------------------------------------------------------------------------
function Invoke-ShortcutRepair {
    param($App, [scriptblock]$RepairAction)
    $shortcut = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
    $result   = & $RepairAction $shortcut
    if ($null -ne $result) { $shortcut.Save() }
    return $result
}

function Repair-ShortcutTarget {
    param($App)
    $repaired = Invoke-ShortcutRepair -App $App -RepairAction {
        param($shortcut)
        $targetPath = $shortcut.TargetPath
        Write-Warning "$($App.Name): shortcut target missing/invalid: $targetPath"
        $searchRoot = Get-ParentFolder -BrokenTargetPath $targetPath
        if ($searchRoot) {
            Write-Host "$($App.Name): searching for $($App.ExpectedExe) under $searchRoot..."
            $foundExe = Find-ExeWithinDepth -RootFolder $searchRoot -ExpectedExe $App.ExpectedExe
            if ($foundExe) {
                if (-not (Test-ExeAcceptable -ExePath $foundExe.FullName -ExpectedPublisher $App.ExpectedPublisher)) {
                    Write-Warning "$($App.Name): exe failed validation. Skipping."
                    return $null
                }
                Write-Host "$($App.Name): found replacement at $($foundExe.FullName)."
                $shortcut.TargetPath       = $foundExe.FullName
                $shortcut.WorkingDirectory = Split-Path -Path $foundExe.FullName -Parent
                if (-not [string]::IsNullOrWhiteSpace($App.ExpectedArguments)) { $shortcut.Arguments = $App.ExpectedArguments }
                return $foundExe.FullName
            }
            Write-Warning "$($App.Name): $($App.ExpectedExe) not found under $searchRoot."
        } else { Write-Warning "$($App.Name): could not determine parent folder from broken target." }
        $manualPath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe -ExpectedPublisher $App.ExpectedPublisher
        if ($manualPath) {
            $shortcut.TargetPath       = $manualPath
            $shortcut.WorkingDirectory = Split-Path -Path $manualPath -Parent
            if (-not [string]::IsNullOrWhiteSpace($App.ExpectedArguments)) { $shortcut.Arguments = $App.ExpectedArguments }
            return $manualPath
        }
        return $null
    }
}

# ---------------------------------------------------------------------------
# Shared AUMID fallback resolver
# ---------------------------------------------------------------------------
function Resolve-AumidWithFallback {
    param($App, [int]$CountdownSeconds = 30)
    $result = [PSCustomObject]@{ Resolved=$false; Aumid=$null; LaunchType=$App.LaunchType; ExePath=$null }
    if (-not [Environment]::UserInteractive) {
        Write-Warning "$($App.Name): non-interactive session. Skipping manual launch steps."
        return $result
    }
    Write-Host ""
    Write-Host "-----------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  $($App.Name): automatic AUMID resolution failed." -ForegroundColor Yellow
    Write-Host "  Step 2: Please launch '$($App.Name)' manually now" -ForegroundColor Cyan
    Write-Host "  (e.g. from Start Menu or taskbar)." -ForegroundColor Cyan
    Write-Host "-----------------------------------------------------" -ForegroundColor DarkGray
    Write-Host ""
    $beforeIds = @(Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -ExpandProperty Id)
    $detectedAumid = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $CountdownSeconds) {
        $remaining = [int]($CountdownSeconds - $sw.Elapsed.TotalSeconds)
        Write-Host ("`r  Waiting for '$($App.Name)' to appear... {0}s  " -f $remaining) -NoNewline
        $newProcs = @(Get-Process -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero -and $_.Id -notin $beforeIds })
        if ($newProcs) {
            $terms = @($App.Name,$App.StartAppName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $matched = $newProcs | Where-Object { $t=$_.MainWindowTitle; $terms | Where-Object { $t -like "*$_*" } } | Select-Object -First 1
            if (-not $matched) {
                $cands = @($newProcs | Where-Object { -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle) })
                if ($cands.Count -gt 0) {
                    Write-Host ""
                    Write-Host "  New window(s) detected. Select the one that is '$($App.Name)':" -ForegroundColor Cyan
                    for ($ci=0; $ci -lt $cands.Count; $ci++) {
                        Write-Host ("  [{0}] {1}  (process: {2})" -f ($ci+1),$cands[$ci].MainWindowTitle,$cands[$ci].ProcessName)
                    }
                    Write-Host "  [0] None -- continue waiting"
                    $pick = Read-Host "Select"
                    if ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $cands.Count) { $matched = $cands[[int]$pick-1] }
                }
            }
            if ($matched) {
                Write-Host ""
                Write-Host ("  Detected: '{0}' (process: {1})" -f $matched.MainWindowTitle,$matched.ProcessName) -ForegroundColor Green
                if ([string]::IsNullOrWhiteSpace($App.ProcessName)) { $App.ProcessName = $matched.ProcessName }
                $sa = Get-StartApps | Where-Object {
                    $_.Name -like "*$($App.Name)*" -or
                    (-not [string]::IsNullOrWhiteSpace($App.StartAppName) -and $_.Name -like "*$($App.StartAppName)*")
                } | Select-Object -First 1
                if ($sa) {
                    $detectedAumid = $sa.AppID
                    Write-Host ("  AUMID resolved via running app: {0}" -f $detectedAumid) -ForegroundColor Green
                } else {
                    $pkg = Get-AppxPackage -ErrorAction SilentlyContinue |
                        Where-Object { $_.PackageFamilyName -like "*$($matched.ProcessName)*" -or $_.Name -like "*$($App.AppxName)*" } |
                        Select-Object -First 1
                    if ($pkg) {
                        try {
                            $ids   = (Get-AppxPackageManifest $pkg).Package.Applications.Application.Id
                            $appId = if ($ids -contains 'App') { 'App' } else { $ids | Select-Object -First 1 }
                            $detectedAumid = "$($pkg.PackageFamilyName)!$appId"
                            Write-Host ("  AUMID from package: {0}" -f $detectedAumid) -ForegroundColor Green
                        } catch { Write-Warning "  Could not read manifest: $_" }
                    }
                }
                break
            }
        }
        Start-Sleep -Milliseconds 1000
    }
    Write-Host ""
    if ($detectedAumid) {
        $repairedArgs = "shell:appsFolder\$detectedAumid"
        if (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf) {
            try {
                $sc = (New-Object -ComObject WScript.Shell).CreateShortcut($App.ShortcutPath)
                $sc.Arguments = $repairedArgs; $sc.Save()
            } catch { Write-Warning "$($App.Name): could not write shortcut. $_" }
        }
        $App.ExpectedArguments = $repairedArgs
        $App | Add-Member -NotePropertyName KnownAumid -NotePropertyValue $detectedAumid -Force
        Export-AppsConfig
        Write-Host ("$($App.Name): AUMID saved as '{0}'." -f $detectedAumid) -ForegroundColor Green
        $result.Resolved=$true; $result.Aumid=$detectedAumid; return $result
    }
    Write-Host ""
    Write-Host "-----------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  $($App.Name): could not detect automatically." -ForegroundColor Yellow
    Write-Host "  Step 4: Provide the path/AUMID manually." -ForegroundColor Cyan
    Write-Host "-----------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "  [1] Windows Store app -- enter AUMID (e.g. Microsoft.YourPhone_8wekyb3d8bbwe!App)"
    Write-Host "  [2] Win32 app         -- enter full .exe path"
    Write-Host "  [3] Skip"
    $choice = Read-Host "Select"
    switch ($choice) {
        '1' {
            $raw = (Read-Host "Enter AUMID").Trim().Trim('"')
            if ([string]::IsNullOrWhiteSpace($raw)) { Write-Host "$($App.Name): no input. Skipping."; return $result }
            $validated = Get-StartApps | Where-Object { $_.AppID -eq $raw -or $_.AppID -like "*$raw*" } | Select-Object -First 1
            if ($validated) { Write-Host ("  Validated: '{0}'" -f $validated.AppID) -ForegroundColor Green; $raw=$validated.AppID }
            else { Write-Warning "$($App.Name): '$raw' not found in Get-StartApps. Saving anyway." }
            $repairedArgs = if ($raw -like 'shell:appsFolder\*') { $raw } else { "shell:appsFolder\$raw" }
            if (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf) {
                try {
                    $sc = (New-Object -ComObject WScript.Shell).CreateShortcut($App.ShortcutPath)
                    $sc.Arguments = $repairedArgs; $sc.Save()
                } catch { Write-Warning "$($App.Name): could not write shortcut. $_" }
            }
            $App.ExpectedArguments = $repairedArgs
            $App | Add-Member -NotePropertyName KnownAumid -NotePropertyValue $raw -Force
            # Clear ProcessName so it is re-detected via open window on next run.
            $App.ProcessName = ''
            Export-AppsConfig
            Write-Host ("$($App.Name): AUMID saved, ProcessName cleared for re-detection." ) -ForegroundColor Green
            $result.Resolved=$true; $result.Aumid=$raw
        }
        '2' {
            $exePath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe -ExpectedPublisher $App.ExpectedPublisher
            if ([string]::IsNullOrWhiteSpace($exePath)) { Write-Host "$($App.Name): no path. Skipping."; return $result }
            New-AppShortcut -Path $App.ShortcutPath -TargetPath $exePath -WorkingDirectory (Split-Path $exePath -Parent)
            $App.LaunchType='Win32'; $App.ExpectedArguments=''; $App.ExpectedExe=[System.IO.Path]::GetFileName($exePath)
            Export-AppsConfig
            Write-Host ("$($App.Name): converted to Win32 -> '{0}'." -f $exePath) -ForegroundColor Green
            $result.Resolved=$true; $result.LaunchType='Win32'; $result.ExePath=$exePath
        }
        default { Write-Host "$($App.Name): skipped." }
    }
    return $result
}

# ---------------------------------------------------------------------------
# Repair-ShortcutArguments
# ---------------------------------------------------------------------------
function Repair-ShortcutArguments {
    param($App)
    $repaired = Invoke-ShortcutRepair -App $App -RepairAction {
        param($shortcut)
        Write-Warning "$($App.Name): shortcut Arguments invalid: '$($shortcut.Arguments)'"
        $aumidFragment = $null
        if ($App.ExpectedArguments -match '^shell:appsFolder\\([A-Za-z0-9][A-Za-z0-9._]*_[A-Za-z0-9]+)![A-Za-z0-9._-]+$') {
            $aumidFragment = $Matches[1]
        }
        if ([string]::IsNullOrWhiteSpace($aumidFragment)) {
            Write-Warning "$($App.Name): cannot extract AUMID fragment. Skipping."
            return $null
        }
        $windowsApps = Join-Path $env:ProgramFiles 'WindowsApps'
        $pkgFolder = Get-ChildItem -LiteralPath $windowsApps -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*$aumidFragment*" } | Sort-Object Name -Descending | Select-Object -First 1
        if (-not $pkgFolder) { Write-Warning "$($App.Name): no matching folder in WindowsApps."; return $null }
        $pfn   = (Get-AppxPackage | Where-Object { $_.PackageFamilyName -like "*$aumidFragment*" } | Select-Object -First 1).PackageFamilyName
        $appId = $null
        $manifestPath = Join-Path $pkgFolder.FullName "AppxManifest.xml"
        if (Test-Path -LiteralPath $manifestPath) {
            try {
                $manifest = [xml]::new(); $manifest.Load($manifestPath)
                $appIds = $manifest.Package.Applications.Application.Id
                $appId  = if ($appIds -contains 'App') { 'App' } else { $appIds | Select-Object -First 1 }
            } catch { Write-Warning "$($App.Name): failed to parse AppxManifest.xml. $_" }
        }
        if ([string]::IsNullOrWhiteSpace($pfn) -or [string]::IsNullOrWhiteSpace($appId)) {
            Write-Warning "$($App.Name): could not build new AUMID from manifest/package."
            return $null
        }
        $newAumid = "$pfn!$appId"
        $newArgs  = "shell:appsFolder\$newAumid"
        Write-Host "$($App.Name): repaired Arguments -> '$newArgs'"
        $shortcut.Arguments = $newArgs
        $App.ExpectedArguments = $newArgs
        $App.KnownAumid        = $newAumid
        Export-AppsConfig
        return $newArgs
    }
    return $repaired
}

# ---------------------------------------------------------------------------
# Invoke-LaunchAttempt
# FIX: when Repair-ShortcutArguments returns $null, calls
# Invoke-ManualAumidPrompt instead of silently returning 'Abort'.
# Returns 'Retry' if the user provides a valid AUMID so the caller
# can re-attempt the launch immediately.
# FIX 2/3: passes $App.PresenceMode to Test-AppAlreadyOpen so tray apps
# (OneDrive, ShareFile, Greenshot) are correctly identified as running
# via process presence alone, without requiring a visible window.
# ---------------------------------------------------------------------------
function Invoke-LaunchAttempt {
    param($App)

    # Bootstrap: ensure the .lnk exists before any launch or health-check.
    Initialize-Shortcut -App $App

    # Already open? Pass PresenceMode so tray apps are not relaunched.
    if (Test-AppAlreadyOpen -ProcessName $App.ProcessName -ExpectedExe $App.ExpectedExe `
            -AppName $App.Name -StartAppName $App.StartAppName `
            -PresenceMode $App.PresenceMode) {
        Write-Host "$($App.Name): already open. Skipping.`n"
        return 'Skip'
    }

    if ($App.LaunchType -eq 'Appx') {

        $aumid = Resolve-Aumid -App $App
        if (-not $aumid) {
            $fallback = Resolve-AumidWithFallback -App $App
            if (-not $fallback.Resolved) { return 'Abort' }
            if ($fallback.LaunchType -eq 'Win32') {
                Write-Host "$($App.Name): launching as Win32 after conversion.`n"
                $script:WshShell.Run("`"$($fallback.ExePath)`"", 1, $false)
                Start-Sleep -Seconds $script:PostLaunchPauseSeconds
                return 'Launched'
            }
            $aumid = $fallback.Aumid
        }

        # Verify the shortcut argument matches the resolved AUMID; repair if stale.
        $expectedArgs = "shell:appsFolder\$aumid"
        if ($App.ExpectedArguments -ne $expectedArgs) {
            Write-Host "$($App.Name): stale AUMID in shortcut. Repairing..."
            $repairedArgs = Repair-ShortcutArguments -App $App
            if (-not $repairedArgs) {
                # Automatic repair failed -- offer manual AUMID entry.
                $manual = Invoke-ManualAumidPrompt -App $App
                if ($manual) { return 'Retry' }
                Write-Warning "$($App.Name): argument repair failed. Skipping.`n"
                return 'Abort'
            }
        }

        Write-Host "$($App.Name): launching via Start Menu shortcut (Appx)..."
        $script:WshShell.Run("`"$($App.ShortcutPath)`"", 1, $false)

    } else {

        # Win32 path
        if (-not (Test-ShortcutHealthy -ShortcutPath $App.ShortcutPath -ExpectedPublisher $App.ExpectedPublisher)) {
            Write-Warning "$($App.Name): shortcut unhealthy. Attempting repair..."
            Repair-ShortcutTarget -App $App
            if (-not (Test-ShortcutHealthy -ShortcutPath $App.ShortcutPath -ExpectedPublisher $App.ExpectedPublisher)) {
                Write-Warning "$($App.Name): repair unsuccessful. Skipping.`n"; return 'Abort'
            }
        }
        Write-Host "$($App.Name): launching via shortcut (Win32)..."
        $script:WshShell.Run("`"$($App.ShortcutPath)`"", 1, $false)
    }

    # Back-fill ProcessName from the first visible window if it was empty.
    if ([string]::IsNullOrWhiteSpace($App.ProcessName)) {
        Write-Host "$($App.Name): ProcessName unknown; waiting for window to detect process..."
        $detected = Wait-ForWindowByTitle -App $App -TimeoutSeconds $script:LaunchTimeoutSeconds
        if ($detected -and -not [string]::IsNullOrWhiteSpace($detected.ProcessName)) {
            $App.ProcessName = $detected.ProcessName
            Write-Host "$($App.Name): ProcessName detected as '$($App.ProcessName)'. Saving."
            Export-AppsConfig
        } else {
            Write-Warning "$($App.Name): could not detect ProcessName. Will retry on next run."
        }
    } else {
        $ready = Wait-ForAppReady -ProcessName $App.ProcessName -TimeoutSeconds $script:LaunchTimeoutSeconds
        if (-not $ready) { Write-Warning "$($App.Name): did not become ready within $($script:LaunchTimeoutSeconds)s." }
    }

    Start-Sleep -Seconds $script:PostLaunchPauseSeconds
    return 'Launched'
}

# ---------------------------------------------------------------------------
# Launch sequence
# ---------------------------------------------------------------------------
function Start-AppSequence {
    Write-Host "`n=== Starting app launch sequence ==="
    Write-Host "Initial delay: $script:InitialDelaySeconds seconds..."
    Start-Sleep -Seconds $script:InitialDelaySeconds

    foreach ($app in $script:apps) {
        $maxRetries = 2
        $attempt    = 0
        do {
            $attempt++
            $outcome = Invoke-LaunchAttempt -App $app
            if ($outcome -eq 'Retry' -and $attempt -le $maxRetries) {
                Write-Host "$($app.Name): retrying launch (attempt $attempt of $maxRetries)..."
            }
        } while ($outcome -eq 'Retry' -and $attempt -le $maxRetries)

        if ($outcome -eq 'Retry') {
            Write-Warning "$($app.Name): max retries reached after manual AUMID update. Skipping."
        }
    }

    Write-Host "`n=== Launch sequence complete ===`n"
}

# ---------------------------------------------------------------------------
# Main menu
# ---------------------------------------------------------------------------
function Show-MainMenu {
    while ($true) {
        Write-Host "`n========================================="
        Write-Host "  Win11 Startup Manager"
        Write-Host "========================================="
        Write-Host "  [1] Show configured apps"
        Write-Host "  [2] Add shortcut"
        Write-Host "  [3] Remove shortcut"
        Write-Host "  [4] Modify shortcut"
        Write-Host "  [5] Sync apps from Start Menu"
        Write-Host "  [6] Launch all apps now"
        Write-Host "  [0] Exit"
        $choice = Read-Host "`nSelect"
        switch ($choice) {
            '1' { Show-AppList }
            '2' { Add-Shortcut  -App (Show-AppPicker -Prompt "Select app to ADD shortcut for:"    -AllowNew) }
            '3' { Remove-Shortcut -App (Show-AppPicker -Prompt "Select app to REMOVE shortcut for:") }
            '4' { $picked = Show-AppPicker -Prompt "Select app to MODIFY shortcut for:"
                  if ($picked) { Edit-Shortcut -App $picked } }
            '5' { Sync-AppsFromStartMenu }
            '6' { Start-AppSequence }
            '0' { Write-Host "Exiting."; return }
            default { Write-Warning "Invalid selection. Enter 0-6." }
        }
    }
}

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if ($env:PS_STARTUP_TESTMODE -eq '1') {
    Write-Host "[TEST MODE] Functions loaded. Menu and launch sequence skipped."
} else {
    Show-MainMenu
}
