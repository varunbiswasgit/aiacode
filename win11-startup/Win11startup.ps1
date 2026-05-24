# Win11 Startup Manager
# Launches configured apps at startup via numbered .lnk shortcuts.
# Win32 apps  : launched via WshShell.Run; self-healing shortcut repair on broken target/args.
# Appx apps   : AUMID resolved at runtime (Get-StartApps -> KnownAumid -> AppxPackage manifest).
# Bootstrap   : ensures every Win32 .lnk exists before launch; renames misnumbered or creates fresh.
# Config      : Win11startupapps.json (same folder). Add/Delete/Modify via main menu.
# Test mode   : set $env:PS_STARTUP_TESTMODE = '1' to dot-source without running the menu/sequence.

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
        if ($null -eq $entry.PresenceMode)       { $entry | Add-Member -NotePropertyName PresenceMode       -NotePropertyValue $null -Force }
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
# Sync from Start Menu  [FIX-TEST-08: optional -StartMenuPath param]
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

        if ($leafName -ieq 'explorer.exe' -and $scArgs -like 'shell:appsFolder\*') {
            $expectedExe       = 'explorer.exe'
            $processName       = ''
            $expectedArguments = $scArgs.Trim()
            $knownAumid        = ($scArgs -replace '^shell:appsFolder\\', '').Trim()
            $startAppName      = $appName
            $appxName          = ($knownAumid -split '_')[0]
            Write-Warning "${appName}: ProcessName unknown after sync. Will be auto-detected on first run."
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
        [string]$AppxName          = ''
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
        PresenceMode      = $null
    }
}

# ---------------------------------------------------------------------------
# Boot: resolve config path, then load  [FIX-TEST-09: optional -Path param]
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
function Get-AppPresenceMode {
    param([string]$ProcessName, [int]$SettleSecs = $script:SettleSeconds)
    if ([string]::IsNullOrWhiteSpace($ProcessName)) { return $null }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $SettleSecs) {
        $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) { return 'Window' }
        Start-Sleep -Seconds 1
    }
    if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) { return 'Tray' }
    return $null
}

# FIX-TEST-06: thin wrapper over Get-AppPresenceMode with normalised return values
function Get-AppPresence {
    param([string]$ProcessName, [int]$SettleSecs = $script:SettleSeconds)
    $raw = Get-AppPresenceMode -ProcessName $ProcessName -SettleSecs $SettleSecs
    switch ($raw) {
        'Tray'   { return 'Running' }
        'Window' { return 'WindowVisible' }
        default  { return $null }
    }
}

function Test-AppAlreadyOpen {
    param(
        [string]$ProcessName,
        [string]$ExpectedExe = "",
        [switch]$RequireWindow
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
    if ($RequireWindow) { return [bool]($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) }
    return $true
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
    $mode       = Get-AppPresenceMode -ProcessName $ProcessName -SettleSecs $phase1Secs
    $remaining  = $TimeoutSeconds - $phase1Secs
    if ($null -eq $mode) {
        return Wait-ForProcessCondition -Remaining $remaining -Condition {
            [bool](Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
        }
    }
    Write-Host "  (presence mode: $mode)"
    if ($mode -eq 'Tray') { return $true }
    return Wait-ForProcessCondition -Remaining $remaining -Condition {
        [bool](Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
            Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })
    }
}

# FIX-TEST-07: added -TitleFragment / -TimeoutSeconds overload returning [bool]
function Wait-ForWindowByTitle {
    param(
        $App = $null,
        [string]$TitleFragment = '',
        [int]$WaitSecs = $script:SettleSeconds,
        [int]$TimeoutSeconds = -1
    )
    # -TimeoutSeconds is an alias for -WaitSecs for the testable overload
    if ($TimeoutSeconds -ge 0) { $WaitSecs = $TimeoutSeconds }

    if (-not [string]::IsNullOrWhiteSpace($TitleFragment)) {
        # Testable overload: search by title fragment only, return [bool]
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

    # Original App-object overload: returns matched process or $null
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
    Write-Host "`n================================================"
    Write-Host "  Configured Startup Apps ($($script:apps.Count) total)"
    Write-Host "================================================"
    Write-Host ("{0,-4} {1,-22} {2,-6} {3,-10} {4}" -f '#', 'Name', 'Type', 'Shortcut', 'Process')
    Write-Host ("{0,-4} {1,-22} {2,-6} {3,-10} {4}" -f '-'*3, '-'*21, '-'*5, '-'*8, '-'*15)
    for ($i = 0; $i -lt $script:apps.Count; $i++) {
        $app    = $script:apps[$i]
        $status = if (Test-Path -LiteralPath $app.ShortcutPath -PathType Leaf) { 'OK' } else { 'MISSING' }
        $color  = if ($status -eq 'OK') { 'Green' } else { 'Yellow' }
        Write-Host ("{0,-4} {1,-22} {2,-6} {3,-10} {4}" -f ($i+1), $app.Name, $app.LaunchType, $status, $app.ProcessName) -ForegroundColor $color
    }
    Write-Host ""
}

# ---------------------------------------------------------------------------
# Shortcut writer (single .lnk creation point)  [FIX-TEST-03: -App param overload]
# ---------------------------------------------------------------------------
function New-AppShortcut {
    param(
        [string]$Path = '',
        [string]$TargetPath = '',
        [string]$Arguments = '',
        [string]$WorkingDirectory = '',
        [PSCustomObject]$App = $null
    )
    # -App overload: derives Path/TargetPath/Arguments/WorkingDirectory from App object
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
# Exe validation (allowlist + signature)
# ---------------------------------------------------------------------------
function Test-ExePathAllowed {
    param([string]$ExePath)
    $full = [System.IO.Path]::GetFullPath($ExePath)
    foreach ($root in $script:AllowedExeRoots) {
        $rootFull = [System.IO.Path]::GetFullPath($root.TrimEnd('\'))
        if ($full.StartsWith($rootFull + '\', [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Test-ExeSignatureTrusted {
    param([string]$ExePath, [string]$ExpectedPublisher = "")
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

function Test-ExeAcceptable {
    param([string]$ExePath, [string]$ExpectedPublisher = "")
    if (-not (Test-ExePathAllowed -ExePath $ExePath)) {
        Write-Warning "Path outside allowed roots: $ExePath"; return $false
    }
    if (-not (Test-ExeSignatureTrusted -ExePath $ExePath -ExpectedPublisher $ExpectedPublisher)) {
        return $false
    }
    return $true
}

# ---------------------------------------------------------------------------
# Shortcut health check  [FIX-TEST-05: standalone Test-ShortcutHealthy]
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
        Repair-ShortcutArguments -App $App; return
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

# FIX-TEST-01: added [Alias('LnkPath')] to -ShortcutPath param
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

# FIX-TEST-02: added -SearchRoot / -ExeName aliases; optional -MaxDepth param
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

function Initialize-Shortcut {
    param($App)
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
# Shortcut repair (Invoke-ShortcutRepair owns Get-ShortcutObject + .Save())
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
    return Invoke-ShortcutRepair -App $App -RepairAction {
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

function Repair-ShortcutArguments {
    param($App)
    return Invoke-ShortcutRepair -App $App -RepairAction {
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
            } catch { Write-Warning "$($App.Name): cannot read manifest. $_" }
        }
        if ([string]::IsNullOrWhiteSpace($appId)) {
            Write-Warning "$($App.Name): manifest unreadable. Falling back to ExpectedArguments AppId."
            $appId = ($App.ExpectedArguments -split '!') | Select-Object -Last 1
        }
        if ([string]::IsNullOrWhiteSpace($pfn)) { Write-Warning "$($App.Name): package not found for '$aumidFragment'."; return $null }
        $repairedArgs = "shell:appsFolder\$pfn!$appId"
        Write-Host "$($App.Name): AUMID repaired to $pfn!$appId."
        $shortcut.Arguments        = $repairedArgs
        $shortcut.WorkingDirectory = Split-Path -Path $shortcut.TargetPath -Parent
        return $repairedArgs
    }
}

# ---------------------------------------------------------------------------
# Failure menu / recovery
# ---------------------------------------------------------------------------
function Show-FailureMenu {
    param([string]$AppName, [string]$Context)
    Write-Host "  [1] Add / fix shortcut for $AppName and retry"
    Write-Host "  [2] Modify a different shortcut"
    Write-Host "  [3] Skip"
    Write-Host "  [4] Delete '$AppName' entry from config"
    return (Read-Host "Select ($Context)")
}

function Invoke-AppLaunchWait {
    param($App, [int]$TimeoutSeconds = $script:LaunchTimeoutSeconds)
    if (Wait-ForAppReady -ProcessName $App.ProcessName -TimeoutSeconds $TimeoutSeconds) {
        Write-Host "$($App.Name): ready.`n"; Start-Sleep -Seconds $script:PostLaunchPauseSeconds; return $true
    }
    Write-Warning "$($App.Name): did not become ready within $TimeoutSeconds seconds."
    return $false
}

function Invoke-FailureRecovery {
    param($App, [string]$Context, [scriptblock]$PreRetryAction = $null)
    $choice = Show-FailureMenu -AppName $App.Name -Context $Context
    switch ($choice) {
        '1' { if ($PreRetryAction) { & $PreRetryAction }; return $true }
        '2' { $target = Show-AppPicker -Prompt "Select shortcut to modify:"; if ($target) { Edit-Shortcut -App $target }; return $false }
        '4' { Remove-Shortcut -App $App; return $false }
        default { Write-Host "$($App.Name): skipped.`n"; return $false }
    }
}

# ---------------------------------------------------------------------------
# Launch functions
# ---------------------------------------------------------------------------
function Start-AppxApp {
    # Appx has no failure-recovery retry by design: failures are AUMID-resolution
    # failures that Resolve-Aumid already exhausts three paths for. Use menu [4]
    # Modify to update Win11startupapps.json directly if AUMID changes.
    param($App)
    if (Test-AppAlreadyOpen -ProcessName $App.ProcessName) { Write-Host "$($App.Name): already open. Skipping.`n"; return $true }
    $aumid = Resolve-Aumid -App $App
    if (-not $aumid) { Write-Warning "$($App.Name): no AUMID found. Skipping.`n"; return $false }
    try {
        Write-Host "$($App.Name): launching via shell:appsFolder\$aumid"
        Start-Process explorer.exe "shell:appsFolder\$aumid" -ErrorAction Stop
        return Invoke-AppLaunchWait -App $App
    } catch { Write-Warning "$($App.Name): launch failed. $_`n"; return $false }
}

# FIX-TEST-04: thin wrapper Invoke-AppLaunch over Invoke-LaunchAttempt
function Invoke-AppLaunch {
    param($App)
    return Invoke-LaunchAttempt -App $App
}

function Invoke-LaunchAttempt {
    # One repair+launch+wait cycle. Returns 'Success', 'Retry', or 'Abort'.
    param($App)

    if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
        Write-Warning "$($App.Name): shortcut not found: $($App.ShortcutPath)"
        $recover = Invoke-FailureRecovery -App $App -Context "missing shortcut" -PreRetryAction { Initialize-Shortcut -App $App }
        if ($recover) { return 'Retry' } else { return 'Abort' }
    }

    try {
        $shortcut   = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
        $targetPath = $shortcut.TargetPath
        if ([string]::IsNullOrWhiteSpace($targetPath) -or -not (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
            $repairedPath = Repair-ShortcutTarget -App $App
            if (-not $repairedPath) { Write-Warning "$($App.Name): repair failed. Skipping.`n"; return 'Abort' }
            $shortcut = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
        }

        if (-not [string]::IsNullOrWhiteSpace($App.ExpectedArguments)) {
            $currentArgs = $shortcut.Arguments
            if ([string]::IsNullOrWhiteSpace($currentArgs) -or $currentArgs -notlike "*$($App.ExpectedArguments)*") {
                $repairedArgs = Repair-ShortcutArguments -App $App
                if (-not $repairedArgs) { Write-Warning "$($App.Name): argument repair failed. Skipping.`n"; return 'Abort' }
            }
        }

        Write-Host "$($App.Name): launching via $($App.ShortcutPath)"
        $script:WshShell.Run("`"$($App.ShortcutPath)`"", 1, $false)

        if ([string]::IsNullOrWhiteSpace($App.ProcessName)) {
            Write-Host "$($App.Name): ProcessName blank - waiting for active window to confirm launch..."
            $matchedProc = Wait-ForWindowByTitle -App $App -WaitSecs $script:SettleSeconds
            if ($matchedProc) {
                $App.ProcessName = $matchedProc.ProcessName
                Export-AppsConfig
                Write-Host "$($App.Name): active window detected. ProcessName back-filled as '$($App.ProcessName)' and saved."
            } else {
                Write-Warning "$($App.Name): no matching active window detected within $script:SettleSeconds seconds. Launch assumed but ready-check skipped."
            }
        }

        $ready = Invoke-AppLaunchWait -App $App
        if ($null -eq $App.PresenceMode -and -not [string]::IsNullOrWhiteSpace($App.ProcessName)) {
            $resolvedMode = Get-AppPresenceMode -ProcessName $App.ProcessName -SettleSecs 0
            if ($resolvedMode) { $App.PresenceMode = $resolvedMode }
            Export-AppsConfig
        }
        if ($ready) { return 'Success' }

        $recover = Invoke-FailureRecovery -App $App -Context "launch timeout" -PreRetryAction { Edit-Shortcut -App $App }
        if ($recover) { return 'Retry' } else { return 'Abort' }

    } catch {
        Write-Warning "$($App.Name): launch exception. $_`n"
        Write-ErrorLog -Message "$($App.Name): launch exception" -ErrorRecord $_
        $recover = Invoke-FailureRecovery -App $App -Context "launch exception" -PreRetryAction { Edit-Shortcut -App $App }
        if ($recover) { return 'Retry' } else { return 'Abort' }
    }
}

# FIX-TEST-10: -MaxAttempts [int] param defaulting to 3
function Start-Win32App {
    param($App, [int]$MaxAttempts = 3)
    $requireWin = ($App.PresenceMode -eq 'Window')
    if (Test-AppAlreadyOpen -ProcessName $App.ProcessName -ExpectedExe $App.ExpectedExe -RequireWindow:$requireWin) {
        Write-Host "$($App.Name): already open. Skipping.`n"; return $true
    }
    for ($attempt = 0; $attempt -lt $MaxAttempts; $attempt++) {
        switch (Invoke-LaunchAttempt -App $App) {
            'Success' { return $true }
            'Abort'   { return $false }
        }
    }
    Write-Warning "$($App.Name): max attempts reached. Skipping.`n"
    return $false
}

# ---------------------------------------------------------------------------
# Main menu + startup sequence
# ---------------------------------------------------------------------------
if ($env:PS_STARTUP_TESTMODE -eq '1') { return }

Write-Host "`n================================================"
Write-Host "  Win11 Startup Manager"
Write-Host "================================================"
Write-Host "  [1] Run startup sequence"
Write-Host "  [2] Add shortcut"
Write-Host "  [3] Delete shortcut"
Write-Host "  [4] Modify shortcut"
Write-Host "  [5] List startup apps"
Write-Host "  [6] Sync from Start Menu"
Write-Host "  [7] Exit"
Write-Host "------------------------------------------------"
$mainChoice = Read-Host "Select"

switch ($mainChoice) {
    '2' { $app = Show-AppPicker -Prompt "Select app to re-initialise, or [N] to add a new entry:" -AllowNew; Add-Shortcut -App $app; exit }
    '3' { $app = Show-AppPicker -Prompt "Select app to DELETE shortcut for:"; if ($app) { Remove-Shortcut -App $app }; exit }
    '4' { $app = Show-AppPicker -Prompt "Select app to MODIFY shortcut for:"; if ($app) { Edit-Shortcut -App $app }; exit }
    '5' { Show-AppList; exit }
    '6' { Sync-AppsFromStartMenu; exit }
    '7' { exit }
}

Write-Host "Waiting $script:InitialDelaySeconds seconds for system to stabilize..."
Start-Sleep -Seconds $script:InitialDelaySeconds

Write-Host "`n--- Shortcut bootstrap ---"
foreach ($app in $script:apps) {
    if ($app.LaunchType -eq 'Win32') { Initialize-Shortcut -App $app }
}
Write-Host "--- Bootstrap complete ---`n"

$failedApps = @()
foreach ($app in $script:apps) {
    $ok = if ($app.LaunchType -eq 'Appx') { Start-AppxApp -App $app } else { Start-Win32App -App $app }
    if (-not $ok) { $failedApps += "$($app.Name) [$($app.ProcessName)]" }
}

if ($failedApps.Count -gt 0) {
    Write-Host "`n--- Startup completed with failures ---"
    foreach ($entry in $failedApps) { Write-Host "  - $entry" }
} else {
    Write-Host "`nStartup sequence completed successfully."
}
