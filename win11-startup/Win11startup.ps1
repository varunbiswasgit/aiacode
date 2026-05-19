# Curated startup launcher with self-healing shortcut repair
# - Win32 apps     : shortcut invoked via WshShell.Run when target is valid (preserves baked-in arguments)
#                    self-healing repair + user-prompt fallback used only when shortcut target is broken
#                    repaired shortcuts are then invoked via WshShell.Run as well
# - Appx apps      : AUMID resolved at runtime (Get-StartApps -> KnownAumid verification -> AppxPackage manifest)
#                    KnownAumid used only as primary candidate, not sole source of truth
# - Sticky Notes   : Win32 shortcut with /memoryWindow start baked into the .lnk Target field
#                    WshShell.Run fires the shortcut as-is; no separate Arguments field needed
# - Packaged apps  : Win32 entries that cannot be launched via direct .exe invocation (e.g. apps installed
#                    under WindowsApps which is ACL-locked) are stored as .lnk targeting explorer.exe
#                    with shell:appsFolder\<AUMID> as the Arguments field. ExpectedArguments triggers
#                    argument self-healing when the AUMID becomes stale after a package update.
# - Bootstrap      : before launch loop, ensures every Win32 .lnk exists at the expected path;
#                    renames misnumbered matches found in the same folder, or creates fresh if absent
# - Main menu      : on launch, user chooses Run / Add / Delete / Modify / List / Exit
#                    inline failure menu (Add+retry / Modify / Skip) appears when a shortcut
#                    is missing or an app fails to start during the startup sequence
# - Presence mode  : after launch, Get-AppPresenceMode polls MainWindowHandle for $script:SettleSeconds;
#                    if a window appears -> 'Window' mode (skip only when window visible);
#                    if no window appears -> 'Tray' mode (skip when process running).
#                    No per-app flags needed; detection is fully automatic at runtime.
# - Exe allowlist  : Test-ExePathAllowed enforces that any exe accepted via user prompt or
#                    auto-discovery during repair lives under Program Files, Program Files (x86),
#                    or Windows. Paths outside these roots are rejected before any shortcut write.
# - Signature gate : Test-ExeSignatureTrusted requires Get-AuthenticodeSignature status 'Valid'
#                    before any repaired or user-supplied exe is persisted to a shortcut target.
#                    Called after the allowlist check in both Prompt-ForExactExePath and
#                    Repair-ShortcutTarget.
# - XML load       : Repair-ShortcutArguments uses [xml]::new() + Load() instead of
#                    [xml]$m = Get-Content to correctly handle BOMs and large manifest files.
# - Publisher gate : Test-ExeSignatureTrusted accepts an optional -ExpectedPublisher string.
#                    When present, the SignerCertificate.Subject must contain that string.
#                    Each app entry carries an optional ExpectedPublisher field; Microsoft apps
#                    use 'CN=Microsoft Corporation', Chrome uses 'CN=Google LLC'.
# - Process guard  : Test-AppAlreadyOpen now compares each matching process MainModule filename
#                    against ExpectedExe before treating the app as already open, preventing
#                    false skips caused by unrelated same-named processes.
# - Regex anchor   : Repair-ShortcutArguments uses a fully anchored regex to extract the PFN
#                    from ExpectedArguments. Pattern requires the full value to match
#                    ^shell:appsFolder\<PFN>!<AppId>$ and accepts any valid PackageFamilyName
#                    prefix (not limited to Microsoft.) so non-Microsoft packaged apps are
#                    handled correctly.
# - Script scope   : All shared vars ($WshShell, $apps, $startMenu,
#                    $InitialDelaySeconds, $LaunchTimeoutSeconds, $PostLaunchPauseSeconds,
#                    $SettleSeconds, $AllowedExeRoots) use $script: scope so dot-sourced
#                    Pester runs cannot leak or shadow globals.
# - Test mode      : When $env:PS_STARTUP_TESTMODE = '1', the script dot-sources cleanly
#                    (functions and vars available) but skips the interactive menu and
#                    startup sequence entirely. Used by Win11startup.Tests.ps1.
# - Config file    : $script:apps is loaded from apps.json in the same folder as the script.
#                    Required fields: Name, LaunchType, ShortcutPath, ProcessName, ExpectedExe.
#                    Add/Delete menu flows write changes back to apps.json automatically.
# - Add flow       : Show-AppPicker -AllowNew shows [N] for brand-new entry and returns
#                    '__NEW__' sentinel; Add-Shortcut distinguishes re-init (real app object),
#                    new entry ('__NEW__'), and cancel ($null). Appx entries collect
#                    StartAppName, KnownAumid, AppxName during the add prompt.
# - Timeout math   : Wait-ForAppReady stores actual phase-1 settle duration in $phase1Secs
#                    and subtracts that exact value for phase-2 remaining, so the total
#                    timeout is always honoured even when TimeoutSeconds < SettleSeconds.
# - Failure menu   : Show-FailureMenu centralises the 3-option inline prompt (Add+retry /
#                    Modify / Skip) used by both the missing-shortcut and launch-timeout
#                    paths in Start-Win32App.
# - Launch wait    : Invoke-AppLaunchWait centralises the Wait-ForAppReady + ready/warning
#                    output + PostLaunchPauseSeconds sleep shared by Start-AppxApp and
#                    the happy path of Start-Win32App.
# - Shortcut write : New-AppShortcut -Path -TargetPath -Arguments -WorkingDirectory is the
#                    single place all .lnk files are written. Add-Shortcut, Edit-Shortcut,
#                    and Initialize-Shortcut all call it; WshShell.CreateShortcut is invoked
#                    in exactly one function.
# - File search    : Get-ParentFolder walks two Split-Path levels above the broken target's
#                    folder (grandparent). Falls back to one level up when the grandparent
#                    path is empty or does not exist. Repair-ShortcutTarget then searches
#                    recursively from that root with no depth cap.
# - Appx enum      : Resolve-Aumid calls Get-AppxPackage once, stores the result in $pkgs,
#                    and reuses it for both KnownAumid verification and the AppxName fallback,
#                    avoiding a second full pipeline enumeration.
# - Error logging  : All unhandled errors and terminating exceptions are written to
#                    startup-error.log in the same folder as the script, with timestamps,
#                    so crashes can be diagnosed after PowerShell has closed.
# - Depth filter   : Find-ExeWithinDepth no longer applies a Get-RelativeDepth depth cap.
#                    Get-ChildItem -Recurse searches freely; Get-RelativeDepth is kept for
#                    test coverage but is no longer called in the main search path. (T-06)
# - Stopwatch      : Get-AppPresenceMode and Wait-ForAppReady use
#                    [System.Diagnostics.Stopwatch] instead of manual $elapsed++ counters
#                    so elapsed time reflects wall clock even when Get-Process is slow. (T-07)
# - PFN source     : Repair-ShortcutArguments uses $pkg.PackageFamilyName directly from
#                    Get-AppxPackage instead of reconstructing it via regex, avoiding
#                    breakage when version/arch segments vary. (FIX-05)
# - SystemRoot     : Initialize-Shortcut uses $env:SystemRoot instead of hardcoded
#                    'C:\Windows' for the explorer.exe target and working directory
#                    in Appx shortcut creation. (FIX-06)
# - Stale object   : Start-Win32App no longer re-reads the shortcut object after
#                    Repair-ShortcutTarget; WshShell.Run always uses $App.ShortcutPath
#                    directly so the extra Get-ShortcutObject call was misleading. (FIX-07)
# - Export guard   : Export-AppsConfig wraps Set-Content in try/catch; a write failure
#                    is logged to startup-error.log and reported to the console instead
#                    of silently discarding the update. (ROB-01)
# - Edit guard     : Edit-Shortcut argument-repair branch guards Get-ShortcutObject with
#                    Test-Path before calling it, matching all other callers. (ROB-02)
# - Retry loop     : Start-Win32App uses Invoke-FailureRecovery with a bounded for-loop
#                    (max 2 retries) instead of unbounded self-recursion. Both the
#                    missing-shortcut and launch-timeout paths share the same helper,
#                    eliminating the duplicated switch blocks. (ROB-04 + QOL-04)
# - Path attempts  : Prompt-ForExactExePath limits validation retries to 3 attempts;
#                    returns $null after exhaustion so the caller handles it gracefully
#                    instead of looping forever. (HARD-04)
# - Number valid   : Add-Shortcut validates the shortcut number is 1-2 digits (re-prompts
#                    on blank or non-numeric input) before constructing the .lnk filename.
#                    (HARD-05)
# - Picker perf    : Show-AppPicker pre-computes each entry's shortcut-existence status
#                    before the display loop, avoiding repeated Test-Path calls per
#                    input attempt. (QOL-01)
# - Window switch  : Test-AppAlreadyOpen accepts -RequireWindow; when set, a process with
#                    no MainWindowHandle is not treated as open. Startup sequence passes
#                    -RequireWindow for Window-mode apps. (FIX-04)
# - AUMID log      : Resolve-Aumid logs all-paths failure to startup-error.log via
#                    Write-ErrorLog in addition to emitting Write-Warning. (QOL-02)
# - Schema version : Import-AppsConfig checks for top-level "schemaVersion" field;
#                    warns if absent or not 1. apps.json written by Export-AppsConfig
#                    includes schemaVersion:1 wrapper. (QOL-03)
# - List apps      : Main menu option [5] prints a formatted table of all configured
#                    startup apps (name, type, shortcut status, process) then exits.
#                    Implemented via Show-AppList. (QOL-05)
# - Stale shortcut : Start-Win32App re-reads the shortcut object after Repair-ShortcutTarget
#                    so the argument-repair check uses the freshly written .lnk, not the
#                    pre-repair in-memory object. (INT-02)
# - WindowsApps    : Repair-ShortcutArguments uses Join-Path $env:ProgramFiles 'WindowsApps'
#                    instead of a hardcoded 'C:\Program Files\WindowsApps' path. (INT-01)
# - AppList fix    : Show-AppList assigns status via bare string ('OK'/'MISSING') instead of
#                    Write-Output inside an if-expression to avoid pipeline capture. (BUG-02)

# ---------------------------------------------------------------------------
# Error log helper — writes timestamped entries to $PSScriptRoot\startup-error.log
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

$script:startMenu              = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
$script:InitialDelaySeconds    = 10
$script:LaunchTimeoutSeconds   = 30
$script:PostLaunchPauseSeconds = 2
$script:SettleSeconds          = 5
$script:AppsConfigPath         = Join-Path $PSScriptRoot "apps.json"
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
        throw "apps.json not found at '$Path'. Cannot continue without app configuration."
    }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $parsed = $raw | ConvertFrom-Json

    if ($parsed -is [System.Array]) {
        Write-Warning "apps.json has no schemaVersion wrapper. Expected schemaVersion $script:AppsConfigSchemaVersion. Processing as legacy format."
        $entries = $parsed
    } else {
        $sv = $parsed.schemaVersion
        if ($null -eq $sv) {
            Write-Warning "apps.json schemaVersion field missing. Expected $script:AppsConfigSchemaVersion."
        } elseif ([int]$sv -ne $script:AppsConfigSchemaVersion) {
            Write-Warning "apps.json schemaVersion is '$sv'; expected '$script:AppsConfigSchemaVersion'. Proceeding with caution."
        }
        $entries = $parsed.apps
    }

    $required = @('Name','LaunchType','ShortcutPath','ProcessName','ExpectedExe')
    $validated = @()
    foreach ($entry in $entries) {
        foreach ($field in $required) {
            if ([string]::IsNullOrWhiteSpace($entry.$field)) {
                throw "apps.json entry missing required field '$field': $(ConvertTo-Json $entry -Compress)"
            }
        }
        if ($null -eq $entry.ExpectedPublisher)  { $entry | Add-Member -NotePropertyName ExpectedPublisher  -NotePropertyValue '' -Force }
        if ($null -eq $entry.ExpectedArguments)  { $entry | Add-Member -NotePropertyName ExpectedArguments  -NotePropertyValue '' -Force }
        if ($null -eq $entry.StartAppName)        { $entry | Add-Member -NotePropertyName StartAppName        -NotePropertyValue '' -Force }
        if ($null -eq $entry.KnownAumid)          { $entry | Add-Member -NotePropertyName KnownAumid          -NotePropertyValue '' -Force }
        if ($null -eq $entry.AppxName)            { $entry | Add-Member -NotePropertyName AppxName            -NotePropertyValue '' -Force }
        $validated += $entry
    }
    return $validated
}

function Export-AppsConfig {
    param([string]$Path = $script:AppsConfigPath)
    try {
        $wrapper = [PSCustomObject]@{
            schemaVersion = $script:AppsConfigSchemaVersion
            apps          = $script:apps
        }
        $wrapper | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
        Write-Host "apps.json saved (schemaVersion $script:AppsConfigSchemaVersion, $($script:apps.Count) entries)."
    } catch {
        Write-ErrorLog -Message "Export-AppsConfig failed to write '$Path'" -ErrorRecord $_
        Write-Warning "apps.json could not be saved: $($_.Exception.Message)"
        Write-Host "Error details written to: $script:ErrorLogPath" -ForegroundColor Yellow
    }
}

try {
    $script:apps = Import-AppsConfig
} catch {
    Write-ErrorLog -Message "FATAL: Import-AppsConfig failed" -ErrorRecord $_
    Write-Error $_
    Write-Host "Error details written to: $script:ErrorLogPath" -ForegroundColor Yellow
    exit 1
}

$script:WshShell = New-Object -ComObject WScript.Shell

# ---------------------------------------------------------------------------
# Presence mode detection
# ---------------------------------------------------------------------------
function Get-AppPresenceMode {
    param([string]$ProcessName, [int]$SettleSecs = $script:SettleSeconds)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $SettleSecs) {
        $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) {
            return 'Window'
        }
        Start-Sleep -Seconds 1
    }
    if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) {
        return 'Tray'
    }
    return $null
}

function Test-AppAlreadyOpen {
    param(
        [string]$ProcessName,
        [string]$ExpectedExe    = "",
        [switch]$RequireWindow
    )
    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if (-not $procs) { return $false }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedExe)) {
        $procs = $procs | Where-Object {
            try {
                [System.IO.Path]::GetFileName($_.MainModule.FileName) -ieq $ExpectedExe
            } catch {
                $false
            }
        }
        if (-not $procs) { return $false }
    }

    if ($RequireWindow) {
        return [bool]($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })
    }

    return $true
}

function Wait-ForAppReady {
    param([string]$ProcessName, [int]$TimeoutSeconds = $script:LaunchTimeoutSeconds)
    $phase1Secs = [Math]::Min($script:SettleSeconds, $TimeoutSeconds)
    $mode = Get-AppPresenceMode -ProcessName $ProcessName -SettleSecs $phase1Secs
    if ($null -eq $mode) {
        $remaining = $TimeoutSeconds - $phase1Secs
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        while ($sw.Elapsed.TotalSeconds -lt $remaining) {
            if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) { return $true }
            Start-Sleep -Seconds 1
        }
        return $false
    }

    Write-Host "  (presence mode: $mode)"

    if ($mode -eq 'Tray') { return $true }

    $remaining = $TimeoutSeconds - $phase1Secs
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $remaining) {
        $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) { return $true }
        Start-Sleep -Seconds 1
    }
    return $false
}

# ---------------------------------------------------------------------------
# Helper: numbered app picker
# ---------------------------------------------------------------------------
function Show-AppPicker {
    param([string]$Prompt, [switch]$AllowNew)
    $statuses = $script:apps | ForEach-Object {
        if (Test-Path -LiteralPath $_.ShortcutPath -PathType Leaf) { 'exists' } else { 'missing' }
    }
    Write-Host "`n$Prompt"
    for ($i = 0; $i -lt $script:apps.Count; $i++) {
        Write-Host ("  [{0}] {1,-20}  {2}  ({3})" -f ($i + 1), $script:apps[$i].Name, $statuses[$i], $script:apps[$i].ShortcutPath)
    }
    if ($AllowNew) {
        Write-Host "  [N] Add a brand-new app entry"
    }
    Write-Host "  [0] Cancel"
    while ($true) {
        $choice = Read-Host "Select"
        if ($AllowNew -and $choice -imatch '^n$')  { return '__NEW__' }
        if ($choice -match '^\d+$') {
            $idx = [int]$choice
            if ($idx -eq 0)                                          { return $null }
            if ($idx -ge 1 -and $idx -le $script:apps.Count)        { return $script:apps[$idx - 1] }
        }
        $hint = if ($AllowNew) { "0-$($script:apps.Count) or N" } else { "0-$($script:apps.Count)" }
        Write-Warning "Invalid selection. Enter $hint."
    }
}

# ---------------------------------------------------------------------------
# QOL-05: Show-AppList — formatted table of all configured startup apps.
# BUG-02: status assigned via bare string, not Write-Output inside if-expression.
# ---------------------------------------------------------------------------
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
        $line   = "{0,-4} {1,-22} {2,-6} {3,-10} {4}" -f ($i + 1), $app.Name, $app.LaunchType, $status, $app.ProcessName
        Write-Host $line -ForegroundColor $color
    }
    Write-Host ""
}

# ---------------------------------------------------------------------------
# Single shortcut writer
# ---------------------------------------------------------------------------
function New-AppShortcut {
    param(
        [string]$Path,
        [string]$TargetPath,
        [string]$Arguments        = "",
        [string]$WorkingDirectory = ""
    )
    $sc = $script:WshShell.CreateShortcut($Path)
    $sc.TargetPath = $TargetPath
    if (-not [string]::IsNullOrWhiteSpace($Arguments))        { $sc.Arguments        = $Arguments }
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) { $sc.WorkingDirectory = $WorkingDirectory }
    $sc.Save()
}

# ---------------------------------------------------------------------------
# Shortcut management
# ---------------------------------------------------------------------------
function Add-Shortcut {
    param($App)
    if ($null -eq $App)          { Write-Host "Add cancelled."; return }
    if ($App -ne '__NEW__') {
        Initialize-Shortcut -App $App
        return
    }

    Write-Host "`n--- Add new app entry ---"
    $name = Read-Host "App name (display label)"
    if ([string]::IsNullOrWhiteSpace($name)) { Write-Host "Cancelled."; return }

    $launchType = ''
    while ($launchType -notin @('Win32','Appx')) {
        $launchType = Read-Host "Launch type [Win32 / Appx]"
    }

    $number = ''
    while ($number -notmatch '^\d{1,2}$') {
        $number = Read-Host "Shortcut number (1-2 digits, e.g. 09)"
        if ($number -notmatch '^\d{1,2}$') {
            Write-Warning "Shortcut number must be 1-2 digits (e.g. 01, 9). Try again."
        }
    }
    $lnkName      = "$number $name.lnk"
    $shortcutPath = Join-Path $script:startMenu $lnkName

    $processName = Read-Host "Process name (without .exe, e.g. chrome)"

    $expectedExe = if ($launchType -eq 'Appx') { 'explorer.exe' } else {
        Read-Host "Expected exe filename (e.g. chrome.exe)"
    }

    $expectedPublisher = Read-Host "Expected publisher CN string (optional, press Enter to skip)"
    $expectedArguments = ''

    $startAppName = ''
    $knownAumid   = ''
    $appxName     = ''
    if ($launchType -eq 'Appx') {
        Write-Host "(Appx fields — used by AUMID resolver at launch time)"
        $startAppName      = Read-Host "Start menu display name fragment (e.g. Your Phone)"
        $knownAumid        = Read-Host "Known AUMID (e.g. Microsoft.YourPhone_8wekyb3d8bbwe!App, press Enter to skip)"
        $appxName          = Read-Host "Appx package name fragment (e.g. YourPhone, press Enter to skip)"
        $expectedArguments = if (-not [string]::IsNullOrWhiteSpace($knownAumid)) {
                                 "shell:appsFolder\$knownAumid"
                             } else { '' }
    } else {
        $expectedArguments = Read-Host "Expected arguments (optional, e.g. shell:appsFolder\PFN!App, press Enter to skip)"
    }

    $exePath = $null
    if ($launchType -eq 'Win32' -and [string]::IsNullOrWhiteSpace($expectedArguments)) {
        $exePath = Prompt-ForExactExePath -AppName $name -ExpectedExe $expectedExe -ExpectedPublisher $expectedPublisher
        if (-not $exePath) { Write-Host "Cancelled."; return }
    }

    $newEntry = [PSCustomObject]@{
        Name              = $name
        LaunchType        = $launchType
        ShortcutPath      = $shortcutPath
        ProcessName       = $processName
        ExpectedExe       = $expectedExe
        ExpectedPublisher = $expectedPublisher
        ExpectedArguments = $expectedArguments
        StartAppName      = $startAppName
        KnownAumid        = $knownAumid
        AppxName          = $appxName
    }

    if ($exePath) {
        New-AppShortcut -Path $shortcutPath -TargetPath $exePath `
                        -WorkingDirectory (Split-Path -Path $exePath -Parent)
        Write-Host "Shortcut created: $shortcutPath"
    } elseif (-not [string]::IsNullOrWhiteSpace($expectedArguments)) {
        New-AppShortcut -Path $shortcutPath -TargetPath "$env:SystemRoot\explorer.exe" `
                        -Arguments $expectedArguments -WorkingDirectory $env:SystemRoot
        Write-Host "Shortcut created with arguments: $shortcutPath"
    }

    $script:apps += $newEntry
    Export-AppsConfig
    Write-Host "'$name' added to apps.json."
}

function Remove-Shortcut {
    param($App)
    if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
        Write-Warning "$($App.Name): no shortcut found at '$($App.ShortcutPath)'."
    } else {
        $confirm = Read-Host "Delete '$($App.ShortcutPath)'? (Y/N)"
        if ($confirm -ieq 'Y') {
            Remove-Item -LiteralPath $App.ShortcutPath -Force
            Write-Host "$($App.Name): shortcut deleted."
        } else {
            Write-Host "$($App.Name): deletion cancelled."
            return
        }
    }

    $confirm2 = Read-Host "Also remove '$($App.Name)' from apps.json? (Y/N)"
    if ($confirm2 -ieq 'Y') {
        $script:apps = $script:apps | Where-Object { $_ -ne $App }
        Export-AppsConfig
        Write-Host "$($App.Name): removed from apps.json."
    }
}

function Edit-Shortcut {
    param($App)
    if (-not [string]::IsNullOrWhiteSpace($App.ExpectedArguments)) {
        if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
            Write-Warning "$($App.Name): shortcut does not exist. Creating it first."
            Initialize-Shortcut -App $App
            return
        }
        Write-Host "$($App.Name): re-running argument repair..."
        Repair-ShortcutArguments -App $App
        return
    }
    $exePath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe -ExpectedPublisher $App.ExpectedPublisher
    if ($exePath) {
        if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
            New-AppShortcut -Path $App.ShortcutPath -TargetPath $exePath `
                            -WorkingDirectory (Split-Path -Path $exePath -Parent)
            Write-Host "$($App.Name): shortcut created at '$($App.ShortcutPath)'."
        } else {
            Update-ShortcutTarget -ShortcutPath $App.ShortcutPath -ExePath $exePath
            Write-Host "$($App.Name): shortcut updated to '$exePath'."
        }
    } else {
        Write-Host "$($App.Name): modify cancelled."
    }
}

# ---------------------------------------------------------------------------
# Core helpers
# ---------------------------------------------------------------------------
function Resolve-Aumid {
    param($App)
    $startApp = Get-StartApps | Where-Object { $_.Name -like "*$($App.StartAppName)*" } | Select-Object -First 1
    if ($startApp) {
        Write-Host "$($App.Name): AUMID resolved via Get-StartApps: $($startApp.AppID)"
        return $startApp.AppID
    }
    $pkgs = @(Get-AppxPackage)
    if (-not [string]::IsNullOrWhiteSpace($App.KnownAumid)) {
        $knownPfn  = ($App.KnownAumid -split '!')[0]
        $installed = $pkgs | Where-Object { $_.PackageFamilyName -eq $knownPfn } | Select-Object -First 1
        if ($installed) {
            Write-Host "$($App.Name): KnownAumid verified as installed: $($App.KnownAumid)"
            return $App.KnownAumid
        }
        Write-Warning "$($App.Name): KnownAumid package family '$knownPfn' not found on this system."
    }
    $pkg = $pkgs | Where-Object { $_.Name -like "*$($App.AppxName)*" } | Select-Object -First 1
    if ($pkg) {
        try {
            $appIds = (Get-AppxPackageManifest $pkg).Package.Applications.Application.Id
            $appId  = if ($appIds -contains 'App') { 'App' } else { $appIds | Select-Object -First 1 }
            if ($appId) {
                $aumid = "$($pkg.PackageFamilyName)!$appId"
                Write-Host "$($App.Name): AUMID discovered via AppxPackage manifest: $aumid"
                return $aumid
            }
        } catch {
            Write-Warning "$($App.Name): could not read AppxPackage manifest. $_"
        }
    }
    $msg = "$($App.Name): AUMID could not be resolved automatically (all paths exhausted)."
    Write-Warning $msg
    Write-ErrorLog -Message $msg
    return $null
}

function Get-ShortcutObject {
    param([string]$ShortcutPath)
    if (-not (Test-Path -LiteralPath $ShortcutPath)) { throw "Shortcut not found: $ShortcutPath" }
    return $script:WshShell.CreateShortcut($ShortcutPath)
}

function Get-ParentFolder {
    param([string]$BrokenTargetPath)
    if ([string]::IsNullOrWhiteSpace($BrokenTargetPath)) { return $null }
    $targetFolder = Split-Path -Path $BrokenTargetPath -Parent
    $grandparent  = Split-Path -Path $targetFolder -Parent
    if (-not [string]::IsNullOrWhiteSpace($grandparent) -and
        (Test-Path -LiteralPath $grandparent -PathType Container)) {
        return $grandparent
    }
    $parent = $targetFolder
    if ([string]::IsNullOrWhiteSpace($parent) -or
        -not (Test-Path -LiteralPath $parent -PathType Container)) {
        return $null
    }
    return $parent
}

function Get-RelativeDepth {
    param([string]$BasePath, [string]$CandidatePath)
    $baseFull      = [System.IO.Path]::GetFullPath($BasePath)
    $candidateFull = [System.IO.Path]::GetFullPath($CandidatePath)
    if (-not $candidateFull.StartsWith($baseFull, [System.StringComparison]::OrdinalIgnoreCase)) { return [int]::MaxValue }
    $relative = $candidateFull.Substring($baseFull.Length).TrimStart('\\')
    if ([string]::IsNullOrWhiteSpace($relative)) { return 0 }
    return ($relative -split '[\\//]').Count
}

# T-06: depth cap removed; Get-RelativeDepth kept for test coverage only.
function Find-ExeWithinDepth {
    param([string]$RootFolder, [string]$ExpectedExe)
    if (-not (Test-Path -LiteralPath $RootFolder -PathType Container)) { return $null }
    $results = Get-ChildItem -LiteralPath $RootFolder -Filter $ExpectedExe -File -Recurse -ErrorAction SilentlyContinue |
        Sort-Object FullName
    return $results | Select-Object -First 1
}

function Update-ShortcutTarget {
    param([string]$ShortcutPath, [string]$ExePath, [string]$Arguments = "")
    $shortcut = Get-ShortcutObject -ShortcutPath $ShortcutPath
    $shortcut.TargetPath       = $ExePath
    $shortcut.WorkingDirectory = Split-Path -Path $ExePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($Arguments)) {
        $shortcut.Arguments = $Arguments
    }
    $shortcut.Save()
}

function Test-ExePathAllowed {
    param([string]$ExePath)
    $full = [System.IO.Path]::GetFullPath($ExePath)
    foreach ($root in $script:AllowedExeRoots) {
        $rootFull = [System.IO.Path]::GetFullPath($root.TrimEnd('\'))
        if ($full.StartsWith($rootFull + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function Test-ExeSignatureTrusted {
    param(
        [string]$ExePath,
        [string]$ExpectedPublisher = ""
    )
    try {
        $sig = Get-AuthenticodeSignature -FilePath $ExePath -ErrorAction Stop
        if ($sig.Status -ne 'Valid') {
            Write-Warning "Signature status for '$ExePath' is '$($sig.Status)'. Only 'Valid' is accepted."
            return $false
        }
        if (-not [string]::IsNullOrWhiteSpace($ExpectedPublisher)) {
            $subject = $sig.SignerCertificate.Subject
            if ($subject -notlike "*$ExpectedPublisher*") {
                Write-Warning "Publisher mismatch for '$ExePath'. Expected subject to contain '$ExpectedPublisher', got '$subject'."
                return $false
            }
        }
        return $true
    } catch {
        Write-Warning "Could not verify Authenticode signature for '$ExePath'. $_"
        return $false
    }
}

function Prompt-ForExactExePath {
    param(
        [string]$AppName,
        [string]$ExpectedExe,
        [string]$ExpectedPublisher = "",
        [int]$MaxAttempts = 3
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $inputPath = Read-Host "Enter the full path for $AppName ($ExpectedExe), or press Enter to skip (attempt $attempt of $MaxAttempts)"
        if ([string]::IsNullOrWhiteSpace($inputPath)) { return $null }
        $trimmed = $inputPath.Trim('"').Trim()
        if (-not (Test-Path -LiteralPath $trimmed -PathType Leaf)) {
            Write-Warning "Path does not exist or is not a file: $trimmed"; continue
        }
        if ([System.IO.Path]::GetFileName($trimmed) -ine $ExpectedExe) {
            Write-Warning "File name must be exactly $ExpectedExe"; continue
        }
        if (-not (Test-ExePathAllowed -ExePath $trimmed)) {
            Write-Warning "Path is outside allowed roots ($($script:AllowedExeRoots -join ', ')): $trimmed"
            continue
        }
        if (-not (Test-ExeSignatureTrusted -ExePath $trimmed -ExpectedPublisher $ExpectedPublisher)) {
            Write-Warning "File did not pass signature/publisher verification: $trimmed"
            continue
        }
        return $trimmed
    }
    Write-Warning "$AppName: maximum path attempts ($MaxAttempts) reached. Skipping."
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
        Write-Host "$($App.Name): misnumbered shortcut found ('$($misnumbered.Name)'). Renaming to expected name."
        Rename-Item -LiteralPath $misnumbered.FullName -NewName ([System.IO.Path]::GetFileName($App.ShortcutPath))
        return
    }

    Write-Warning "$($App.Name): no shortcut found at '$($App.ShortcutPath)'. Creating..."

    if (-not [string]::IsNullOrWhiteSpace($App.ExpectedArguments)) {
        New-AppShortcut -Path $App.ShortcutPath `
                        -TargetPath "$env:SystemRoot\explorer.exe" `
                        -Arguments $App.ExpectedArguments `
                        -WorkingDirectory $env:SystemRoot
        Write-Host "$($App.Name): shortcut created with Arguments: $($App.ExpectedArguments)"
        return
    }

    $exePath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe -ExpectedPublisher $App.ExpectedPublisher
    if ($exePath) {
        New-AppShortcut -Path $App.ShortcutPath -TargetPath $exePath `
                        -WorkingDirectory (Split-Path -Path $exePath -Parent)
        Write-Host "$($App.Name): shortcut created at '$($App.ShortcutPath)'."
    } else {
        Write-Warning "$($App.Name): shortcut creation skipped by user."
    }
}

function Repair-ShortcutTarget {
    param($App)
    $shortcut   = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
    $targetPath = $shortcut.TargetPath
    Write-Warning "$($App.Name): shortcut target missing or invalid: $targetPath"
    $searchRoot = Get-ParentFolder -BrokenTargetPath $targetPath
    if ($searchRoot) {
        Write-Host "$($App.Name): searching for $($App.ExpectedExe) under $searchRoot (all subfolders)..."
        $foundExe = Find-ExeWithinDepth -RootFolder $searchRoot -ExpectedExe $App.ExpectedExe
        if ($foundExe) {
            if (-not (Test-ExePathAllowed -ExePath $foundExe.FullName)) {
                Write-Warning "$($App.Name): discovered exe is outside allowed roots. Skipping auto-repair: $($foundExe.FullName)"
            } elseif (-not (Test-ExeSignatureTrusted -ExePath $foundExe.FullName -ExpectedPublisher $App.ExpectedPublisher)) {
                Write-Warning "$($App.Name): discovered exe did not pass signature/publisher verification. Skipping auto-repair: $($foundExe.FullName)"
            } else {
                Write-Host "$($App.Name): found replacement at $($foundExe.FullName). Updating shortcut."
                Update-ShortcutTarget -ShortcutPath $App.ShortcutPath -ExePath $foundExe.FullName -Arguments $App.ExpectedArguments
                return $foundExe.FullName
            }
        }
        Write-Warning "$($App.Name): $($App.ExpectedExe) not found under $searchRoot."
    } else {
        Write-Warning "$($App.Name): could not determine an existing parent folder from the broken target."
    }
    $manualPath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe -ExpectedPublisher $App.ExpectedPublisher
    if ($manualPath) {
        Update-ShortcutTarget -ShortcutPath $App.ShortcutPath -ExePath $manualPath -Arguments $App.ExpectedArguments
        return $manualPath
    }
    return $null
}

# INT-01: uses $env:ProgramFiles instead of hardcoded 'C:\Program Files'.
function Repair-ShortcutArguments {
    param($App)
    $shortcut = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
    Write-Warning "$($App.Name): shortcut Arguments missing or invalid: '$($shortcut.Arguments)'"

    $aumidFragment = $null
    if ($App.ExpectedArguments -match '^shell:appsFolder\\([A-Za-z0-9][A-Za-z0-9._]*_[A-Za-z0-9]+)![A-Za-z0-9._-]+$') {
        $fullPfn       = $Matches[1]
        $aumidFragment = ($fullPfn -split '_', 2)[1]
    }

    if ([string]::IsNullOrWhiteSpace($aumidFragment)) {
        Write-Warning "$($App.Name): cannot extract AUMID fragment from ExpectedArguments. Skipping argument repair."
        return $null
    }

    $windowsApps = Join-Path $env:ProgramFiles 'WindowsApps'
    Write-Host "$($App.Name): scanning $windowsApps for '*$aumidFragment*'..."
    $pkgFolder = Get-ChildItem -LiteralPath $windowsApps -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like "*$aumidFragment*" } |
        Sort-Object Name -Descending |
        Select-Object -First 1

    if (-not $pkgFolder) {
        Write-Warning "$($App.Name): no folder matching '*$aumidFragment*' found in WindowsApps."
        return $null
    }

    Write-Host "$($App.Name): found package folder: $($pkgFolder.Name)"

    $installedPkg = Get-AppxPackage | Where-Object { $_.PackageFamilyName -like "*$aumidFragment*" } | Select-Object -First 1
    $pfn = if ($installedPkg) { $installedPkg.PackageFamilyName } else { $null }

    $manifestPath = Join-Path $pkgFolder.FullName "AppxManifest.xml"
    $appId = $null
    if (Test-Path -LiteralPath $manifestPath) {
        try {
            $manifest = [xml]::new()
            $manifest.Load($manifestPath)
            $appIds = $manifest.Package.Applications.Application.Id
            $appId  = if ($appIds -contains 'App') { 'App' } else { $appIds | Select-Object -First 1 }
        } catch {
            Write-Warning "$($App.Name): could not read AppxManifest.xml. $_"
        }
    }

    if ([string]::IsNullOrWhiteSpace($appId)) {
        Write-Warning "$($App.Name): manifest unreadable (likely ACL-blocked). Falling back to ExpectedArguments AppId."
        $appId = ($App.ExpectedArguments -split '!') | Select-Object -Last 1
    }

    if ([string]::IsNullOrWhiteSpace($pfn)) {
        Write-Warning "$($App.Name): installed package not found via Get-AppxPackage for fragment '$aumidFragment'. Cannot repair arguments."
        return $null
    }

    $aumid        = "$pfn!$appId"
    $repairedArgs = "shell:appsFolder\$aumid"

    Write-Host "$($App.Name): reconstructed AUMID: $aumid. Updating shortcut Arguments."
    $shortcut.Arguments        = $repairedArgs
    $shortcut.WorkingDirectory = Split-Path -Path $shortcut.TargetPath -Parent
    $shortcut.Save()
    return $repairedArgs
}

# ---------------------------------------------------------------------------
# Inline failure menu
# ---------------------------------------------------------------------------
function Show-FailureMenu {
    param([string]$AppName, [string]$Context)
    Write-Host "  [1] Add / fix shortcut for $AppName and retry"
    Write-Host "  [2] Modify a different shortcut"
    Write-Host "  [3] Skip"
    return (Read-Host "Select ($Context)")
}

# ---------------------------------------------------------------------------
# Shared launch-wait tail
# ---------------------------------------------------------------------------
function Invoke-AppLaunchWait {
    param($App, [int]$TimeoutSeconds = $script:LaunchTimeoutSeconds)
    if (Wait-ForAppReady -ProcessName $App.ProcessName -TimeoutSeconds $TimeoutSeconds) {
        Write-Host "$($App.Name): ready.`n"
        Start-Sleep -Seconds $script:PostLaunchPauseSeconds
        return $true
    }
    Write-Warning "$($App.Name): did not become ready within $TimeoutSeconds seconds."
    return $false
}

# ---------------------------------------------------------------------------
# Invoke-FailureRecovery — centralises Add+retry / Modify / Skip
# ---------------------------------------------------------------------------
function Invoke-FailureRecovery {
    param(
        $App,
        [string]$Context,
        [scriptblock]$PreRetryAction = $null
    )
    $choice = Show-FailureMenu -AppName $App.Name -Context $Context
    switch ($choice) {
        '1' {
            if ($PreRetryAction) { & $PreRetryAction }
            return $true
        }
        '2' {
            $target = Show-AppPicker -Prompt "Select shortcut to modify:"
            if ($target) { Edit-Shortcut -App $target }
            return $false
        }
        default {
            Write-Host "$($App.Name): skipped.`n"
            return $false
        }
    }
}

# ---------------------------------------------------------------------------
# Launch functions
# ---------------------------------------------------------------------------
function Start-AppxApp {
    param($App)
    if (Test-AppAlreadyOpen -ProcessName $App.ProcessName) {
        Write-Host "$($App.Name): already open. Skipping.`n"
        return $true
    }
    $aumid = Resolve-Aumid -App $App
    if (-not $aumid) {
        Write-Warning "$($App.Name): no AUMID found. Skipping.`n"
        return $false
    }
    try {
        Write-Host "$($App.Name): launching via shell:appsFolder\$aumid"
        Start-Process explorer.exe "shell:appsFolder\$aumid" -ErrorAction Stop
        return Invoke-AppLaunchWait -App $App
    } catch {
        Write-Warning "$($App.Name): launch failed. $_`n"
        return $false
    }
}

# INT-02: $shortcut is re-read via Get-ShortcutObject after Repair-ShortcutTarget
#         so the argument-repair check works on the freshly written .lnk.
function Start-Win32App {
    param($App)
    $requireWin = (Get-AppPresenceMode -ProcessName $App.ProcessName -SettleSecs 0) -eq 'Window'
    if (Test-AppAlreadyOpen -ProcessName $App.ProcessName -ExpectedExe $App.ExpectedExe `
            -RequireWindow:$requireWin) {
        Write-Host "$($App.Name): already open. Skipping.`n"
        return $true
    }

    for ($attempt = 0; $attempt -le 2; $attempt++) {

        if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
            Write-Warning "$($App.Name): shortcut file not found: $($App.ShortcutPath)"
            $recover = Invoke-FailureRecovery -App $App -Context "missing shortcut" `
                -PreRetryAction { Initialize-Shortcut -App $App }
            if (-not $recover) { return $false }
            continue
        }

        try {
            $shortcut   = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
            $targetPath = $shortcut.TargetPath
            if ([string]::IsNullOrWhiteSpace($targetPath) -or -not (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
                $repairedPath = Repair-ShortcutTarget -App $App
                if (-not $repairedPath) {
                    Write-Warning "$($App.Name): shortcut could not be repaired. Skipping.`n"
                    return $false
                }
                Write-Host "$($App.Name): shortcut repaired. Proceeding with launch."
                # INT-02: re-read after repair so $shortcut reflects the updated .lnk.
                $shortcut = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
            }

            if (-not [string]::IsNullOrWhiteSpace($App.ExpectedArguments)) {
                $currentArgs = $shortcut.Arguments
                if ([string]::IsNullOrWhiteSpace($currentArgs) -or
                    $currentArgs -notlike "*$($App.ExpectedArguments)*") {
                    $repairedArgs = Repair-ShortcutArguments -App $App
                    if (-not $repairedArgs) {
                        Write-Warning "$($App.Name): argument repair failed. Skipping.`n"
                        return $false
                    }
                }
            }

            Write-Host "$($App.Name): launching via shortcut: $($App.ShortcutPath)"
            $script:WshShell.Run("`"$($App.ShortcutPath)`"", 1, $false)

            if (Invoke-AppLaunchWait -App $App) { return $true }

            $recover = Invoke-FailureRecovery -App $App -Context "launch timeout" `
                -PreRetryAction { Edit-Shortcut -App $App }
            if (-not $recover) { return $false }

        } catch {
            Write-Warning "$($App.Name): launch failed. $_`n"
            Write-ErrorLog -Message "$($App.Name): launch exception" -ErrorRecord $_
            return $false
        }
    }

    Write-Warning "$($App.Name): max retries reached. Skipping.`n"
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
Write-Host "  [6] Exit"
Write-Host "------------------------------------------------"
$mainChoice = Read-Host "Select"

switch ($mainChoice) {
    '2' {
        $app = Show-AppPicker -Prompt "Select app to re-initialise, or [N] to add a new entry:" -AllowNew
        Add-Shortcut -App $app
        exit
    }
    '3' {
        $app = Show-AppPicker -Prompt "Select app to DELETE shortcut for:"
        if ($app) { Remove-Shortcut -App $app }
        exit
    }
    '4' {
        $app = Show-AppPicker -Prompt "Select app to MODIFY shortcut for:"
        if ($app) { Edit-Shortcut -App $app }
        exit
    }
    '5' {
        Show-AppList
        exit
    }
    '6' { exit }
}

Write-Host "Waiting $script:InitialDelaySeconds seconds for system to stabilize..."
Start-Sleep -Seconds $script:InitialDelaySeconds

Write-Host "`n--- Shortcut bootstrap ---"
foreach ($app in $script:apps) {
    if ($app.LaunchType -eq "Win32") {
        Initialize-Shortcut -App $app
    }
}
Write-Host "--- Bootstrap complete ---`n"

$failedApps = @()
foreach ($app in $script:apps) {
    $ok = if ($app.LaunchType -eq "Appx") {
        Start-AppxApp -App $app
    } else {
        Start-Win32App -App $app
    }
    if (-not $ok) { $failedApps += "$($app.Name) [$($app.ProcessName)]" }
}

if ($failedApps.Count -gt 0) {
    Write-Host "`n--- Startup completed with failures ---"
    foreach ($entry in $failedApps) {
        Write-Host "  - $entry"
    }
} else {
    Write-Host "`nStartup sequence completed successfully."
}
