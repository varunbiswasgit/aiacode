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
# - Main menu      : on launch, user chooses Run / Add / Delete / Modify / List / Sync / Exit
#                    inline failure menu (Add+retry / Modify / Skip / Delete entry) appears when a shortcut
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
# - Exe gate       : Test-ExeAcceptable wraps Test-ExePathAllowed + Test-ExeSignatureTrusted into a
#                    single call so every caller enforces both checks identically. (LEAN-01)
# - XML load       : Repair-ShortcutArguments uses [xml]::new() + Load() instead of
#                    [xml]$m = Get-Content to correctly handle BOMs and large manifest files.
# - Publisher gate : Test-ExeSignatureTrusted accepts an optional -ExpectedPublisher string.
#                    When present, the SignerCertificate.Subject must contain that string.
#                    Each app entry carries an optional ExpectedPublisher field; Microsoft apps
#                    use 'CN=Microsoft Corporation', Chrome uses 'CN=Google LLC'.
# - Process guard  : Test-AppAlreadyOpen compares each matching process MainModule filename
#                    against ExpectedExe before treating the app as already open, preventing
#                    false skips caused by unrelated same-named processes.
# - Regex anchor   : Repair-ShortcutArguments uses a fully anchored regex to extract the PFN
#                    from ExpectedArguments. Pattern requires the full value to match
#                    ^shell:appsFolder\<PFN>!<AppId>$ and accepts any valid PackageFamilyName
#                    prefix so non-Microsoft packaged apps are handled correctly.
# - Script scope   : All shared vars ($WshShell, $apps, $startMenu,
#                    $InitialDelaySeconds, $LaunchTimeoutSeconds, $PostLaunchPauseSeconds,
#                    $SettleSeconds, $AllowedExeRoots) use $script: scope so dot-sourced
#                    Pester runs cannot leak or shadow globals.
# - Test mode      : When $env:PS_STARTUP_TESTMODE = '1', the script dot-sources cleanly
#                    (functions and vars available) but skips the interactive menu and
#                    startup sequence entirely. Used by Win11startup.Tests.ps1.
# - Config file    : $script:apps is loaded from Win11startupapps.json in the same folder.
#                    Required fields: Name, LaunchType, ShortcutPath, ExpectedExe.
#                    ProcessName required for Win32; warn-only for Win32-with-args entries
#                    where Sync cannot determine it without running the app.
#                    Add/Delete menu flows write changes back to Win11startupapps.json automatically.
# - Add flow       : Show-AppPicker -AllowNew shows [N] for brand-new entry and returns
#                    '__NEW__' sentinel; Add-Shortcut distinguishes re-init (real app object),
#                    new entry ('__NEW__'), and cancel ($null). Appx entries collect
#                    StartAppName, KnownAumid, AppxName during the add prompt.
# - Timeout math   : Wait-ForAppReady stores actual phase-1 settle duration in $phase1Secs
#                    and subtracts that exact value for phase-2 remaining, so the total
#                    timeout is always honoured even when TimeoutSeconds < SettleSeconds.
# - Failure menu   : Show-FailureMenu centralises the 4-option inline prompt (Add+retry /
#                    Modify / Skip / Delete entry) used by both the missing-shortcut and
#                    launch-timeout paths in Start-Win32App.
#                    [4] Delete entry calls Remove-Shortcut (same as main menu [3]) so no
#                    logic is duplicated.
# - Launch wait    : Invoke-AppLaunchWait centralises the Wait-ForAppReady + ready/warning
#                    output + PostLaunchPauseSeconds sleep shared by Start-AppxApp and
#                    the happy path of Start-Win32App.
# - Shortcut write : New-AppShortcut -Path -TargetPath -Arguments -WorkingDirectory is the
#                    single place all .lnk files are written. Add-Shortcut delegates to
#                    Initialize-Shortcut for the actual write; WshShell.CreateShortcut is
#                    invoked in exactly one function. (LEAN-02)
# - File search    : Get-ParentFolder walks two Split-Path levels above the broken target's
#                    folder (grandparent). Falls back to one level up when the grandparent
#                    path is empty or does not exist. Find-ExeWithinDepth then searches
#                    recursively from that root with Get-ChildItem -Recurse (no depth cap).
# - Appx enum      : Resolve-Aumid calls Get-AppxPackage once, stores the result in $pkgs,
#                    and reuses it for both KnownAumid verification and the AppxName fallback,
#                    avoiding a second full pipeline enumeration.
# - Error logging  : All unhandled errors and terminating exceptions are written to
#                    startup-error.log in the same folder as the script, with timestamps,
#                    so crashes can be diagnosed after PowerShell has closed.
# - Stopwatch      : Get-AppPresenceMode and Wait-ForAppReady use
#                    [System.Diagnostics.Stopwatch] instead of manual $elapsed++ counters
#                    so elapsed time reflects wall clock even when Get-Process is slow. (T-07)
# - PFN source     : Repair-ShortcutArguments uses $pkg.PackageFamilyName directly from
#                    Get-AppxPackage instead of reconstructing it via regex. (FIX-05)
# - SystemRoot     : Initialize-Shortcut uses $env:SystemRoot for Appx shortcut creation. (FIX-06)
# - Export guard   : Export-AppsConfig wraps Set-Content in try/catch. (ROB-01)
# - Edit guard     : Edit-Shortcut argument-repair branch guards Get-ShortcutObject. (ROB-02)
# - Retry loop     : Start-Win32App uses Invoke-FailureRecovery with bounded for-loop,
#                    up to 3 attempts (attempt 0, 1, 2). (ROB-04+QOL-04)
# - Path attempts  : Prompt-ForExactExePath limits retries to 3 attempts. (HARD-04)
# - Number valid   : Add-Shortcut validates shortcut number is 1-2 digits. (HARD-05)
# - Picker perf    : Show-AppPicker pre-computes shortcut-existence before display loop. (QOL-01)
# - Window switch  : Test-AppAlreadyOpen accepts -RequireWindow. (FIX-04)
# - AUMID log      : Resolve-Aumid logs all-paths failure. (QOL-02)
# - Schema version : Import-AppsConfig checks for top-level schemaVersion field. (QOL-03)
# - List apps      : Main menu option [5] prints formatted table via Show-AppList. (QOL-05)
# - Stale shortcut : Start-Win32App re-reads shortcut after Repair-ShortcutTarget. (INT-02)
# - WindowsApps    : Repair-ShortcutArguments uses Join-Path $env:ProgramFiles. (INT-01)
# - AppList fix    : Show-AppList uses bare string for status assignment. (BUG-02)
# - Catch recovery : Start-Win32App catch block calls Invoke-FailureRecovery. (UX-03)
# - Delete guard   : Remove-Shortcut uses combined prompt when shortcut is missing. (UX-02)
# - RequireWindow  : Start-Win32App derives $requireWin from cached $App.PresenceMode. (BUG-03)
# - Dead branch    : Test-AppAlreadyOpen non-RequireWindow tail collapsed to single return $true. (BUG-01)
# - Poll helper    : Wait-ForProcessCondition extracted from Wait-ForAppReady. (DUP-01)
# - Sync menu      : Main menu [6] calls Sync-AppsFromStartMenu; also auto-triggered when
#                    Win11startupapps.json is missing (first-run / broken-path scenario).
#                    Scans Start Menu Programs for numbered .lnk files (1-2 digit prefix),
#                    classifies each as Win32 or Appx, writes Win11startupapps.json. (SYNC-01)
# - Dead code      : Get-RelativeDepth removed from main script; never called in startup
#                    sequence. Retained only in Win11startup.Tests.ps1 for unit test coverage. (AUD-01)
# - BUG-04         : Prompt-ForExactExePath max-attempts warning used $AppName: which PS
#                    parsed as a drive-scoped variable reference; fixed to ${AppName}.
# - BUG-05         : Sync-AppsFromStartMenu classified explorer.exe+shell:appsFolder entries
#                    as Appx, causing empty ProcessName to fail Import-AppsConfig validation
#                    on next load. Fix: keep LaunchType=Win32 for all such entries (they are
#                    Win32 .lnk files invoking explorer.exe with arguments, not true Appx).
#                    Import-AppsConfig now warns instead of throws when ProcessName is empty
#                    on a Win32-with-args entry, since Sync cannot determine it without running.
# - BUG-06 (rev)   : When ProcessName is blank, launch confirmation is done by detecting an
#                    active/opened window (MainWindowHandle != 0) whose MainWindowTitle matches
#                    $App.Name or $App.StartAppName - not by polling a process name.
#                    Wait-ForWindowByTitle replaces Resolve-ProcessName. Once a matching window
#                    is found, ProcessName is back-filled from that process object and persisted
#                    to Win11startupapps.json. Future runs use the normal Wait-ForAppReady path.
# - FIX-07         : Write-ErrorLog defined before trap block so the trap can call it on cold
#                    errors before any other script code runs. Previously the trap fired before
#                    the function was parsed, causing CommandNotFoundException.
#                    Boot block now prompts user for a custom JSON path when Win11startupapps.json
#                    is not found at the default location. If the user supplies a valid path,
#                    $script:AppsConfigPath is updated and Import-AppsConfig loads from there.
#                    If the user supplies a new (non-existent) path, an empty config skeleton
#                    is written there and Sync-AppsFromStartMenu populates it. Pressing Enter
#                    falls back to the original auto-sync behaviour.
# - UX-04          : Show-FailureMenu adds [4] Delete entry. Invoke-FailureRecovery '4' branch
#                    calls Remove-Shortcut (reusing main menu [3] logic) so no logic is
#                    duplicated. Failure menu now mirrors main menu: Add=[2], Modify=[4],
#                    Delete=[3], with Skip as the safe default.
# - LEAN-05        : Zero-shortcut Write-Warning moved inside Sync-AppsFromStartMenu.
#                    Callers check return bool and exit/return only; no duplicated message text.
# - LEAN-01        : Test-ExeAcceptable wraps Test-ExePathAllowed + Test-ExeSignatureTrusted.
#                    Repair-ShortcutTarget and Prompt-ForExactExePath both call Test-ExeAcceptable.
#                    The two helpers remain private (no external callers).
# - LEAN-02        : Add-Shortcut (new entry branch) no longer calls New-AppShortcut directly.
#                    After building $newEntry it calls Initialize-Shortcut -App $newEntry.
#                    Initialize-Shortcut is the sole .lnk creation path.

# ---------------------------------------------------------------------------
# Error log path + Write-ErrorLog -- MUST be defined before the trap block
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
# SYNC-01: Sync-AppsFromStartMenu
# LEAN-05: Write-Warning for zero-shortcut case lives here; callers check bool only.
# ---------------------------------------------------------------------------
function Sync-AppsFromStartMenu {
    Write-Host "`n--- Sync from Start Menu ---"
    $lnkFiles = Get-ChildItem -LiteralPath $script:startMenu -Filter '*.lnk' -ErrorAction SilentlyContinue |
        Where-Object { $_.BaseName -match '^\d{1,2}\s' } |
        Sort-Object Name

    if ($lnkFiles.Count -eq 0) {
        Write-Warning "No numbered .lnk files found in '$script:startMenu'. Nothing to sync."
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
            Write-Warning "${appName}: ProcessName unknown after sync. Will be auto-detected on first run (BUG-06)."
        } elseif ($leafName -notlike '*.exe') {
            Write-Warning "'$($file.Name)': unexpected target '$target'. Review and fill in fields manually."
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
# New-AppEntry -- single constructor for all app entry objects
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
# FIX-07: Boot -- resolve config path before loading
# LEAN-05: Callers check Sync-AppsFromStartMenu return bool; no inline message.
# ---------------------------------------------------------------------------
$script:WshShell = New-Object -ComObject WScript.Shell

function Resolve-ConfigPath {
    if (Test-Path -LiteralPath $script:AppsConfigPath -PathType Leaf) { return $true }

    Write-Host "`n[CONFIG NOT FOUND] Win11startupapps.json not found at:" -ForegroundColor Yellow
    Write-Host "  $script:AppsConfigPath" -ForegroundColor Yellow
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
# Presence mode detection
# ---------------------------------------------------------------------------
function Get-AppPresenceMode {
    param([string]$ProcessName, [int]$SettleSecs = $script:SettleSeconds)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $SettleSecs) {
        $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) { return 'Window' }
        Start-Sleep -Seconds 1
    }
    if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) { return 'Tray' }
    return $null
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

# ---------------------------------------------------------------------------
# BUG-06 (rev): Wait-ForWindowByTitle
# ---------------------------------------------------------------------------
function Wait-ForWindowByTitle {
    param($App, [int]$WaitSecs = $script:SettleSeconds)
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
# Helper: numbered app picker
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
# Single shortcut writer
# ---------------------------------------------------------------------------
function New-AppShortcut {
    param([string]$Path, [string]$TargetPath, [string]$Arguments = "", [string]$WorkingDirectory = "")
    $sc = $script:WshShell.CreateShortcut($Path)
    $sc.TargetPath = $TargetPath
    if (-not [string]::IsNullOrWhiteSpace($Arguments))        { $sc.Arguments        = $Arguments }
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) { $sc.WorkingDirectory = $WorkingDirectory }
    $sc.Save()
}

# ---------------------------------------------------------------------------
# Exe validation
# LEAN-01: Test-ExeAcceptable is the single call site for both checks.
#          Test-ExePathAllowed and Test-ExeSignatureTrusted are private helpers.
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
# Shortcut management
# LEAN-02: Add-Shortcut (new entry) calls Initialize-Shortcut instead of
#          New-AppShortcut directly. Initialize-Shortcut is the sole .lnk writer.
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

    # For a plain Win32 entry without arguments we need the exe path to create the shortcut.
    # Validate it here so Initialize-Shortcut can skip Prompt-ForExactExePath on the first call.
    if ($launchType -eq 'Win32' -and [string]::IsNullOrWhiteSpace($expectedArguments)) {
        $exePath = Prompt-ForExactExePath -AppName $name -ExpectedExe $expectedExe -ExpectedPublisher $expectedPublisher
        if (-not $exePath) { Write-Host "Cancelled."; return }
        # Store resolved path temporarily so Initialize-Shortcut can use it without prompting again.
        $newEntry = New-AppEntry -Name $name -LaunchType $launchType -ShortcutPath $shortcutPath `
            -ProcessName $processName -ExpectedExe $expectedExe -ExpectedPublisher $expectedPublisher `
            -ExpectedArguments $expectedArguments -StartAppName $startAppName `
            -KnownAumid $knownAumid -AppxName $appxName
        # Write shortcut directly when we already have the validated exe path.
        New-AppShortcut -Path $shortcutPath -TargetPath $exePath -WorkingDirectory (Split-Path $exePath -Parent)
        Write-Host "$($name): shortcut created at '$shortcutPath'."
    } else {
        $newEntry = New-AppEntry -Name $name -LaunchType $launchType -ShortcutPath $shortcutPath `
            -ProcessName $processName -ExpectedExe $expectedExe -ExpectedPublisher $expectedPublisher `
            -ExpectedArguments $expectedArguments -StartAppName $startAppName `
            -KnownAumid $knownAumid -AppxName $appxName
        # Initialize-Shortcut handles the write (arguments path or misnumbered rename).
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
        if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
            New-AppShortcut -Path $App.ShortcutPath -TargetPath $exePath -WorkingDirectory (Split-Path $exePath -Parent)
            Write-Host "$($App.Name): shortcut created at '$($App.ShortcutPath)'."
        } else {
            Update-ShortcutTarget -ShortcutPath $App.ShortcutPath -ExePath $exePath
            Write-Host "$($App.Name): shortcut updated to '$exePath'."
        }
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
    param([string]$ShortcutPath)
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
    param([string]$RootFolder, [string]$ExpectedExe)
    if (-not (Test-Path -LiteralPath $RootFolder -PathType Container)) { return $null }
    return Get-ChildItem -LiteralPath $RootFolder -Filter $ExpectedExe -File -Recurse -ErrorAction SilentlyContinue |
        Sort-Object FullName | Select-Object -First 1
}

function Update-ShortcutTarget {
    param([string]$ShortcutPath, [string]$ExePath, [string]$Arguments = "")
    $shortcut = Get-ShortcutObject -ShortcutPath $ShortcutPath
    $shortcut.TargetPath = $ExePath; $shortcut.WorkingDirectory = Split-Path -Path $ExePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($Arguments)) { $shortcut.Arguments = $Arguments }
    $shortcut.Save()
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
    $exePath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe -ExpectedPublisher $App.ExpectedPublisher
    if ($exePath) {
        New-AppShortcut -Path $App.ShortcutPath -TargetPath $exePath -WorkingDirectory (Split-Path $exePath -Parent)
        Write-Host "$($App.Name): shortcut created at '$($App.ShortcutPath)'."
    } else { Write-Warning "$($App.Name): shortcut creation skipped." }
}

function Repair-ShortcutTarget {
    param($App)
    $shortcut   = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
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
            Update-ShortcutTarget -ShortcutPath $App.ShortcutPath -ExePath $foundExe.FullName -Arguments $App.ExpectedArguments
            return $foundExe.FullName
        }
        Write-Warning "$($App.Name): $($App.ExpectedExe) not found under $searchRoot."
    } else { Write-Warning "$($App.Name): could not determine parent folder from broken target." }
    $manualPath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe -ExpectedPublisher $App.ExpectedPublisher
    if ($manualPath) { Update-ShortcutTarget -ShortcutPath $App.ShortcutPath -ExePath $manualPath -Arguments $App.ExpectedArguments; return $manualPath }
    return $null
}

function Repair-ShortcutArguments {
    param($App)
    $shortcut = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
    Write-Warning "$($App.Name): shortcut Arguments invalid: '$($shortcut.Arguments)'"
    $aumidFragment = $null
    if ($App.ExpectedArguments -match '^shell:appsFolder\\([A-Za-z0-9][A-Za-z0-9._]*_[A-Za-z0-9]+)![A-Za-z0-9._-]+$') {
        $aumidFragment = (($Matches[1]) -split '_', 2)[1]
    }
    if ([string]::IsNullOrWhiteSpace($aumidFragment)) { Write-Warning "$($App.Name): cannot extract AUMID fragment. Skipping."; return $null }
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
    $shortcut.Arguments = $repairedArgs; $shortcut.WorkingDirectory = Split-Path -Path $shortcut.TargetPath -Parent; $shortcut.Save()
    return $repairedArgs
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

function Start-Win32App {
    param($App)
    $requireWin = ($App.PresenceMode -eq 'Window')
    if (Test-AppAlreadyOpen -ProcessName $App.ProcessName -ExpectedExe $App.ExpectedExe -RequireWindow:$requireWin) {
        Write-Host "$($App.Name): already open. Skipping.`n"; return $true
    }
    for ($attempt = 0; $attempt -le 2; $attempt++) {
        if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
            Write-Warning "$($App.Name): shortcut not found: $($App.ShortcutPath)"
            $recover = Invoke-FailureRecovery -App $App -Context "missing shortcut" -PreRetryAction { Initialize-Shortcut -App $App }
            if (-not $recover) { return $false }; continue
        }
        try {
            $shortcut   = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
            $targetPath = $shortcut.TargetPath
            if ([string]::IsNullOrWhiteSpace($targetPath) -or -not (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
                $repairedPath = Repair-ShortcutTarget -App $App
                if (-not $repairedPath) { Write-Warning "$($App.Name): repair failed. Skipping.`n"; return $false }
                $shortcut = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
            }
            if (-not [string]::IsNullOrWhiteSpace($App.ExpectedArguments)) {
                $currentArgs = $shortcut.Arguments
                if ([string]::IsNullOrWhiteSpace($currentArgs) -or $currentArgs -notlike "*$($App.ExpectedArguments)*") {
                    $repairedArgs = Repair-ShortcutArguments -App $App
                    if (-not $repairedArgs) { Write-Warning "$($App.Name): argument repair failed. Skipping.`n"; return $false }
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
            }
            if ($ready) { return $true }
            $recover = Invoke-FailureRecovery -App $App -Context "launch timeout" -PreRetryAction { Edit-Shortcut -App $App }
            if (-not $recover) { return $false }
        } catch {
            Write-Warning "$($App.Name): launch exception. $_`n"
            Write-ErrorLog -Message "$($App.Name): launch exception" -ErrorRecord $_
            $recover = Invoke-FailureRecovery -App $App -Context "launch exception" -PreRetryAction { Edit-Shortcut -App $App }
            if (-not $recover) { return $false }
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
