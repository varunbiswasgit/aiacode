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
# - Main menu      : on launch, user chooses Run / Add / Delete / Modify / Exit
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

$script:startMenu              = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
$script:InitialDelaySeconds    = 10
$script:LaunchTimeoutSeconds   = 30
$script:PostLaunchPauseSeconds = 2
$script:SettleSeconds          = 5
$script:AppsConfigPath         = Join-Path $PSScriptRoot "apps.json"

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
    $entries = $raw | ConvertFrom-Json
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
    $script:apps | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Path -Encoding UTF8
    Write-Host "apps.json saved ($($script:apps.Count) entries)."
}

try {
    $script:apps = Import-AppsConfig
} catch {
    Write-Error $_
    exit 1
}

$script:WshShell = New-Object -ComObject WScript.Shell

# ---------------------------------------------------------------------------
# Presence mode detection
# ---------------------------------------------------------------------------
function Get-AppPresenceMode {
    param([string]$ProcessName, [int]$SettleSecs = $script:SettleSeconds)
    $elapsed = 0
    while ($elapsed -lt $SettleSecs) {
        $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) {
            return 'Window'
        }
        Start-Sleep -Seconds 1
        $elapsed++
    }
    if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) {
        return 'Tray'
    }
    return $null
}

function Test-AppAlreadyOpen {
    param(
        [string]$ProcessName,
        [string]$ExpectedExe = ""
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

    if ($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) { return $true }
    return $true
}

function Wait-ForAppReady {
    param([string]$ProcessName, [int]$TimeoutSeconds = $script:LaunchTimeoutSeconds)
    $phase1Secs = [Math]::Min($script:SettleSeconds, $TimeoutSeconds)
    $mode = Get-AppPresenceMode -ProcessName $ProcessName -SettleSecs $phase1Secs
    if ($null -eq $mode) {
        $remaining = $TimeoutSeconds - $phase1Secs
        $elapsed   = 0
        while ($elapsed -lt $remaining) {
            if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) { return $true }
            Start-Sleep -Seconds 1
            $elapsed++
        }
        return $false
    }

    Write-Host "  (presence mode: $mode)"

    if ($mode -eq 'Tray') { return $true }

    $remaining = $TimeoutSeconds - $phase1Secs
    $elapsed   = 0
    while ($elapsed -lt $remaining) {
        $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) { return $true }
        Start-Sleep -Seconds 1
        $elapsed++
    }
    return $false
}

# ---------------------------------------------------------------------------
# Helper: numbered app picker
# ---------------------------------------------------------------------------
function Show-AppPicker {
    param([string]$Prompt, [switch]$AllowNew)
    Write-Host "`n$Prompt"
    for ($i = 0; $i -lt $script:apps.Count; $i++) {
        $exists = if (Test-Path -LiteralPath $script:apps[$i].ShortcutPath -PathType Leaf) { "exists" } else { "missing" }
        Write-Host ("  [{0}] {1,-20}  {2}  ({3})" -f ($i + 1), $script:apps[$i].Name, $exists, $script:apps[$i].ShortcutPath)
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
# Single shortcut writer — all .lnk creation goes through here
# ---------------------------------------------------------------------------
function New-AppShortcut {
    # -Arguments and -WorkingDirectory are optional; omit for exe-only shortcuts.
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

    # ── Brand-new entry flow ──────────────────────────────────────────────────
    Write-Host "`n--- Add new app entry ---"
    $name = Read-Host "App name (display label)"
    if ([string]::IsNullOrWhiteSpace($name)) { Write-Host "Cancelled."; return }

    $launchType = ''
    while ($launchType -notin @('Win32','Appx')) {
        $launchType = Read-Host "Launch type [Win32 / Appx]"
    }

    $number       = Read-Host "Shortcut number (e.g. 09)"
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
        $script:apps = $script:apps | Where-Object { $_.Name -ne $App.Name }
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
    if (-not [string]::IsNullOrWhiteSpace($App.KnownAumid)) {
        $knownPfn = ($App.KnownAumid -split '!')[0]
        $installed = Get-AppxPackage | Where-Object { $_.PackageFamilyName -eq $knownPfn } | Select-Object -First 1
        if ($installed) {
            Write-Host "$($App.Name): KnownAumid verified as installed: $($App.KnownAumid)"
            return $App.KnownAumid
        }
        Write-Warning "$($App.Name): KnownAumid package family '$knownPfn' not found on this system."
    }
    $pkg = Get-AppxPackage | Where-Object { $_.Name -like "*$($App.AppxName)*" } | Select-Object -First 1
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
    Write-Warning "$($App.Name): AUMID could not be resolved automatically."
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
    $parent       = Split-Path -Path $targetFolder -Parent
    if ([string]::IsNullOrWhiteSpace($parent) -or -not (Test-Path -LiteralPath $parent -PathType Container)) {
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

function Find-ExeWithinDepth {
    param([string]$RootFolder, [string]$ExpectedExe, [int]$MaxDepth = 3)
    if (-not (Test-Path -LiteralPath $RootFolder -PathType Container)) { return $null }
    $results = Get-ChildItem -LiteralPath $RootFolder -Filter $ExpectedExe -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { (Get-RelativeDepth -BasePath $RootFolder -CandidatePath $_.FullName) -le $MaxDepth } |
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
        [string]$ExpectedPublisher = ""
    )
    while ($true) {
        $inputPath = Read-Host "Enter the full path for $AppName ($ExpectedExe), or press Enter to skip"
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
        New-AppShortcut -Path $App.ShortcutPath -TargetPath "C:\Windows\explorer.exe" `
                        -Arguments $App.ExpectedArguments -WorkingDirectory "C:\Windows"
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
        Write-Host "$($App.Name): searching for $($App.ExpectedExe) under $searchRoot (1 level up, all subfolders)..."
        $foundExe = Find-ExeWithinDepth -RootFolder $searchRoot -ExpectedExe $App.ExpectedExe -MaxDepth 10
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

    $windowsApps = "C:\Program Files\WindowsApps"
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

    $pfn          = $pkgFolder.Name -replace '_\d+\.\d+\.\d+\.\d+_[^_]+__', '_'
    $aumid        = "$pfn!$appId"
    $repairedArgs = "shell:appsFolder\$aumid"

    Write-Host "$($App.Name): reconstructed AUMID: $aumid. Updating shortcut Arguments."
    $shortcut.Arguments        = $repairedArgs
    $shortcut.WorkingDirectory = Split-Path -Path $shortcut.TargetPath -Parent
    $shortcut.Save()
    return $repairedArgs
}

# ---------------------------------------------------------------------------
# Inline failure menu (shared by missing-shortcut and launch-timeout paths)
# ---------------------------------------------------------------------------
function Show-FailureMenu {
    param([string]$AppName, [string]$Context)
    Write-Host "  [1] Add / fix shortcut for $AppName and retry"
    Write-Host "  [2] Modify a different shortcut"
    Write-Host "  [3] Skip"
    return (Read-Host "Select ($Context)")
}

# ---------------------------------------------------------------------------
# Shared launch-wait tail (used by Start-AppxApp and Start-Win32App)
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

function Start-Win32App {
    param($App)
    if (Test-AppAlreadyOpen -ProcessName $App.ProcessName -ExpectedExe $App.ExpectedExe) {
        Write-Host "$($App.Name): already open. Skipping.`n"
        return $true
    }

    if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
        Write-Warning "$($App.Name): shortcut file not found: $($App.ShortcutPath)"
        $failChoice = Show-FailureMenu -AppName $App.Name -Context "missing shortcut"
        switch ($failChoice) {
            '1' {
                Initialize-Shortcut -App $App
                if (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf) {
                    return Start-Win32App -App $App
                }
                Write-Warning "$($App.Name): shortcut still missing after add attempt. Skipping.`n"
                return $false
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

        $timeoutChoice = Show-FailureMenu -AppName $App.Name -Context "launch timeout"
        switch ($timeoutChoice) {
            '1' {
                Edit-Shortcut -App $App
                return Start-Win32App -App $App
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
    } catch {
        Write-Warning "$($App.Name): launch failed. $_`n"
        return $false
    }
}

# ---------------------------------------------------------------------------
# Main menu + startup sequence
# Guard: when PS_STARTUP_TESTMODE=1, functions and vars are available for
# Pester dot-source but the interactive menu and sequence are skipped.
# ---------------------------------------------------------------------------
if ($env:PS_STARTUP_TESTMODE -eq '1') { return }

Write-Host "`n================================================"
Write-Host "  Win11 Startup Manager"
Write-Host "================================================"
Write-Host "  [1] Run startup sequence"
Write-Host "  [2] Add shortcut"
Write-Host "  [3] Delete shortcut"
Write-Host "  [4] Modify shortcut"
Write-Host "  [5] Exit"
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
    '5' { exit }
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
    if (-not $ok) { $failedApps += $app }
}

if ($failedApps.Count -gt 0) {
    Write-Host "`n--- Startup completed with failures ---"
    foreach ($app in $failedApps) {
        Write-Host "  - $($app.Name) [$($app.ProcessName)]"
    }
} else {
    Write-Host "`nStartup sequence completed successfully."
}
