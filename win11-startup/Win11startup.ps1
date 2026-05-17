# Curated startup launcher with self-healing shortcut repair
# - Win32 apps     : shortcut invoked via WshShell.Run when target is valid (preserves baked-in arguments)
#                    self-healing repair + user-prompt fallback used only when shortcut target is broken
#                    repaired shortcuts are then invoked via WshShell.Run as well
# - Appx apps      : AUMID resolved at runtime (Get-StartApps -> KnownAumid verification -> AppxPackage manifest)
#                    KnownAumid used only as primary candidate, not sole source of truth
# - Sticky Notes   : Win32 shortcut with /memoryWindow start baked into the .lnk Target field
#                    WshShell.Run fires the shortcut as-is; no separate Arguments field needed
# - Phone Link     : Win32 shortcut targeting explorer.exe with shell:appsFolder AUMID as Arguments
#                    argument self-healing scans WindowsApps for the package family name fragment
#                    and reconstructs the AUMID from the installed folder name + AppxManifest.xml
# - Bootstrap      : before launch loop, ensures every Win32 .lnk exists at the expected path;
#                    renames misnumbered matches found in the same folder, or creates fresh if absent
# - Main menu      : on launch, user chooses Run / Add / Delete / Modify / Exit
#                    inline failure menu (Add+retry / Modify / Skip) appears when a shortcut
#                    is missing or an app fails to start during the startup sequence
# - Presence mode  : after launch, Get-AppPresenceMode polls MainWindowHandle for $SettleSeconds;
#                    if a window appears -> 'Window' mode (skip only when window visible);
#                    if no window appears -> 'Tray' mode (skip when process running).
#                    No per-app flags needed; detection is fully automatic at runtime.
# - Exe allowlist  : Test-ExePathAllowed enforces that any exe accepted via user prompt or
#                    auto-discovery during repair lives under Program Files, Program Files (x86),
#                    or Windows. Paths outside these roots are rejected before any shortcut write.

$startMenu              = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
$MaxRepairDepth         = 3
$InitialDelaySeconds    = 10
$LaunchTimeoutSeconds   = 30
$PostLaunchPauseSeconds = 2
$SettleSeconds          = 5   # how long to wait for MainWindowHandle after launch before classifying as Tray

# Allowlisted root paths for exe repair; user-supplied or auto-discovered paths must live under one of these.
$AllowedExeRoots = @(
    $env:ProgramFiles,
    ${env:ProgramFiles(x86)},
    $env:SystemRoot
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

$apps = @(
    @{ Name = "Outlook";        LaunchType = "Win32"; ShortcutPath = "$startMenu\01 Outlook.lnk";        ProcessName = "OUTLOOK";            ExpectedExe = "OUTLOOK.EXE" },
    @{ Name = "Teams";          LaunchType = "Win32"; ShortcutPath = "$startMenu\02 Teams.lnk";          ProcessName = "ms-teams";           ExpectedExe = "ms-teams.exe" },
    @{ Name = "OneDrive";       LaunchType = "Win32"; ShortcutPath = "$startMenu\03 OneDrive.lnk";       ProcessName = "OneDrive";           ExpectedExe = "OneDrive.exe" },
    @{ Name = "Sticky Notes";   LaunchType = "Win32"; ShortcutPath = "$startMenu\04 Sticky Notes.lnk";   ProcessName = "ONENOTE";            ExpectedExe = "ONENOTE.EXE" },
    @{ Name = "OneNote";        LaunchType = "Win32"; ShortcutPath = "$startMenu\05 OneNote.lnk";        ProcessName = "ONENOTE";            ExpectedExe = "ONENOTE.EXE" },
    @{ Name = "Phone Link";     LaunchType = "Win32"; ShortcutPath = "$startMenu\06 Phone Link.lnk";     ProcessName = "PhoneExperienceHost"; ExpectedExe = "explorer.exe"; ExpectedArguments = "shell:appsFolder\Microsoft.YourPhone_8wekyb3d8bbwe!App" },
    @{ Name = "Microsoft Edge"; LaunchType = "Win32"; ShortcutPath = "$startMenu\07 Microsoft Edge.lnk"; ProcessName = "msedge";             ExpectedExe = "msedge.exe" },
    @{ Name = "Google Chrome";  LaunchType = "Win32"; ShortcutPath = "$startMenu\08 Google Chrome.lnk";  ProcessName = "chrome";             ExpectedExe = "chrome.exe" }
)

$WshShell = New-Object -ComObject WScript.Shell

# ---------------------------------------------------------------------------
# Presence mode detection
# Polls MainWindowHandle for up to $SettleSeconds after launch.
# Returns 'Window' if a visible window appears; 'Tray' if the process stays headless.
# ---------------------------------------------------------------------------
function Get-AppPresenceMode {
    param([string]$ProcessName, [int]$SettleSecs = $SettleSeconds)
    $elapsed = 0
    while ($elapsed -lt $SettleSecs) {
        $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
        if ($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) {
            return 'Window'
        }
        Start-Sleep -Seconds 1
        $elapsed++
    }
    # Process exists but no window appeared within settle time -> tray-only
    if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) {
        return 'Tray'
    }
    return $null   # process never started
}

# Returns $true when the app is considered 'already open' based on its detected presence mode.
# Called at the top of Start-Win32App / Start-AppxApp before attempting a launch.
function Test-AppAlreadyOpen {
    param([string]$ProcessName)
    $procs = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if (-not $procs) { return $false }
    # If any instance owns a visible window, it is open in Window mode
    if ($procs | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero }) { return $true }
    # Process running but no window -> could be Tray mode; treat as open so we don't relaunch
    return $true
}

# Waits for the app to reach a 'ready' state after launch.
# Phase 1 ($SettleSeconds): classify as Window or Tray.
# Phase 2 (remaining timeout): for Window mode, wait for MainWindowHandle; Tray mode is already confirmed.
function Wait-ForAppReady {
    param([string]$ProcessName, [int]$TimeoutSeconds = $LaunchTimeoutSeconds)

    # Phase 1: settle
    $mode = Get-AppPresenceMode -ProcessName $ProcessName -SettleSecs ([Math]::Min($SettleSeconds, $TimeoutSeconds))
    if ($null -eq $mode) {
        # Process never appeared during settle period; keep waiting for process up to full timeout
        $remaining = $TimeoutSeconds - $SettleSeconds
        $elapsed   = 0
        while ($elapsed -lt $remaining) {
            if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) { return $true }
            Start-Sleep -Seconds 1
            $elapsed++
        }
        return $false
    }

    Write-Host "  (presence mode: $mode)"

    if ($mode -eq 'Tray') {
        # Tray apps confirm on process presence alone - already confirmed in Get-AppPresenceMode
        return $true
    }

    # Window mode: wait for MainWindowHandle in remaining time
    $remaining = $TimeoutSeconds - $SettleSeconds
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
# Helper: numbered app picker used by main menu and inline failure menus
# ---------------------------------------------------------------------------
function Show-AppPicker {
    param([string]$Prompt)
    Write-Host "`n$Prompt"
    for ($i = 0; $i -lt $apps.Count; $i++) {
        $exists = if (Test-Path -LiteralPath $apps[$i].ShortcutPath -PathType Leaf) { "exists" } else { "missing" }
        Write-Host ("  [{0}] {1,-20}  {2}  ({3})" -f ($i + 1), $apps[$i].Name, $exists, $apps[$i].ShortcutPath)
    }
    Write-Host "  [0] Cancel"
    while ($true) {
        $choice = Read-Host "Select"
        if ($choice -match '^\d+$') {
            $idx = [int]$choice
            if ($idx -eq 0)                            { return $null }
            if ($idx -ge 1 -and $idx -le $apps.Count) { return $apps[$idx - 1] }
        }
        Write-Warning "Invalid selection. Enter a number between 0 and $($apps.Count)."
    }
}

# ---------------------------------------------------------------------------
# Shortcut management: Add / Delete / Modify
# ---------------------------------------------------------------------------
function Add-Shortcut {
    param($App)
    if (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf) {
        Write-Host "$($App.Name): shortcut already exists at '$($App.ShortcutPath)'."
        return
    }
    Initialize-Shortcut -App $App
}

function Remove-Shortcut {
    param($App)
    if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
        Write-Warning "$($App.Name): no shortcut found at '$($App.ShortcutPath)'."
        return
    }
    $confirm = Read-Host "Delete '$($App.ShortcutPath)'? (Y/N)"
    if ($confirm -ieq 'Y') {
        Remove-Item -LiteralPath $App.ShortcutPath -Force
        Write-Host "$($App.Name): shortcut deleted."
    } else {
        Write-Host "$($App.Name): deletion cancelled."
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
    $exePath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe
    if ($exePath) {
        if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
            $sc = $WshShell.CreateShortcut($App.ShortcutPath)
            $sc.TargetPath       = $exePath
            $sc.WorkingDirectory = Split-Path -Path $exePath -Parent
            $sc.Save()
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
    return $WshShell.CreateShortcut($ShortcutPath)
}

function Get-NearestExistingParent {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $current = Split-Path -Path $Path -Parent
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        if (Test-Path -LiteralPath $current -PathType Container) { return $current }
        $next = Split-Path -Path $current -Parent
        if ($next -eq $current) { break }
        $current = $next
    }
    return $null
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

# ---------------------------------------------------------------------------
# Security: validate an exe path against the allowlisted roots before
# persisting it into any shortcut target field.
# ---------------------------------------------------------------------------
function Test-ExePathAllowed {
    param([string]$ExePath)
    $full = [System.IO.Path]::GetFullPath($ExePath)
    foreach ($root in $AllowedExeRoots) {
        $rootFull = [System.IO.Path]::GetFullPath($root.TrimEnd('\\'))
        if ($full.StartsWith($rootFull + '\\', [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function Prompt-ForExactExePath {
    param([string]$AppName, [string]$ExpectedExe)
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
            Write-Warning "Path is outside allowed roots ($($AllowedExeRoots -join ', ')): $trimmed"
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
        $sc = $WshShell.CreateShortcut($App.ShortcutPath)
        $sc.TargetPath       = "C:\Windows\explorer.exe"
        $sc.Arguments        = $App.ExpectedArguments
        $sc.WorkingDirectory = "C:\Windows"
        $sc.Save()
        Write-Host "$($App.Name): shortcut created with Arguments: $($App.ExpectedArguments)"
        return
    }

    $exePath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe
    if ($exePath) {
        $sc = $WshShell.CreateShortcut($App.ShortcutPath)
        $sc.TargetPath       = $exePath
        $sc.WorkingDirectory = Split-Path -Path $exePath -Parent
        $sc.Save()
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
    $existingParent = Get-NearestExistingParent -Path $targetPath
    if ($existingParent) {
        Write-Host "$($App.Name): searching for $($App.ExpectedExe) under $existingParent (max depth $MaxRepairDepth)..."
        $foundExe = Find-ExeWithinDepth -RootFolder $existingParent -ExpectedExe $App.ExpectedExe -MaxDepth $MaxRepairDepth
        if ($foundExe) {
            if (-not (Test-ExePathAllowed -ExePath $foundExe.FullName)) {
                Write-Warning "$($App.Name): discovered exe is outside allowed roots. Skipping auto-repair: $($foundExe.FullName)"
            } else {
                Write-Host "$($App.Name): found replacement at $($foundExe.FullName). Updating shortcut."
                Update-ShortcutTarget -ShortcutPath $App.ShortcutPath -ExePath $foundExe.FullName -Arguments $App.ExpectedArguments
                return $foundExe.FullName
            }
        }
        Write-Warning "$($App.Name): $($App.ExpectedExe) not found within $MaxRepairDepth levels of $existingParent."
    } else {
        Write-Warning "$($App.Name): could not determine an existing parent folder from the broken target."
    }
    $manualPath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe
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
    if ($App.ExpectedArguments -match '\\([^\\!]+)!') {
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
            [xml]$manifest = Get-Content -LiteralPath $manifestPath -ErrorAction Stop
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
        if (Wait-ForAppReady -ProcessName $App.ProcessName -TimeoutSeconds $LaunchTimeoutSeconds) {
            Write-Host "$($App.Name): ready.`n"
            Start-Sleep -Seconds $PostLaunchPauseSeconds
            return $true
        } else {
            Write-Warning "$($App.Name): did not become ready within $LaunchTimeoutSeconds seconds.`n"
            return $false
        }
    } catch {
        Write-Warning "$($App.Name): launch failed. $_`n"
        return $false
    }
}

function Start-Win32App {
    param($App)
    if (Test-AppAlreadyOpen -ProcessName $App.ProcessName) {
        Write-Host "$($App.Name): already open. Skipping.`n"
        return $true
    }

    # ---- Inline failure menu: shortcut file not found ----
    if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
        Write-Warning "$($App.Name): shortcut file not found: $($App.ShortcutPath)"
        Write-Host "  [1] Add / fix shortcut now and retry launch"
        Write-Host "  [2] Modify an existing shortcut"
        Write-Host "  [3] Skip"
        $failChoice = Read-Host "Select"
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
        $WshShell.Run("`"$($App.ShortcutPath)`"", 1, $false)

        if (Wait-ForAppReady -ProcessName $App.ProcessName -TimeoutSeconds $LaunchTimeoutSeconds) {
            Write-Host "$($App.Name): ready.`n"
            Start-Sleep -Seconds $PostLaunchPauseSeconds
            return $true
        }

        # ---- Inline failure menu: app did not become ready in time ----
        Write-Warning "$($App.Name): did not become ready within $LaunchTimeoutSeconds seconds."
        Write-Host "  [1] Modify shortcut for $($App.Name) and retry"
        Write-Host "  [2] Modify a different shortcut"
        Write-Host "  [3] Skip"
        $timeoutChoice = Read-Host "Select"
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
# Main menu
# ---------------------------------------------------------------------------
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
        $app = Show-AppPicker -Prompt "Select app to ADD shortcut for:"
        if ($app) { Add-Shortcut -App $app }
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

# ---------------------------------------------------------------------------
# Startup sequence (option 1 or Enter)
# ---------------------------------------------------------------------------
Write-Host "Waiting $InitialDelaySeconds seconds for system to stabilize..."
Start-Sleep -Seconds $InitialDelaySeconds

Write-Host "`n--- Shortcut bootstrap ---"
foreach ($app in $apps) {
    if ($app.LaunchType -eq "Win32") {
        Initialize-Shortcut -App $app
    }
}
Write-Host "--- Bootstrap complete ---`n"

$failedApps = @()
foreach ($app in $apps) {
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
