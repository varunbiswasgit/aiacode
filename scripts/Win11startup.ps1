# Curated startup launcher with self-healing shortcut repair
# - Win32 apps  : launched via shortcut with depth-3 repair and user-prompt fallback
#                 optional Arguments field overrides shortcut arguments at launch
# - Appx apps   : AUMID resolved at runtime (Get-StartApps -> KnownAumid verification -> AppxPackage manifest)
#                 KnownAumid used only as primary candidate, not sole source of truth
# - Sticky Notes: Win32, ONENOTE.EXE with Arguments = "/memoryWindow start"

$startMenu = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
$MaxRepairDepth         = 3
$InitialDelaySeconds    = 10
$LaunchTimeoutSeconds   = 30
$PostLaunchPauseSeconds = 2

# LaunchType controls the launch strategy per entry:
#   'Win32' - shortcut-based with self-healing repair and user-prompt fallback
#             optional 'Arguments' field is passed to Start-Process -ArgumentList
#             if absent, app is launched with no extra arguments
#   'Appx'  - AUMID resolved dynamically at runtime; KnownAumid is the primary candidate
#
# Appx fields:
#   KnownAumid   - last-known AUMID (PackageFamilyName!AppId); verified at runtime, not assumed static
#   AppxName     - partial package name used to discover AUMID if KnownAumid is stale
#   StartAppName - display name pattern used in Get-StartApps discovery
#   ProcessName  - process name used to detect if already running and confirm launch

$apps = @(
    @{ Name = "Outlook";        LaunchType = "Win32"; ShortcutPath = "$startMenu\01 Outlook.lnk";        ProcessName = "OUTLOOK";             ExpectedExe = "OUTLOOK.EXE" },
    @{ Name = "Teams";          LaunchType = "Win32"; ShortcutPath = "$startMenu\02 Teams.lnk";          ProcessName = "ms-teams";            ExpectedExe = "ms-teams.exe" },
    @{ Name = "OneDrive";       LaunchType = "Win32"; ShortcutPath = "$startMenu\03 OneDrive.lnk";       ProcessName = "OneDrive";            ExpectedExe = "OneDrive.exe" },
    @{ Name = "ShareFile";      LaunchType = "Win32"; ShortcutPath = "$startMenu\04 ShareFile.lnk";      ProcessName = "ShareFile";           ExpectedExe = "ShareFile.exe" },
    @{ Name = "Greenshot";      LaunchType = "Win32"; ShortcutPath = "$startMenu\05 Greenshot.lnk";      ProcessName = "Greenshot";           ExpectedExe = "Greenshot.exe" },
    @{ Name = "Sticky Notes";   LaunchType = "Win32"; ShortcutPath = "$startMenu\06 Sticky Notes.lnk";   ProcessName = "ONENOTE";             ExpectedExe = "ONENOTE.EXE";  Arguments = "/memoryWindow start" },
    @{ Name = "OneNote";        LaunchType = "Win32"; ShortcutPath = "$startMenu\07 OneNote.lnk";        ProcessName = "ONENOTE";             ExpectedExe = "ONENOTE.EXE" },
    @{ Name = "SAP GUI";        LaunchType = "Win32"; ShortcutPath = "$startMenu\08 SAP GUI.lnk";        ProcessName = "saplogon";            ExpectedExe = "saplogon.exe" },
    @{ Name = "Notepad++";      LaunchType = "Win32"; ShortcutPath = "$startMenu\09 notepad++.lnk";      ProcessName = "notepad++";           ExpectedExe = "notepad++.exe" },
    @{ Name = "Phone Link";     LaunchType = "Appx";  KnownAumid = "Microsoft.YourPhone_8wekyb3d8bbwe!App"; AppxName = "YourPhone"; StartAppName = "Phone Link"; ProcessName = "PhoneExperienceHost" },
    @{ Name = "Microsoft Edge"; LaunchType = "Win32"; ShortcutPath = "$startMenu\11 Microsoft Edge.lnk"; ProcessName = "msedge";              ExpectedExe = "msedge.exe" },
    @{ Name = "Google Chrome";  LaunchType = "Win32"; ShortcutPath = "$startMenu\12 Google Chrome.lnk";  ProcessName = "chrome";              ExpectedExe = "chrome.exe" }
)

$WshShell = New-Object -ComObject WScript.Shell

# ---------------------------------------------------------------------------
# AUMID resolution: Get-StartApps -> KnownAumid verification -> AppxPackage manifest
# ---------------------------------------------------------------------------
function Resolve-Aumid {
    param($App)

    # Step 1: Get-StartApps by display name (most reliable - always reflects installed state)
    $startApp = Get-StartApps | Where-Object { $_.Name -like "*$($App.StartAppName)*" } | Select-Object -First 1
    if ($startApp) {
        Write-Host "$($App.Name): AUMID resolved via Get-StartApps: $($startApp.AppID)"
        return $startApp.AppID
    }

    # Step 2: Verify KnownAumid is still installed before trusting it
    if (-not [string]::IsNullOrWhiteSpace($App.KnownAumid)) {
        $knownPfn = ($App.KnownAumid -split '!')[0]
        $installed = Get-AppxPackage | Where-Object { $_.PackageFamilyName -eq $knownPfn } | Select-Object -First 1
        if ($installed) {
            Write-Host "$($App.Name): KnownAumid verified as installed: $($App.KnownAumid)"
            return $App.KnownAumid
        }
        Write-Warning "$($App.Name): KnownAumid package family '$knownPfn' not found on this system."
    }

    # Step 3: Discover via AppxPackage name search and read the manifest for AppId
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

# ---------------------------------------------------------------------------
# Win32 helpers
# ---------------------------------------------------------------------------
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
    param([string]$ShortcutPath, [string]$ExePath)
    $shortcut = Get-ShortcutObject -ShortcutPath $ShortcutPath
    $shortcut.TargetPath       = $ExePath
    $shortcut.WorkingDirectory = Split-Path -Path $ExePath -Parent
    $shortcut.Save()
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
        return $trimmed
    }
}

function Resolve-LaunchPath {
    param($App)
    $shortcut   = Get-ShortcutObject -ShortcutPath $App.ShortcutPath
    $targetPath = $shortcut.TargetPath

    if (-not [string]::IsNullOrWhiteSpace($targetPath) -and (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
        return $targetPath
    }

    Write-Warning "$($App.Name): shortcut target missing or invalid: $targetPath"

    $existingParent = Get-NearestExistingParent -Path $targetPath
    if ($existingParent) {
        Write-Host "$($App.Name): searching for $($App.ExpectedExe) under $existingParent (max depth $MaxRepairDepth)..."
        $foundExe = Find-ExeWithinDepth -RootFolder $existingParent -ExpectedExe $App.ExpectedExe -MaxDepth $MaxRepairDepth
        if ($foundExe) {
            Write-Host "$($App.Name): found replacement at $($foundExe.FullName). Updating shortcut."
            Update-ShortcutTarget -ShortcutPath $App.ShortcutPath -ExePath $foundExe.FullName
            return $foundExe.FullName
        }
        Write-Warning "$($App.Name): $($App.ExpectedExe) not found within $MaxRepairDepth levels of $existingParent."
    } else {
        Write-Warning "$($App.Name): could not determine an existing parent folder from the broken target."
    }

    $manualPath = Prompt-ForExactExePath -AppName $App.Name -ExpectedExe $App.ExpectedExe
    if ($manualPath) {
        Update-ShortcutTarget -ShortcutPath $App.ShortcutPath -ExePath $manualPath
        return $manualPath
    }

    return $null
}

# ---------------------------------------------------------------------------
# Process wait helper
# ---------------------------------------------------------------------------
function Wait-ForProcessStart {
    param([string]$ProcessName, [int]$TimeoutSeconds = 30)
    $elapsed = 0
    while ($elapsed -lt $TimeoutSeconds) {
        if (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue) { return $true }
        Start-Sleep -Seconds 1
        $elapsed++
    }
    return $false
}

# ---------------------------------------------------------------------------
# Launch functions
# ---------------------------------------------------------------------------
function Start-AppxApp {
    param($App)

    if (Get-Process -Name $App.ProcessName -ErrorAction SilentlyContinue) {
        Write-Host "$($App.Name): '$($App.ProcessName)' already running. Skipping.`n"
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

        if (Wait-ForProcessStart -ProcessName $App.ProcessName -TimeoutSeconds $LaunchTimeoutSeconds) {
            Write-Host "$($App.Name): '$($App.ProcessName)' is now running.`n"
            Start-Sleep -Seconds $PostLaunchPauseSeconds
            return $true
        } else {
            Write-Warning "$($App.Name): '$($App.ProcessName)' did not appear within $LaunchTimeoutSeconds seconds.`n"
            return $false
        }
    } catch {
        Write-Warning "$($App.Name): launch failed. $_`n"
        return $false
    }
}

function Start-Win32App {
    param($App)

    if (Get-Process -Name $App.ProcessName -ErrorAction SilentlyContinue) {
        Write-Host "$($App.Name): '$($App.ProcessName)' already running. Skipping.`n"
        return $true
    }

    if (-not (Test-Path -LiteralPath $App.ShortcutPath -PathType Leaf)) {
        Write-Warning "$($App.Name): shortcut file not found: $($App.ShortcutPath)`n"
        return $false
    }

    try {
        $launchPath = Resolve-LaunchPath -App $App
        if (-not $launchPath) {
            Write-Warning "$($App.Name): no valid executable path. Skipping.`n"
            return $false
        }

        # Use explicit Arguments field if present; otherwise launch with no extra arguments
        if (-not [string]::IsNullOrWhiteSpace($App.Arguments)) {
            Write-Host "$($App.Name): launching $launchPath $($App.Arguments)"
            Start-Process -FilePath $launchPath -ArgumentList $App.Arguments -ErrorAction Stop
        } else {
            Write-Host "$($App.Name): launching $launchPath"
            Start-Process -FilePath $launchPath -ErrorAction Stop
        }

        if (Wait-ForProcessStart -ProcessName $App.ProcessName -TimeoutSeconds $LaunchTimeoutSeconds) {
            Write-Host "$($App.Name): '$($App.ProcessName)' is now running.`n"
            Start-Sleep -Seconds $PostLaunchPauseSeconds
            return $true
        } else {
            Write-Warning "$($App.Name): '$($App.ProcessName)' did not appear within $LaunchTimeoutSeconds seconds.`n"
            return $false
        }
    } catch {
        Write-Warning "$($App.Name): launch failed. $_`n"
        return $false
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Write-Host "Waiting $InitialDelaySeconds seconds for system to stabilize..."
Start-Sleep -Seconds $InitialDelaySeconds

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
