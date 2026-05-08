# Curated startup launcher with self-healing shortcut repair
# - No UWP logic
# - Checks each shortcut target
# - If target is missing, climbs to nearest existing parent folder
# - Searches downward up to a fixed depth of 3 for the expected EXE
# - If still not found, prompts user for the exact EXE path
# - Updates the shortcut and launches the app

$startMenu = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
$MaxRepairDepth = 3
$InitialDelaySeconds = 10
$LaunchTimeoutSeconds = 30
$PostLaunchPauseSeconds = 2

$apps = @(
    @{ Name = "Outlook";        ShortcutPath = "$startMenu\01 Outlook.lnk";        ProcessName = "OUTLOOK";             ExpectedExe = "OUTLOOK.EXE" },
    @{ Name = "Teams";          ShortcutPath = "$startMenu\02 Teams.lnk";          ProcessName = "ms-teams";            ExpectedExe = "ms-teams.exe" },
    @{ Name = "OneDrive";       ShortcutPath = "$startMenu\03 OneDrive.lnk";       ProcessName = "OneDrive";            ExpectedExe = "OneDrive.exe" },
    @{ Name = "ShareFile";      ShortcutPath = "$startMenu\04 ShareFile.lnk";      ProcessName = "ShareFile";           ExpectedExe = "ShareFile.exe" },
    @{ Name = "Greenshot";      ShortcutPath = "$startMenu\05 Greenshot.lnk";      ProcessName = "Greenshot";           ExpectedExe = "Greenshot.exe" },
    @{ Name = "Sticky Notes";   ShortcutPath = "$startMenu\06 Sticky Notes.lnk";   ProcessName = "Sticky-Note";         ExpectedExe = "StikyNot.exe" },
    @{ Name = "OneNote";        ShortcutPath = "$startMenu\07 OneNote.lnk";        ProcessName = "ONENOTE";             ExpectedExe = "ONENOTE.EXE" },
    @{ Name = "SAP GUI";        ShortcutPath = "$startMenu\08 SAP GUI.lnk";        ProcessName = "saplogon";            ExpectedExe = "saplogon.exe" },
    @{ Name = "Notepad++";      ShortcutPath = "$startMenu\09 notepad++.lnk";      ProcessName = "notepad++";           ExpectedExe = "notepad++.exe" },
    @{ Name = "Phone Link";     ShortcutPath = "$startMenu\10 Phone Link.lnk";     ProcessName = "PhoneExperienceHost"; ExpectedExe = "PhoneExperienceHost.exe" },
    @{ Name = "Microsoft Edge"; ShortcutPath = "$startMenu\11 Microsoft Edge.lnk"; ProcessName = "msedge";              ExpectedExe = "msedge.exe" },
    @{ Name = "Google Chrome";  ShortcutPath = "$startMenu\12 Google Chrome.lnk";  ProcessName = "chrome";              ExpectedExe = "chrome.exe" }
)

$WshShell = New-Object -ComObject WScript.Shell

function Get-ShortcutObject {
    param([string]$ShortcutPath)
    if (-not (Test-Path -LiteralPath $ShortcutPath)) {
        throw "Shortcut not found: $ShortcutPath"
    }
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
    if (-not $candidateFull.StartsWith($baseFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        return [int]::MaxValue
    }
    $relative = $candidateFull.Substring($baseFull.Length).TrimStart('\\')
    if ([string]::IsNullOrWhiteSpace($relative)) { return 0 }
    return ($relative -split '[\\//]').Count
}

function Find-ExeWithinDepth {
    param([string]$RootFolder, [string]$ExpectedExe, [int]$MaxDepth = 3)
    if (-not (Test-Path -LiteralPath $RootFolder -PathType Container)) { return $null }
    $matches = Get-ChildItem -LiteralPath $RootFolder -Filter $ExpectedExe -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { (Get-RelativeDepth -BasePath $RootFolder -CandidatePath $_.FullName) -le $MaxDepth } |
        Sort-Object FullName
    return $matches | Select-Object -First 1
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
            Write-Warning "Path does not exist or is not a file: $trimmed"
            continue
        }
        if ([System.IO.Path]::GetFileName($trimmed) -ine $ExpectedExe) {
            Write-Warning "File name must be exactly $ExpectedExe"
            continue
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

function Start-AppAndRepairShortcut {
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

        Write-Host "$($App.Name): launching $launchPath"
        Start-Process -FilePath $launchPath -ErrorAction Stop

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

Write-Host "Waiting $InitialDelaySeconds seconds for system to stabilize..."
Start-Sleep -Seconds $InitialDelaySeconds

$failedApps = @()
foreach ($app in $apps) {
    $ok = Start-AppAndRepairShortcut -App $app
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
