<#
.SYNOPSIS
    Rebuilds apps.json from numbered .lnk files in the Start Menu Programs folder.

.DESCRIPTION
    Scans C:\ProgramData\Microsoft\Windows\Start Menu\Programs for .lnk files
    whose base names begin with one or two digits followed by a space (e.g. '01 Chrome.lnk',
    '9 Slack.lnk'). Non-numbered shortcuts are ignored.

    For each numbered .lnk the script reads the shortcut via WshShell to determine:
      - TargetPath  : if it ends in .exe and is not explorer.exe -> Win32 direct launch
      - TargetPath explorer.exe + Arguments starting with shell:appsFolder -> Appx packaged

    The resulting entries are written to apps.json in the same folder as this script,
    using the schemaVersion 1 wrapper expected by Win11startup.ps1.

    Fields that cannot be auto-detected (ExpectedPublisher, StartAppName, KnownAumid,
    AppxName, ProcessName for Appx apps) are left blank and can be filled in manually
    or via the Add menu in Win11startup.ps1.

.NOTES
    Run once to bootstrap or rebuild apps.json. Existing apps.json is overwritten.
    Must be run on the local machine where the .lnk files exist.
#>

[CmdletBinding()]
param(
    [string]$StartMenuFolder = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs",
    [string]$OutputPath      = (Join-Path $PSScriptRoot "apps.json")
)

$WshShell = New-Object -ComObject WScript.Shell

$lnkFiles = Get-ChildItem -LiteralPath $StartMenuFolder -Filter '*.lnk' -ErrorAction Stop |
    Where-Object { $_.BaseName -match '^\d{1,2}\s' } |
    Sort-Object Name

if ($lnkFiles.Count -eq 0) {
    Write-Warning "No numbered .lnk files found in '$StartMenuFolder'."
    exit 0
}

Write-Host "Found $($lnkFiles.Count) numbered shortcut(s). Scanning..."

$entries = @()
foreach ($file in $lnkFiles) {
    $sc         = $WshShell.CreateShortcut($file.FullName)
    $target     = $sc.TargetPath
    $args       = $sc.Arguments
    $baseName   = $file.BaseName                            # e.g. '01 Google Chrome'
    $appName    = ($baseName -replace '^\d{1,2}\s+', '')   # strip leading number

    $launchType        = 'Win32'
    $expectedExe       = ''
    $processName       = ''
    $expectedArguments = ''
    $startAppName      = ''
    $knownAumid        = ''
    $appxName          = ''

    $targetLeaf = if ($target) { [System.IO.Path]::GetFileName($target) } else { '' }

    if ($targetLeaf -ieq 'explorer.exe' -and $args -like 'shell:appsFolder\*') {
        # Appx / packaged app launched via shell:appsFolder
        $launchType        = 'Appx'
        $expectedExe       = 'explorer.exe'
        $expectedArguments = $args.Trim()
        $knownAumid        = ($args -replace '^shell:appsFolder\\', '').Trim()
        $processName       = ''   # cannot determine without running; fill in manually
        $startAppName      = $appName
        $appxName          = ($knownAumid -split '_')[0]   # rough package name prefix
    } elseif ($targetLeaf -like '*.exe') {
        # Standard Win32 exe shortcut
        $launchType        = 'Win32'
        $expectedExe       = $targetLeaf
        $processName       = [System.IO.Path]::GetFileNameWithoutExtension($targetLeaf)
        $expectedArguments = if (-not [string]::IsNullOrWhiteSpace($args)) { $args.Trim() } else { '' }
    } else {
        # Unknown / non-exe target - record as Win32 with blank exe; flag for manual review
        $launchType  = 'Win32'
        $expectedExe = $targetLeaf
        $processName = [System.IO.Path]::GetFileNameWithoutExtension($targetLeaf)
        Write-Warning "'$($file.Name)': unexpected target '$target'. Entry created with blank fields - review manually."
    }

    $entries += [PSCustomObject]@{
        Name              = $appName
        LaunchType        = $launchType
        ShortcutPath      = $file.FullName
        ProcessName       = $processName
        ExpectedExe       = $expectedExe
        ExpectedPublisher = ''
        ExpectedArguments = $expectedArguments
        StartAppName      = $startAppName
        KnownAumid        = $knownAumid
        AppxName          = $appxName
        PresenceMode      = $null
    }

    Write-Host ("  [{0}] {1,-28} {2}" -f $launchType, $appName, $file.Name)
}

$wrapper = [PSCustomObject]@{
    schemaVersion = 1
    apps          = $entries
}

$wrapper | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Host "`napps.json written to: $OutputPath ($($entries.Count) entries)"
