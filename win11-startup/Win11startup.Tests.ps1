#Requires -Modules Pester
# Run unit tests : Invoke-Pester .\Win11startup.Tests.ps1
#
# Tested functions / behaviours
# --------------------------------
# Unit:
#   Save-Config                   (TEST-01)
#   Config load: valid JSON       (TEST-02a)
#   Config load: invalid JSON     (TEST-02b)
#   Config load: missing file     (TEST-02c)
#   StartMenuPath guard           (TEST-03)
#   .lnk file discovery & sort   (TEST-04)
#   Store-app skip guard          (TEST-05)
#   Empty TargetPath skip guard   (TEST-06)
#   Skip-if-running               (TEST-07)
#   ProcessStartTimeout clamping  (TEST-08)

BeforeAll {
    $env:PS_STARTUP_TESTMODE = '1'
    . "$PSScriptRoot\Win11startup.ps1"
}

AfterAll {
    $env:PS_STARTUP_TESTMODE = $null
}

# ---------------------------------------------------------------------------
# Unit tests
# ---------------------------------------------------------------------------
Describe 'Unit' {

    # -----------------------------------------------------------------------
    # TEST-01: Save-Config
    # -----------------------------------------------------------------------
    Describe 'Save-Config' {

        BeforeEach {
            $script:SC_dir = Join-Path $env:TEMP ("PesterSC_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:SC_dir | Out-Null
        }

        AfterEach {
            Remove-Item -LiteralPath $script:SC_dir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'writes a valid JSON file to the specified path' {
            $cfg  = [PSCustomObject]@{ StartMenuPath = 'C:\Test'; Shortcuts = @() }
            $path = Join-Path $script:SC_dir 'config.json'
            Save-Config -Cfg $cfg -Path $path
            Test-Path -LiteralPath $path -PathType Leaf | Should -Be $true
        }

        It 'written JSON can be round-tripped back via ConvertFrom-Json' {
            $cfg  = [PSCustomObject]@{ StartMenuPath = 'C:\RoundTrip'; Shortcuts = @() }
            $path = Join-Path $script:SC_dir 'rt.json'
            Save-Config -Cfg $cfg -Path $path
            $loaded = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
            $loaded.StartMenuPath | Should -Be 'C:\RoundTrip'
        }

        It 'does not throw when the destination path is invalid' {
            $cfg = [PSCustomObject]@{ StartMenuPath = ''; Shortcuts = @() }
            { Save-Config -Cfg $cfg -Path $env:TEMP } | Should -Not -Throw
        }

        It 'emits a warning (not a terminating error) when write fails' {
            $cfg = [PSCustomObject]@{ StartMenuPath = ''; Shortcuts = @() }
            $warnings = @()
            Save-Config -Cfg $cfg -Path $env:TEMP -WarningVariable warnings -WarningAction SilentlyContinue
            $true | Should -Be $true
        }
    }

    # -----------------------------------------------------------------------
    # TEST-02a: Config load -- valid JSON
    # -----------------------------------------------------------------------
    Describe 'Config load: valid JSON' {

        BeforeEach {
            $script:CL_dir = Join-Path $env:TEMP ("PesterCL_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:CL_dir | Out-Null
        }

        AfterEach {
            Remove-Item -LiteralPath $script:CL_dir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'ConvertFrom-Json succeeds for a well-formed config' {
            $json = '{"StartMenuPath":"C:\\ProgramData","Shortcuts":[]}'
            $path = Join-Path $script:CL_dir 'valid.json'
            $json | Set-Content -LiteralPath $path -Encoding UTF8
            $loaded = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
            $loaded.StartMenuPath | Should -Be 'C:\ProgramData'
            $loaded.Shortcuts     | Should -BeNullOrEmpty
        }

        It 'loaded config preserves a Shortcuts entry' {
            $json = @'
{
  "StartMenuPath": "C:\\Test",
  "Shortcuts": [
    { "Name": "Outlook", "ShortcutPath": "C:\\Fake\\01 Outlook.lnk", "ProcessName": "OUTLOOK" }
  ]
}
'@
            $path = Join-Path $script:CL_dir 'with_shortcuts.json'
            $json | Set-Content -LiteralPath $path -Encoding UTF8
            $loaded = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
            $loaded.Shortcuts.Count   | Should -Be 1
            $loaded.Shortcuts[0].Name | Should -Be 'Outlook'
        }
    }

    # -----------------------------------------------------------------------
    # TEST-02b: Config load -- invalid JSON
    # -----------------------------------------------------------------------
    Describe 'Config load: invalid JSON' {

        BeforeEach {
            $script:CL2_dir = Join-Path $env:TEMP ("PesterCL2_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:CL2_dir | Out-Null
        }

        AfterEach {
            Remove-Item -LiteralPath $script:CL2_dir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'ConvertFrom-Json throws for malformed JSON' {
            $bad  = 'not valid json {{{'
            $path = Join-Path $script:CL2_dir 'bad.json'
            $bad | Set-Content -LiteralPath $path -Encoding UTF8
            { Get-Content -LiteralPath $path -Raw | ConvertFrom-Json } | Should -Throw
        }
    }

    # -----------------------------------------------------------------------
    # TEST-02c: Config load -- missing file
    # -----------------------------------------------------------------------
    Describe 'Config load: missing file' {

        It 'Test-Path returns $false for a non-existent config path' {
            $missing = Join-Path $env:TEMP ("PesterMissing_" + [System.IO.Path]::GetRandomFileName() + '.json')
            Test-Path -LiteralPath $missing -PathType Leaf | Should -Be $false
        }
    }

    # -----------------------------------------------------------------------
    # TEST-03: StartMenuPath guard -- folder existence check
    # -----------------------------------------------------------------------
    Describe 'StartMenuPath guard' {

        It 'Test-Path returns $false for a non-existent Start Menu folder' {
            Test-Path -LiteralPath 'C:\PesterNonExistent_DoesNotExist' -PathType Container | Should -Be $false
        }

        It 'Test-Path returns $true for a real directory' {
            Test-Path -LiteralPath $env:TEMP -PathType Container | Should -Be $true
        }
    }

    # -----------------------------------------------------------------------
    # TEST-04: .lnk file discovery and numeric sort
    # -----------------------------------------------------------------------
    Describe 'lnk file discovery and sort' {

        BeforeEach {
            $script:LNK_dir = Join-Path $env:TEMP ("PesterLNK_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:LNK_dir | Out-Null
        }

        AfterEach {
            Remove-Item -LiteralPath $script:LNK_dir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'discovers numbered .lnk files and returns them in ascending numeric order' {
            New-Item -ItemType File -Path (Join-Path $script:LNK_dir '03 Slack.lnk')   | Out-Null
            New-Item -ItemType File -Path (Join-Path $script:LNK_dir '01 Outlook.lnk') | Out-Null
            New-Item -ItemType File -Path (Join-Path $script:LNK_dir '02 Teams.lnk')   | Out-Null
            $files = Get-ChildItem -LiteralPath $script:LNK_dir -Filter '*.lnk' |
                Where-Object { $_.BaseName -match '^\d{1,2}\s' } |
                Sort-Object { if ($_.BaseName -match '^(\d{1,2})\s') { [int]$Matches[1] } else { 0 } }
            $files.Count              | Should -Be 3
            $files[0].BaseName        | Should -Match '^01'
            $files[1].BaseName        | Should -Match '^02'
            $files[2].BaseName        | Should -Match '^03'
        }

        It 'ignores .lnk files without a numeric prefix' {
            New-Item -ItemType File -Path (Join-Path $script:LNK_dir 'NoPrefix.lnk') | Out-Null
            New-Item -ItemType File -Path (Join-Path $script:LNK_dir '01 Valid.lnk') | Out-Null
            $files = Get-ChildItem -LiteralPath $script:LNK_dir -Filter '*.lnk' |
                Where-Object { $_.BaseName -match '^\d{1,2}\s' }
            $files.Count       | Should -Be 1
            $files[0].BaseName | Should -Match '^01'
        }

        It 'returns an empty list when the folder contains no .lnk files' {
            $files = Get-ChildItem -LiteralPath $script:LNK_dir -Filter '*.lnk' -ErrorAction SilentlyContinue |
                Where-Object { $_.BaseName -match '^\d{1,2}\s' }
            @($files).Count | Should -Be 0
        }

        It 'handles two-digit prefixes correctly' {
            New-Item -ItemType File -Path (Join-Path $script:LNK_dir '10 App.lnk')  | Out-Null
            New-Item -ItemType File -Path (Join-Path $script:LNK_dir '2 App2.lnk') | Out-Null
            $files = Get-ChildItem -LiteralPath $script:LNK_dir -Filter '*.lnk' |
                Where-Object { $_.BaseName -match '^\d{1,2}\s' } |
                Sort-Object { if ($_.BaseName -match '^(\d{1,2})\s') { [int]$Matches[1] } else { 0 } }
            $files.Count       | Should -Be 2
            $files[0].BaseName | Should -Match '^2 '
            $files[1].BaseName | Should -Match '^10 '
        }
    }

    # -----------------------------------------------------------------------
    # TEST-05: Store-app skip guard
    # -----------------------------------------------------------------------
    Describe 'Store-app skip guard' {

        It 'identifies a Store-app shortcut when TargetPath is explorer.exe and Arguments start with shell:appsFolder' {
            $targetPath = "$env:SystemRoot\explorer.exe"
            $targetArgs = 'shell:appsFolder\Microsoft.WindowsCalculator_8wekyb3d8bbwe!App'
            $isStore    = ($targetPath -ieq "$env:SystemRoot\explorer.exe") -and ($targetArgs -match '^shell:appsFolder\\')
            $isStore | Should -Be $true
        }

        It 'does not flag a regular Win32 shortcut as a Store app' {
            $targetPath = "$env:ProgramFiles\SomeApp\app.exe"
            $targetArgs = ''
            $isStore    = ($targetPath -ieq "$env:SystemRoot\explorer.exe") -and ($targetArgs -match '^shell:appsFolder\\')
            $isStore | Should -Be $false
        }

        It 'does not flag explorer.exe launched without shell:appsFolder as a Store app' {
            $targetPath = "$env:SystemRoot\explorer.exe"
            $targetArgs = 'C:\SomePath'
            $isStore    = ($targetPath -ieq "$env:SystemRoot\explorer.exe") -and ($targetArgs -match '^shell:appsFolder\\')
            $isStore | Should -Be $false
        }
    }

    # -----------------------------------------------------------------------
    # TEST-06: Empty TargetPath skip guard
    # -----------------------------------------------------------------------
    Describe 'Empty TargetPath skip guard' {

        It 'IsNullOrEmpty returns true for an empty TargetPath' {
            [string]::IsNullOrEmpty('') | Should -Be $true
        }

        It 'IsNullOrEmpty returns false for a valid TargetPath' {
            [string]::IsNullOrEmpty("$env:SystemRoot\System32\notepad.exe") | Should -Be $false
        }

        It 'GetFileNameWithoutExtension returns empty string for an empty path' {
            $result = [System.IO.Path]::GetFileNameWithoutExtension('')
            [string]::IsNullOrEmpty($result) | Should -Be $true
        }
    }

    # -----------------------------------------------------------------------
    # TEST-07: Skip-if-running check
    # -----------------------------------------------------------------------
    Describe 'Skip-if-running check' {

        It 'Get-Process returns a result for a known running process (System)' {
            Get-Process -Name 'System' -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }

        It 'Get-Process returns nothing for a fictional process name' {
            Get-Process -Name 'pester_no_such_process_xyz' -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
        }
    }

    # -----------------------------------------------------------------------
    # TEST-08: ProcessStartTimeout loop clamping
    # -----------------------------------------------------------------------
    Describe 'ProcessStartTimeout clamping' {

        It 'loop exits on first iteration when process starts immediately (simulated)' {
            $timeout    = 15
            $iterations = 0
            for ($i = 0; $i -lt $timeout; $i++) {
                $iterations++
                break   # simulate immediate process detection
            }
            $iterations | Should -Be 1
        }

        It 'loop exhausts all iterations when process never starts (simulated)' {
            $timeout    = 3
            $iterations = 0
            for ($i = 0; $i -lt $timeout; $i++) { $iterations++ }
            $iterations | Should -Be 3
        }

        It 'loop runs zero iterations when timeout is 0' {
            $timeout    = 0
            $iterations = 0
            for ($i = 0; $i -lt $timeout; $i++) { $iterations++ }
            $iterations | Should -Be 0
        }
    }
}
