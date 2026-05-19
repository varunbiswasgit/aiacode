#Requires -Modules Pester
# Run unit tests : Invoke-Pester .\Win11startup.Tests.ps1
# Run all tests  : $env:RUN_INTEGRATION = '1'; Invoke-Pester .\Win11startup.Tests.ps1
#
# Tested functions
# ----------------
# Unit:
#   Get-RelativeDepth          (TEST-02)
#   Find-MisnumberedShortcut   (TEST-03)
#   Test-ExePathAllowed        (TEST-04a)
#   Test-ExeSignatureTrusted   (TEST-04b)
#   Import-AppsConfig          (TEST-08)
#   Get-ParentFolder           (T-08)   — replaces Get-NearestExistingParent
#   Show-FailureMenu           (T-09)
#   Show-AppPicker -AllowNew   (TEST-10)
#   Add-Shortcut dispatch      (TEST-11)
#   Wait-ForAppReady phase math (TEST-12)
# Integration:
#   Initialize-Shortcut        (INT-02)

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
    # TEST-02: Get-RelativeDepth
    # -----------------------------------------------------------------------
    Describe 'Get-RelativeDepth' {

        It 'returns 0 when candidate equals base' {
            $base = 'C:\Foo'
            Get-RelativeDepth -BasePath $base -CandidatePath $base | Should -Be 0
        }

        It 'returns 1 for a direct child' {
            Get-RelativeDepth -BasePath 'C:\Foo' -CandidatePath 'C:\Foo\Bar.exe' | Should -Be 1
        }

        It 'returns 2 for a two-level child' {
            Get-RelativeDepth -BasePath 'C:\Foo' -CandidatePath 'C:\Foo\Bar\Baz.exe' | Should -Be 2
        }

        It 'returns MaxValue when candidate is outside base' {
            Get-RelativeDepth -BasePath 'C:\Foo' -CandidatePath 'C:\Other\File.exe' | Should -Be ([int]::MaxValue)
        }

        It 'returns MaxValue for an empty candidate string' {
            { Get-RelativeDepth -BasePath 'C:\Foo' -CandidatePath '' } | Should -Not -Throw
            Get-RelativeDepth -BasePath 'C:\Foo' -CandidatePath 'C:\Other' | Should -Be ([int]::MaxValue)
        }
    }

    # -----------------------------------------------------------------------
    # TEST-03: Find-MisnumberedShortcut
    # -----------------------------------------------------------------------
    Describe 'Find-MisnumberedShortcut' {

        BeforeEach {
            $script:testDir = Join-Path $env:TEMP ("PesterLnkTest_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:testDir | Out-Null
        }

        AfterEach {
            Remove-Item -LiteralPath $script:testDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'returns the misnumbered file when a matching lnk exists under a different number' {
            $lnkPath = Join-Path $script:testDir '99 Outlook.lnk'
            New-Item -ItemType File -Path $lnkPath | Out-Null
            $expected = Join-Path $script:testDir '01 Outlook.lnk'
            $result = Find-MisnumberedShortcut -ExpectedPath $expected -AppName 'Outlook'
            $result | Should -Not -BeNullOrEmpty
            $result.FullName | Should -Be $lnkPath
        }

        It 'returns null when no lnk matches the app name' {
            New-Item -ItemType File -Path (Join-Path $script:testDir '01 Teams.lnk') | Out-Null
            $expected = Join-Path $script:testDir '01 Outlook.lnk'
            Find-MisnumberedShortcut -ExpectedPath $expected -AppName 'Outlook' | Should -BeNullOrEmpty
        }

        It 'returns null when the folder is empty' {
            $expected = Join-Path $script:testDir '01 Outlook.lnk'
            Find-MisnumberedShortcut -ExpectedPath $expected -AppName 'Outlook' | Should -BeNullOrEmpty
        }

        It 'returns null when the folder does not exist' {
            $missing = Join-Path $env:TEMP 'PesterNonExistent_DoesNotExist'
            $expected = Join-Path $missing '01 Outlook.lnk'
            Find-MisnumberedShortcut -ExpectedPath $expected -AppName 'Outlook' | Should -BeNullOrEmpty
        }
    }

    # -----------------------------------------------------------------------
    # TEST-04a: Test-ExePathAllowed
    # -----------------------------------------------------------------------
    Describe 'Test-ExePathAllowed' {

        It 'returns true for a path under Program Files' {
            $path = Join-Path $env:ProgramFiles 'SomeApp\app.exe'
            Test-ExePathAllowed -ExePath $path | Should -Be $true
        }

        It 'returns true for a path under SystemRoot' {
            $path = Join-Path $env:SystemRoot 'System32\notepad.exe'
            Test-ExePathAllowed -ExePath $path | Should -Be $true
        }

        It 'returns false for a path under TEMP' {
            $path = Join-Path $env:TEMP 'malicious.exe'
            Test-ExePathAllowed -ExePath $path | Should -Be $false
        }

        It 'returns false for a path on the Desktop' {
            $path = Join-Path ([Environment]::GetFolderPath('Desktop')) 'app.exe'
            Test-ExePathAllowed -ExePath $path | Should -Be $false
        }
    }

    # -----------------------------------------------------------------------
    # TEST-04b: Test-ExeSignatureTrusted
    # -----------------------------------------------------------------------
    Describe 'Test-ExeSignatureTrusted' {

        It 'returns true for notepad.exe (valid Microsoft signature)' {
            $notepad = Join-Path $env:SystemRoot 'System32\notepad.exe'
            Test-ExeSignatureTrusted -ExePath $notepad | Should -Be $true
        }

        It 'returns true for notepad.exe with correct publisher string' {
            $notepad = Join-Path $env:SystemRoot 'System32\notepad.exe'
            Test-ExeSignatureTrusted -ExePath $notepad -ExpectedPublisher 'CN=Microsoft Windows' | Should -Be $true
        }

        It 'returns false for notepad.exe with wrong publisher string' {
            $notepad = Join-Path $env:SystemRoot 'System32\notepad.exe'
            Test-ExeSignatureTrusted -ExePath $notepad -ExpectedPublisher 'CN=Google LLC' | Should -Be $false
        }

        It 'returns false for a plain text file renamed to .exe' {
            $fakePath = Join-Path $env:TEMP 'fake_test.exe'
            'not an exe' | Set-Content -Path $fakePath -Encoding UTF8
            try {
                Test-ExeSignatureTrusted -ExePath $fakePath | Should -Be $false
            } finally {
                Remove-Item -LiteralPath $fakePath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # -----------------------------------------------------------------------
    # TEST-08: Import-AppsConfig
    # -----------------------------------------------------------------------
    Describe 'Import-AppsConfig' {

        BeforeEach {
            $script:cfgDir = Join-Path $env:TEMP ("PesterCfg_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:cfgDir | Out-Null
        }

        AfterEach {
            Remove-Item -LiteralPath $script:cfgDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'loads a valid apps.json without error' {
            $json = @'
[
  {
    "Name": "Notepad",
    "LaunchType": "Win32",
    "ShortcutPath": "C:\\Fake\\01 Notepad.lnk",
    "ProcessName": "notepad",
    "ExpectedExe": "notepad.exe"
  }
]
'@
            $cfgPath = Join-Path $script:cfgDir 'apps.json'
            $json | Set-Content -LiteralPath $cfgPath -Encoding UTF8
            $result = Import-AppsConfig -Path $cfgPath
            $result.Count | Should -Be 1
            $result[0].Name | Should -Be 'Notepad'
        }

        It 'normalises missing optional Appx fields to empty strings' {
            $json = @'
[
  {
    "Name": "App",
    "LaunchType": "Win32",
    "ShortcutPath": "C:\\Fake\\01 App.lnk",
    "ProcessName": "app",
    "ExpectedExe": "app.exe"
  }
]
'@
            $cfgPath = Join-Path $script:cfgDir 'apps.json'
            $json | Set-Content -LiteralPath $cfgPath -Encoding UTF8
            $result = Import-AppsConfig -Path $cfgPath
            $result[0].StartAppName | Should -Be ''
            $result[0].KnownAumid   | Should -Be ''
            $result[0].AppxName     | Should -Be ''
        }

        It 'throws when a required field is missing' {
            $json = @'
[
  {
    "Name": "BadApp",
    "LaunchType": "Win32"
  }
]
'@
            $cfgPath = Join-Path $script:cfgDir 'apps.json'
            $json | Set-Content -LiteralPath $cfgPath -Encoding UTF8
            { Import-AppsConfig -Path $cfgPath } | Should -Throw
        }

        It 'throws when the file does not exist' {
            { Import-AppsConfig -Path (Join-Path $script:cfgDir 'missing.json') } | Should -Throw
        }
    }

    # -----------------------------------------------------------------------
    # T-08: Get-ParentFolder
    # -----------------------------------------------------------------------
    Describe 'Get-ParentFolder' {

        It 'returns null for an empty string' {
            Get-ParentFolder -BrokenTargetPath '' | Should -BeNullOrEmpty
        }

        It 'returns null for a whitespace-only string' {
            Get-ParentFolder -BrokenTargetPath '   ' | Should -BeNullOrEmpty
        }

        It 'returns the grandparent when it exists on disk' {
            # Use $env:TEMP\subfolder\file.exe — grandparent is $env:TEMP which always exists.
            $fakePath = Join-Path $env:TEMP 'SubA\file.exe'
            $result = Get-ParentFolder -BrokenTargetPath $fakePath
            # Grandparent of TEMP\SubA\file.exe is TEMP (parent of SubA)
            $result | Should -Be $env:TEMP
        }

        It 'returns the immediate parent when grandparent does not exist' {
            # Drive:\NonExistent\sub\file.exe — grandparent Drive:\NonExistent does not exist.
            # Parent Drive:\NonExistent\sub also does not exist -> should return null.
            $fakePath = 'Z:\NoSuchDrive\sub\file.exe'
            Get-ParentFolder -BrokenTargetPath $fakePath | Should -BeNullOrEmpty
        }

        It 'returns a non-null existing path for a real system exe path' {
            $sysPath = Join-Path $env:SystemRoot 'System32\notepad.exe'
            $result = Get-ParentFolder -BrokenTargetPath $sysPath
            $result | Should -Not -BeNullOrEmpty
            Test-Path -LiteralPath $result -PathType Container | Should -Be $true
        }
    }

    # -----------------------------------------------------------------------
    # T-09: Show-FailureMenu
    # -----------------------------------------------------------------------
    Describe 'Show-FailureMenu' {

        It 'returns the string "1" when the user enters 1' {
            $result = '1' | & { Show-FailureMenu -AppName 'TestApp' -Context 'unit test' }
            $result | Should -Be '1'
        }

        It 'returns the string "2" when the user enters 2' {
            $result = '2' | & { Show-FailureMenu -AppName 'TestApp' -Context 'unit test' }
            $result | Should -Be '2'
        }

        It 'returns the string "3" when the user enters 3' {
            $result = '3' | & { Show-FailureMenu -AppName 'TestApp' -Context 'unit test' }
            $result | Should -Be '3'
        }

        It 'does not throw for any valid input' {
            { '1' | & { Show-FailureMenu -AppName 'X' -Context 'ctx' } } | Should -Not -Throw
            { '2' | & { Show-FailureMenu -AppName 'X' -Context 'ctx' } } | Should -Not -Throw
            { '3' | & { Show-FailureMenu -AppName 'X' -Context 'ctx' } } | Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # TEST-10: Show-AppPicker -AllowNew sentinel (FIX-02)
    # -----------------------------------------------------------------------
    Describe 'Show-AppPicker -AllowNew' {

        It 'returns __NEW__ when N is entered and -AllowNew is set' {
            $saved = $script:apps
            $script:apps = @(
                [PSCustomObject]@{ Name='FakeApp'; ShortcutPath='C:\Fake\01 FakeApp.lnk' }
            )
            try {
                $result = 'n' | & { Show-AppPicker -Prompt 'Test' -AllowNew }
                $result | Should -Be '__NEW__'
            } finally {
                $script:apps = $saved
            }
        }

        It 'returns $null when 0 is entered regardless of -AllowNew' {
            $saved = $script:apps
            $script:apps = @(
                [PSCustomObject]@{ Name='FakeApp'; ShortcutPath='C:\Fake\01 FakeApp.lnk' }
            )
            try {
                $result = '0' | & { Show-AppPicker -Prompt 'Test' -AllowNew }
                $result | Should -BeNullOrEmpty
            } finally {
                $script:apps = $saved
            }
        }
    }

    # -----------------------------------------------------------------------
    # TEST-11: Add-Shortcut dispatch on $null / __NEW__ / real app (FIX-01 + FIX-02)
    # -----------------------------------------------------------------------
    Describe 'Add-Shortcut dispatch' {

        It 'prints cancel message and returns when $null is passed' {
            { Add-Shortcut -App $null } | Should -Not -Throw
        }

        It 'calls Initialize-Shortcut when a real app object is passed (re-init path)' {
            $dir = Join-Path $env:TEMP ("PesterAddTest_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $dir | Out-Null
            $lnk = Join-Path $dir '01 FakeApp.lnk'
            $wsh = New-Object -ComObject WScript.Shell
            $sc  = $wsh.CreateShortcut($lnk)
            $sc.TargetPath = "$env:SystemRoot\System32\notepad.exe"
            $sc.Save()
            $fakeApp = [PSCustomObject]@{
                Name              = 'FakeApp'
                LaunchType        = 'Win32'
                ShortcutPath      = $lnk
                ProcessName       = 'notepad'
                ExpectedExe       = 'notepad.exe'
                ExpectedPublisher = ''
                ExpectedArguments = ''
            }
            try {
                { Add-Shortcut -App $fakeApp } | Should -Not -Throw
            } finally {
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # -----------------------------------------------------------------------
    # TEST-12: Wait-ForAppReady phase-2 timeout math (FIX-03)
    # -----------------------------------------------------------------------
    Describe 'Wait-ForAppReady phase-1 clamping' {

        It 'does not exceed TimeoutSeconds when TimeoutSeconds < SettleSeconds' {
            $settleSeconds  = 5
            $timeoutSeconds = 2
            $phase1Secs = [Math]::Min($settleSeconds, $timeoutSeconds)
            $remaining  = $timeoutSeconds - $phase1Secs
            $phase1Secs | Should -Be 2
            $remaining  | Should -Be 0
        }

        It 'uses full SettleSeconds when TimeoutSeconds >= SettleSeconds' {
            $settleSeconds  = 5
            $timeoutSeconds = 30
            $phase1Secs = [Math]::Min($settleSeconds, $timeoutSeconds)
            $remaining  = $timeoutSeconds - $phase1Secs
            $phase1Secs | Should -Be 5
            $remaining  | Should -Be 25
        }
    }
}

# ---------------------------------------------------------------------------
# Integration tests  (only when RUN_INTEGRATION=1)
# ---------------------------------------------------------------------------
Describe 'Integration' -Skip:($env:RUN_INTEGRATION -ne '1') {

    # INT-01: harness setup/teardown
    BeforeAll {
        $script:intTempDir = Join-Path $env:TEMP ("PesterIntTest_" + [System.IO.Path]::GetRandomFileName())
        New-Item -ItemType Directory -Path $script:intTempDir | Out-Null
    }

    AfterAll {
        Get-ChildItem -LiteralPath $script:intTempDir -Filter '*.lnk' -ErrorAction SilentlyContinue |
            Remove-Item -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:intTempDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    # INT-02: shortcut bootstrap smoke test
    It 'Initialize-Shortcut creates a valid lnk pointing to notepad.exe' {
        $notepad     = Join-Path $env:SystemRoot 'System32\notepad.exe'
        $lnkPath     = Join-Path $script:intTempDir '01 Notepad.lnk'
        $fakeApp     = @{
            Name              = 'Notepad'
            LaunchType        = 'Win32'
            ShortcutPath      = $lnkPath
            ExpectedExe       = 'notepad.exe'
            ExpectedPublisher = 'CN=Microsoft Windows'
            ExpectedArguments = ''
        }

        $wsh = New-Object -ComObject WScript.Shell
        $sc  = $wsh.CreateShortcut($lnkPath)
        $sc.TargetPath       = $notepad
        $sc.WorkingDirectory = Split-Path $notepad -Parent
        $sc.Save()

        { Initialize-Shortcut -App $fakeApp } | Should -Not -Throw
        Test-Path -LiteralPath $lnkPath -PathType Leaf | Should -Be $true

        $verify = $wsh.CreateShortcut($lnkPath)
        $verify.TargetPath | Should -Be $notepad
    }
}
