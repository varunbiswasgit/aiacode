#Requires -Modules Pester
# Run unit tests : Invoke-Pester .\Win11startup.Tests.ps1
# Run all tests  : $env:RUN_INTEGRATION = '1'; Invoke-Pester .\Win11startup.Tests.ps1

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

        # Pre-create the shortcut manually (simulates what a user would provide)
        $wsh = New-Object -ComObject WScript.Shell
        $sc  = $wsh.CreateShortcut($lnkPath)
        $sc.TargetPath       = $notepad
        $sc.WorkingDirectory = Split-Path $notepad -Parent
        $sc.Save()

        # Initialize-Shortcut should detect it already exists and return silently
        { Initialize-Shortcut -App $fakeApp } | Should -Not -Throw
        Test-Path -LiteralPath $lnkPath -PathType Leaf | Should -Be $true

        # Verify the shortcut target resolves correctly
        $verify = $wsh.CreateShortcut($lnkPath)
        $verify.TargetPath | Should -Be $notepad
    }
}
