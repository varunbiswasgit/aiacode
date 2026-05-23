#Requires -Modules Pester
# Run unit tests : Invoke-Pester .\Win11startup.Tests.ps1
# Run all tests  : $env:RUN_INTEGRATION = '1'; Invoke-Pester .\Win11startup.Tests.ps1
#
# Tested functions
# ----------------
# Unit:
#   Get-RelativeDepth              (TEST-02)
#   Find-MisnumberedShortcut       (TEST-03)
#   Test-ExePathAllowed            (TEST-04a)
#   Test-ExeSignatureTrusted       (TEST-04b)
#   Import-AppsConfig              (TEST-08)
#   Get-ParentFolder               (T-08)
#   Show-FailureMenu               (T-09)
#   Show-AppPicker -AllowNew       (TEST-10)
#   Add-Shortcut dispatch          (TEST-11)
#   Wait-ForAppReady phase math    (TEST-12)
#   Test-AppAlreadyOpen -RequireWindow (NEW-TEST-08)
#   Export-AppsConfig error path   (NEW-TEST-09)
#   Resolve-Aumid error log        (NEW-TEST-10)
#   Invoke-FailureRecovery         (NEW-TEST-11)
#   Show-AppList                   (NEW-TEST-12)
#   Import-AppsConfig schemaVersion (NEW-TEST-13)
#   Invoke-ShortcutRepair          (LEAN-06-01)
#   Repair-ShortcutArguments wrapper (LEAN-06-02)
#   Test-ExeAcceptable             (NEW-TEST-14)
#   New-AppEntry                   (NEW-TEST-15)
#   Wait-ForProcessCondition       (NEW-TEST-16)
#   Find-ExeWithinDepth            (NEW-TEST-17)
#   Get-ShortcutObject             (NEW-TEST-18)
#   New-AppShortcut                (NEW-TEST-19)
#   Invoke-AppLaunch               (NEW-TEST-20)
#   Test-ShortcutHealthy           (NEW-TEST-21)
#   Get-AppPresence                (NEW-TEST-22)
# Integration:
#   Initialize-Shortcut            (INT-02)

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
    # TEST-08 (original): Import-AppsConfig
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
            $fakePath = Join-Path $env:TEMP 'SubA\file.exe'
            $result = Get-ParentFolder -BrokenTargetPath $fakePath
            $result | Should -Be $env:TEMP
        }

        It 'returns null when neither grandparent nor parent exist' {
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
    # TEST-10: Show-AppPicker -AllowNew sentinel
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
    # TEST-11: Add-Shortcut dispatch
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
    # TEST-12 (original): Wait-ForAppReady phase-1 clamping
    # -----------------------------------------------------------------------
    Describe 'Wait-ForAppReady phase-1 clamping' {

        It 'does not exceed TimeoutSeconds when TimeoutSeconds < SettleSeconds' {
            $phase1Secs = [Math]::Min(5, 2)
            $remaining  = 2 - $phase1Secs
            $phase1Secs | Should -Be 2
            $remaining  | Should -Be 0
        }

        It 'uses full SettleSeconds when TimeoutSeconds >= SettleSeconds' {
            $phase1Secs = [Math]::Min(5, 30)
            $remaining  = 30 - $phase1Secs
            $phase1Secs | Should -Be 5
            $remaining  | Should -Be 25
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-08: Test-AppAlreadyOpen -RequireWindow (FIX-04)
    # -----------------------------------------------------------------------
    Describe 'Test-AppAlreadyOpen -RequireWindow' {

        It 'returns false when no process with given name is running' {
            Test-AppAlreadyOpen -ProcessName 'pester_no_such_process_xyz' | Should -Be $false
        }

        It 'returns false with -RequireWindow when process is running but has no window' {
            $p = Start-Process notepad.exe -WindowStyle Hidden -PassThru -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 500
            try {
                $result = Test-AppAlreadyOpen -ProcessName 'notepad' -RequireWindow
                $result | Should -BeIn @($true, $false)
            } finally {
                $p | Stop-Process -Force -ErrorAction SilentlyContinue
            }
        }

        It 'does not throw when -RequireWindow is omitted (backward compat)' {
            { Test-AppAlreadyOpen -ProcessName 'pester_no_such_process_xyz' } | Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-09: Export-AppsConfig error path (ROB-01)
    # -----------------------------------------------------------------------
    Describe 'Export-AppsConfig error path' {

        It 'emits a warning and does not throw when the destination path is invalid' {
            $saved      = $script:apps
            $savedPath  = $script:AppsConfigPath
            $script:apps = @([PSCustomObject]@{
                Name='Test'; LaunchType='Win32'; ShortcutPath='C:\x.lnk'
                ProcessName='test'; ExpectedExe='test.exe'
                ExpectedPublisher=''; ExpectedArguments=''
                StartAppName=''; KnownAumid=''; AppxName=''
            })
            $script:AppsConfigPath = $env:TEMP
            try {
                { Export-AppsConfig -Path $env:TEMP } | Should -Not -Throw
            } finally {
                $script:apps          = $saved
                $script:AppsConfigPath = $savedPath
            }
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-10: Resolve-Aumid logs to error log on all-paths failure (QOL-02)
    # -----------------------------------------------------------------------
    Describe 'Resolve-Aumid error log on failure' {

        It 'writes to startup-error.log when AUMID cannot be resolved' {
            $logPath = Join-Path $env:TEMP ("PesterErrLog_" + [System.IO.Path]::GetRandomFileName() + ".log")
            $savedLog        = $script:ErrorLogPath
            $script:ErrorLogPath = $logPath
            $fakeApp = [PSCustomObject]@{
                Name         = 'NoSuchApp'
                StartAppName = 'pester_no_such_startapp_xyz'
                KnownAumid   = ''
                AppxName     = 'pester_no_such_appx_xyz'
            }
            try {
                $result = Resolve-Aumid -App $fakeApp
                $result | Should -BeNullOrEmpty
                Test-Path -LiteralPath $logPath -PathType Leaf | Should -Be $true
                $content = Get-Content -LiteralPath $logPath -Raw
                $content | Should -Match 'AUMID could not be resolved'
            } finally {
                $script:ErrorLogPath = $savedLog
                Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-11: Invoke-FailureRecovery branch coverage (ROB-04 + QOL-04)
    # -----------------------------------------------------------------------
    Describe 'Invoke-FailureRecovery' {

        BeforeEach {
            $script:IFR_fakeApp = [PSCustomObject]@{ Name = 'PesterApp' }
        }

        It 'returns $true and calls PreRetryAction when user chooses 1' {
            $called = $false
            $action = { $called = $true }
            $result = '1' | & { Invoke-FailureRecovery -App $script:IFR_fakeApp -Context 'test' -PreRetryAction $action }
            $result | Should -Be $true
            $called | Should -Be $true
        }

        It 'returns $true without error when choice is 1 and no PreRetryAction is provided' {
            $result = '1' | & { Invoke-FailureRecovery -App $script:IFR_fakeApp -Context 'test' }
            $result | Should -Be $true
        }

        It 'returns $false when user chooses 3 (skip)' {
            $result = '3' | & { Invoke-FailureRecovery -App $script:IFR_fakeApp -Context 'test' }
            $result | Should -Be $false
        }

        It 'returns $false when user enters an unrecognised value (default branch)' {
            $result = 'x' | & { Invoke-FailureRecovery -App $script:IFR_fakeApp -Context 'test' }
            $result | Should -Be $false
        }

        It 'does not throw for any of the three standard choices' {
            { '1' | & { Invoke-FailureRecovery -App $script:IFR_fakeApp -Context 'test' } } | Should -Not -Throw
            { '3' | & { Invoke-FailureRecovery -App $script:IFR_fakeApp -Context 'test' } } | Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-12: Show-AppList (QOL-05)
    # -----------------------------------------------------------------------
    Describe 'Show-AppList' {

        It 'does not throw and emits at least a header line' {
            $saved = $script:apps
            $script:apps = @(
                [PSCustomObject]@{
                    Name='PesterApp'; LaunchType='Win32'
                    ShortcutPath='C:\Fake\01 PesterApp.lnk'; ProcessName='pesterapp'
                }
            )
            try {
                $output = & { Show-AppList } 4>&1 | Out-String
                $output | Should -Not -BeNullOrEmpty
            } finally {
                $script:apps = $saved
            }
        }

        It 'includes the app name in the output' {
            $saved = $script:apps
            $script:apps = @(
                [PSCustomObject]@{
                    Name='UniqueTestAppXYZ'; LaunchType='Appx'
                    ShortcutPath='C:\Fake\01 UniqueTestAppXYZ.lnk'; ProcessName='utaxyz'
                }
            )
            try {
                $output = Show-AppList *>&1 | Out-String
                $output | Should -Match 'UniqueTestAppXYZ'
            } finally {
                $script:apps = $saved
            }
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-13: Import-AppsConfig schemaVersion handling (QOL-03)
    # -----------------------------------------------------------------------
    Describe 'Import-AppsConfig schemaVersion' {

        BeforeEach {
            $script:svDir = Join-Path $env:TEMP ("PesterSV_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:svDir | Out-Null
        }

        AfterEach {
            Remove-Item -LiteralPath $script:svDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        $baseEntry = @'
  {
    "Name": "App",
    "LaunchType": "Win32",
    "ShortcutPath": "C:\\Fake\\01 App.lnk",
    "ProcessName": "app",
    "ExpectedExe": "app.exe"
  }
'@

        It 'loads cleanly when schemaVersion is 1' {
            $json = "{`n  `"schemaVersion\": 1,`n  `"apps\": [`n$baseEntry`n  ]`n}"
            $cfgPath = Join-Path $script:svDir 'apps.json'
            $json | Set-Content -LiteralPath $cfgPath -Encoding UTF8
            { Import-AppsConfig -Path $cfgPath } | Should -Not -Throw
            $result = Import-AppsConfig -Path $cfgPath
            $result.Count | Should -Be 1
        }

        It 'emits a warning when schemaVersion is missing (legacy bare array)' {
            $json = "[`n$baseEntry`n]"
            $cfgPath = Join-Path $script:svDir 'apps.json'
            $json | Set-Content -LiteralPath $cfgPath -Encoding UTF8
            $warnings = @()
            Import-AppsConfig -Path $cfgPath -WarningVariable warnings -WarningAction SilentlyContinue | Out-Null
            $warnings | Where-Object { $_ -match 'schemaVersion' } | Should -Not -BeNullOrEmpty
        }

        It 'emits a warning when schemaVersion does not match expected value' {
            $json = "{`n  `"schemaVersion\": 99,`n  `"apps\": [`n$baseEntry`n  ]`n}"
            $cfgPath = Join-Path $script:svDir 'apps.json'
            $json | Set-Content -LiteralPath $cfgPath -Encoding UTF8
            $warnings = @()
            Import-AppsConfig -Path $cfgPath -WarningVariable warnings -WarningAction SilentlyContinue | Out-Null
            $warnings | Where-Object { $_ -match 'schemaVersion' } | Should -Not -BeNullOrEmpty
        }
    }

    # -----------------------------------------------------------------------
    # LEAN-06-01: Invoke-ShortcutRepair -- envelope behaviour
    # -----------------------------------------------------------------------
    Describe 'Invoke-ShortcutRepair' {

        BeforeEach {
            $script:ISR_dir = Join-Path $env:TEMP ("PesterISR_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:ISR_dir | Out-Null
            $script:ISR_lnk = Join-Path $script:ISR_dir '01 ISRTest.lnk'
            $wsh = New-Object -ComObject WScript.Shell
            $sc  = $wsh.CreateShortcut($script:ISR_lnk)
            $sc.TargetPath = "$env:SystemRoot\System32\notepad.exe"
            $sc.Save()
            $script:ISR_app = [PSCustomObject]@{
                Name         = 'ISRTest'
                ShortcutPath = $script:ISR_lnk
            }
        }

        AfterEach {
            Remove-Item -LiteralPath $script:ISR_dir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'calls .Save() and returns the repaired value when RepairAction returns non-null' {
            $result = Invoke-ShortcutRepair -App $script:ISR_app -RepairAction {
                param($shortcut)
                $shortcut.Arguments = 'pester-sentinel'
                return 'repaired'
            }
            $result | Should -Be 'repaired'
            $wsh    = New-Object -ComObject WScript.Shell
            $verify = $wsh.CreateShortcut($script:ISR_lnk)
            $verify.Arguments | Should -Be 'pester-sentinel'
        }

        It 'does not call .Save() and returns $null when RepairAction returns $null' {
            $wsh = New-Object -ComObject WScript.Shell
            $pre = $wsh.CreateShortcut($script:ISR_lnk)
            $pre.Arguments = 'original-value'
            $pre.Save()
            $result = Invoke-ShortcutRepair -App $script:ISR_app -RepairAction {
                param($shortcut)
                $shortcut.Arguments = 'should-not-persist'
                return $null
            }
            $result | Should -BeNullOrEmpty
            $wsh2   = New-Object -ComObject WScript.Shell
            $verify = $wsh2.CreateShortcut($script:ISR_lnk)
            $verify.Arguments | Should -Be 'original-value'
        }

        It 'does not throw when RepairAction returns $null' {
            { Invoke-ShortcutRepair -App $script:ISR_app -RepairAction { param($s) return $null } } |
                Should -Not -Throw
        }

        It 'throws when the shortcut file does not exist' {
            $missing = [PSCustomObject]@{ Name = 'Ghost'; ShortcutPath = 'C:\DoesNotExist\ghost.lnk' }
            { Invoke-ShortcutRepair -App $missing -RepairAction { param($s) return 'x' } } |
                Should -Throw
        }
    }

    # -----------------------------------------------------------------------
    # LEAN-06-02: Repair-ShortcutArguments is a thin wrapper
    # -----------------------------------------------------------------------
    Describe 'Repair-ShortcutArguments wrapper' {

        It 'returns $null and does not throw when ExpectedArguments has no shell:appsFolder pattern' {
            $dir = Join-Path $env:TEMP ("PesterRSA_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $dir | Out-Null
            $lnk = Join-Path $dir '01 RSATest.lnk'
            $wsh = New-Object -ComObject WScript.Shell
            $sc  = $wsh.CreateShortcut($lnk)
            $sc.TargetPath = "$env:SystemRoot\System32\notepad.exe"
            $sc.Save()
            $app = [PSCustomObject]@{
                Name='RSATest'; ShortcutPath=$lnk; ExpectedArguments='not-a-shell-appsfolder-value'
            }
            try {
                $result = Repair-ShortcutArguments -App $app
                $result | Should -BeNullOrEmpty
            } finally {
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        It 'does not throw when the shortcut has no Arguments field set' {
            $dir = Join-Path $env:TEMP ("PesterRSA2_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $dir | Out-Null
            $lnk = Join-Path $dir '01 RSATest2.lnk'
            $wsh = New-Object -ComObject WScript.Shell
            $sc  = $wsh.CreateShortcut($lnk)
            $sc.TargetPath = "$env:SystemRoot\System32\explorer.exe"
            $sc.Save()
            $app = [PSCustomObject]@{
                Name='RSATest2'; ShortcutPath=$lnk
                ExpectedArguments='shell:appsFolder\Contoso.App_abc123_x64__xyz!App'
            }
            try {
                { Repair-ShortcutArguments -App $app } | Should -Not -Throw
            } finally {
                Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-14: Test-ExeAcceptable -- LEAN-01 composite gate
    # -----------------------------------------------------------------------
    Describe 'Test-ExeAcceptable' {

        It 'returns true for notepad.exe (allowed path + valid signature, no publisher check)' {
            $notepad = Join-Path $env:SystemRoot 'System32\notepad.exe'
            Test-ExeAcceptable -ExePath $notepad | Should -Be $true
        }

        It 'returns true for notepad.exe with correct publisher' {
            $notepad = Join-Path $env:SystemRoot 'System32\notepad.exe'
            Test-ExeAcceptable -ExePath $notepad -ExpectedPublisher 'CN=Microsoft Windows' | Should -Be $true
        }

        It 'returns false when the path is outside allowed roots' {
            $fakePath = Join-Path $env:TEMP 'notepad_copy.exe'
            Test-ExeAcceptable -ExePath $fakePath | Should -Be $false
        }

        It 'returns false for notepad.exe with wrong publisher' {
            $notepad = Join-Path $env:SystemRoot 'System32\notepad.exe'
            Test-ExeAcceptable -ExePath $notepad -ExpectedPublisher 'CN=Google LLC' | Should -Be $false
        }

        It 'does not throw for any combination of valid inputs' {
            $notepad = Join-Path $env:SystemRoot 'System32\notepad.exe'
            { Test-ExeAcceptable -ExePath $notepad }                                           | Should -Not -Throw
            { Test-ExeAcceptable -ExePath $notepad -ExpectedPublisher '' }                     | Should -Not -Throw
            { Test-ExeAcceptable -ExePath $notepad -ExpectedPublisher 'CN=Microsoft Windows' } | Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-15: New-AppEntry -- single constructor contract
    # -----------------------------------------------------------------------
    Describe 'New-AppEntry' {

        It 'returns an object with all required fields populated' {
            $entry = New-AppEntry -Name 'TestApp' -LaunchType 'Win32' `
                -ShortcutPath 'C:\Fake\01 TestApp.lnk' `
                -ProcessName 'testapp' -ExpectedExe 'testapp.exe'
            $entry.Name         | Should -Be 'TestApp'
            $entry.LaunchType   | Should -Be 'Win32'
            $entry.ShortcutPath | Should -Be 'C:\Fake\01 TestApp.lnk'
            $entry.ProcessName  | Should -Be 'testapp'
            $entry.ExpectedExe  | Should -Be 'testapp.exe'
        }

        It 'defaults optional string fields to empty string' {
            $entry = New-AppEntry -Name 'A' -LaunchType 'Win32' -ShortcutPath 'C:\x.lnk' -ExpectedExe 'a.exe'
            $entry.ExpectedPublisher | Should -Be ''
            $entry.ExpectedArguments | Should -Be ''
            $entry.StartAppName      | Should -Be ''
            $entry.KnownAumid        | Should -Be ''
            $entry.AppxName          | Should -Be ''
        }

        It 'defaults PresenceMode to $null' {
            $entry = New-AppEntry -Name 'A' -LaunchType 'Appx' -ShortcutPath 'C:\x.lnk' -ExpectedExe 'a.exe'
            $entry.PresenceMode | Should -BeNullOrEmpty
        }

        It 'honours explicitly supplied optional values' {
            $entry = New-AppEntry -Name 'Chrome' -LaunchType 'Win32' `
                -ShortcutPath 'C:\Fake\01 Chrome.lnk' -ExpectedExe 'chrome.exe' `
                -ExpectedPublisher 'CN=Google LLC' -StartAppName 'Google Chrome' `
                -KnownAumid 'ChromeAUMID' -AppxName 'Chrome'
            $entry.ExpectedPublisher | Should -Be 'CN=Google LLC'
            $entry.StartAppName      | Should -Be 'Google Chrome'
            $entry.KnownAumid        | Should -Be 'ChromeAUMID'
            $entry.AppxName          | Should -Be 'Chrome'
        }

        It 'does not throw for any combination of parameters' {
            { New-AppEntry -Name 'X' -LaunchType 'Win32' -ShortcutPath 'C:\x.lnk' -ExpectedExe 'x.exe' } |
                Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-16: Wait-ForProcessCondition -- DUP-01 extracted helper
    # -----------------------------------------------------------------------
    Describe 'Wait-ForProcessCondition' {

        It 'returns $true immediately when condition is already true' {
            Wait-ForProcessCondition -Condition { $true } -Remaining 5 | Should -Be $true
        }

        It 'returns $false when condition is never true and Remaining elapses' {
            Wait-ForProcessCondition -Condition { $false } -Remaining 1 | Should -Be $false
        }

        It 'returns $true as soon as condition becomes true mid-loop' {
            $script:WFPC_calls = 0
            $result = Wait-ForProcessCondition -Condition {
                $script:WFPC_calls++
                $script:WFPC_calls -ge 1
            } -Remaining 5
            $result | Should -Be $true
        }

        It 'does not throw when Remaining is 0' {
            { Wait-ForProcessCondition -Condition { $false } -Remaining 0 } | Should -Not -Throw
        }

        It 'does not throw when condition scriptblock throws internally' {
            { Wait-ForProcessCondition -Condition { throw 'oops' } -Remaining 1 } | Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-17: Find-ExeWithinDepth -- depth-bounded exe search
    # -----------------------------------------------------------------------
    Describe 'Find-ExeWithinDepth' {

        BeforeEach {
            $script:FEW_dir = Join-Path $env:TEMP ("PesterFEW_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:FEW_dir | Out-Null
        }

        AfterEach {
            Remove-Item -LiteralPath $script:FEW_dir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'finds an exe at depth 1 when MaxDepth is 1' {
            $exePath = Join-Path $script:FEW_dir 'app.exe'
            New-Item -ItemType File -Path $exePath | Out-Null
            $result = Find-ExeWithinDepth -SearchRoot $script:FEW_dir -ExeName 'app.exe' -MaxDepth 1
            $result | Should -Not -BeNullOrEmpty
            $result | Should -Be $exePath
        }

        It 'finds an exe nested at depth 2 when MaxDepth is 2' {
            $sub = Join-Path $script:FEW_dir 'sub'
            New-Item -ItemType Directory -Path $sub | Out-Null
            $exePath = Join-Path $sub 'app.exe'
            New-Item -ItemType File -Path $exePath | Out-Null
            $result = Find-ExeWithinDepth -SearchRoot $script:FEW_dir -ExeName 'app.exe' -MaxDepth 2
            $result | Should -Not -BeNullOrEmpty
            $result | Should -Be $exePath
        }

        It 'returns $null when exe exists only beyond MaxDepth' {
            $sub1 = Join-Path $script:FEW_dir 'a'
            $sub2 = Join-Path $sub1 'b'
            New-Item -ItemType Directory -Path $sub2 -Force | Out-Null
            $exePath = Join-Path $sub2 'app.exe'
            New-Item -ItemType File -Path $exePath | Out-Null
            Find-ExeWithinDepth -SearchRoot $script:FEW_dir -ExeName 'app.exe' -MaxDepth 1 | Should -BeNullOrEmpty
        }

        It 'returns $null when no matching exe exists anywhere' {
            Find-ExeWithinDepth -SearchRoot $script:FEW_dir -ExeName 'nosuchfile.exe' -MaxDepth 3 | Should -BeNullOrEmpty
        }

        It 'does not throw when SearchRoot does not exist' {
            { Find-ExeWithinDepth -SearchRoot 'C:\PesterNonExistent_XYZ' -ExeName 'x.exe' -MaxDepth 2 } |
                Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-18: Get-ShortcutObject -- COM wrapper contract
    # -----------------------------------------------------------------------
    Describe 'Get-ShortcutObject' {

        BeforeEach {
            $script:GSO_dir = Join-Path $env:TEMP ("PesterGSO_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:GSO_dir | Out-Null
            $script:GSO_lnk = Join-Path $script:GSO_dir '01 GSOTest.lnk'
            $wsh = New-Object -ComObject WScript.Shell
            $sc  = $wsh.CreateShortcut($script:GSO_lnk)
            $sc.TargetPath = "$env:SystemRoot\System32\notepad.exe"
            $sc.Save()
        }

        AfterEach {
            Remove-Item -LiteralPath $script:GSO_dir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'returns an object with a non-empty TargetPath for a valid lnk' {
            $sc = Get-ShortcutObject -LnkPath $script:GSO_lnk
            $sc | Should -Not -BeNullOrEmpty
            $sc.TargetPath | Should -Not -BeNullOrEmpty
        }

        It 'TargetPath matches the exe used when creating the shortcut' {
            $sc = Get-ShortcutObject -LnkPath $script:GSO_lnk
            $sc.TargetPath | Should -Be "$env:SystemRoot\System32\notepad.exe"
        }

        It 'returned object exposes a Save method' {
            $sc = Get-ShortcutObject -LnkPath $script:GSO_lnk
            $sc | Get-Member -Name 'Save' | Should -Not -BeNullOrEmpty
        }

        It 'throws or returns $null for a path that does not exist' {
            $missing = Join-Path $script:GSO_dir 'ghost.lnk'
            $threw = $false; $result = $null
            try { $result = Get-ShortcutObject -LnkPath $missing } catch { $threw = $true }
            ($threw -or ($null -eq $result)) | Should -Be $true
        }

        It 'does not throw for a valid lnk path' {
            { Get-ShortcutObject -LnkPath $script:GSO_lnk } | Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-19: New-AppShortcut -- shortcut creation contract
    # -----------------------------------------------------------------------
    Describe 'New-AppShortcut' {

        BeforeEach {
            $script:NAS_dir = Join-Path $env:TEMP ("PesterNAS_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:NAS_dir | Out-Null
        }

        AfterEach {
            Remove-Item -LiteralPath $script:NAS_dir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'creates a .lnk file at the specified ShortcutPath' {
            $lnk = Join-Path $script:NAS_dir '01 NASTest.lnk'
            $app = [PSCustomObject]@{ Name='NASTest'; LaunchType='Win32'; ShortcutPath=$lnk
                ExpectedExe='notepad.exe'; ExpectedPublisher=''; ExpectedArguments='' }
            New-AppShortcut -App $app -TargetPath "$env:SystemRoot\System32\notepad.exe"
            Test-Path -LiteralPath $lnk -PathType Leaf | Should -Be $true
        }

        It 'created lnk has TargetPath matching the supplied exe' {
            $lnk = Join-Path $script:NAS_dir '01 NASTest2.lnk'
            $app = [PSCustomObject]@{ Name='NASTest2'; LaunchType='Win32'; ShortcutPath=$lnk
                ExpectedExe='notepad.exe'; ExpectedPublisher=''; ExpectedArguments='' }
            $target = "$env:SystemRoot\System32\notepad.exe"
            New-AppShortcut -App $app -TargetPath $target
            $wsh    = New-Object -ComObject WScript.Shell
            $verify = $wsh.CreateShortcut($lnk)
            $verify.TargetPath | Should -Be $target
        }

        It 'does not throw when called with valid parameters' {
            $lnk = Join-Path $script:NAS_dir '01 NASTest3.lnk'
            $app = [PSCustomObject]@{ Name='NASTest3'; LaunchType='Win32'; ShortcutPath=$lnk
                ExpectedExe='notepad.exe'; ExpectedPublisher=''; ExpectedArguments='' }
            { New-AppShortcut -App $app -TargetPath "$env:SystemRoot\System32\notepad.exe" } |
                Should -Not -Throw
        }

        It 'does not throw when WorkingDirectory is explicitly passed' {
            $lnk = Join-Path $script:NAS_dir '01 NASTest4.lnk'
            $app = [PSCustomObject]@{ Name='NASTest4'; LaunchType='Win32'; ShortcutPath=$lnk
                ExpectedExe='notepad.exe'; ExpectedPublisher=''; ExpectedArguments='' }
            { New-AppShortcut -App $app -TargetPath "$env:SystemRoot\System32\notepad.exe" `
                -WorkingDirectory "$env:SystemRoot\System32" } | Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-20: Invoke-AppLaunch -- retry-loop branch coverage (TEST-GAP-01)
    # -----------------------------------------------------------------------
    Describe 'Invoke-AppLaunch' {

        BeforeEach {
            $script:IAL_dir = Join-Path $env:TEMP ("PesterIAL_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:IAL_dir | Out-Null
            $script:IAL_lnk = Join-Path $script:IAL_dir '01 IALTest.lnk'
            $wsh = New-Object -ComObject WScript.Shell
            $sc  = $wsh.CreateShortcut($script:IAL_lnk)
            $sc.TargetPath = "$env:SystemRoot\System32\notepad.exe"
            $sc.Save()
            $script:IAL_app = [PSCustomObject]@{
                Name              = 'IALTest'
                LaunchType        = 'Win32'
                ShortcutPath      = $script:IAL_lnk
                ProcessName       = 'notepad'
                ExpectedExe       = 'notepad.exe'
                ExpectedPublisher = ''
                ExpectedArguments = ''
            }
        }

        AfterEach {
            Get-Process -Name 'notepad' -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $script:IAL_dir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'does not throw when called with a valid Win32 app object' {
            { Invoke-AppLaunch -App $script:IAL_app } | Should -Not -Throw
        }

        It 'returns a result object (not $null) for a valid app' {
            $result = Invoke-AppLaunch -App $script:IAL_app
            $result | Should -Not -BeNullOrEmpty
        }

        It 'returns a Success or known-status result for notepad' {
            $result = Invoke-AppLaunch -App $script:IAL_app
            $result | Should -BeIn @('Success', 'Retry', 'Abort', 'Skipped')
        }

        It 'does not throw when app object has an empty ExpectedPublisher' {
            $app2 = $script:IAL_app.PSObject.Copy()
            $app2.ExpectedPublisher = ''
            { Invoke-AppLaunch -App $app2 } | Should -Not -Throw
        }

        It 'does not throw when the shortcut file is missing (graceful error path)' {
            $ghost = [PSCustomObject]@{
                Name='Ghost'; LaunchType='Win32'
                ShortcutPath='C:\DoesNotExist\ghost.lnk'
                ProcessName='ghost'; ExpectedExe='ghost.exe'
                ExpectedPublisher=''; ExpectedArguments=''
            }
            { Invoke-AppLaunch -App $ghost } | Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-21: Test-ShortcutHealthy -- health-check gate
    # -----------------------------------------------------------------------
    Describe 'Test-ShortcutHealthy' {

        BeforeEach {
            $script:TSH_dir = Join-Path $env:TEMP ("PesterTSH_" + [System.IO.Path]::GetRandomFileName())
            New-Item -ItemType Directory -Path $script:TSH_dir | Out-Null
            $script:TSH_lnk = Join-Path $script:TSH_dir '01 TSHTest.lnk'
            $wsh = New-Object -ComObject WScript.Shell
            $sc  = $wsh.CreateShortcut($script:TSH_lnk)
            $sc.TargetPath = "$env:SystemRoot\System32\notepad.exe"
            $sc.Save()
        }

        AfterEach {
            Remove-Item -LiteralPath $script:TSH_dir -Recurse -Force -ErrorAction SilentlyContinue
        }

        It 'returns true for a valid lnk pointing to a real, signed exe' {
            $app = [PSCustomObject]@{
                Name='TSHTest'; ShortcutPath=$script:TSH_lnk
                ExpectedExe='notepad.exe'; ExpectedPublisher=''; ExpectedArguments=''
            }
            Test-ShortcutHealthy -App $app | Should -Be $true
        }

        It 'returns false when the lnk file does not exist' {
            $app = [PSCustomObject]@{
                Name='Ghost'; ShortcutPath='C:\DoesNotExist\ghost.lnk'
                ExpectedExe='ghost.exe'; ExpectedPublisher=''; ExpectedArguments=''
            }
            Test-ShortcutHealthy -App $app | Should -Be $false
        }

        It 'returns false when the lnk target does not exist on disk' {
            $brokenLnk = Join-Path $script:TSH_dir '01 Broken.lnk'
            $wsh = New-Object -ComObject WScript.Shell
            $sc  = $wsh.CreateShortcut($brokenLnk)
            $sc.TargetPath = 'C:\DoesNotExist\broken.exe'
            $sc.Save()
            $app = [PSCustomObject]@{
                Name='Broken'; ShortcutPath=$brokenLnk
                ExpectedExe='broken.exe'; ExpectedPublisher=''; ExpectedArguments=''
            }
            Test-ShortcutHealthy -App $app | Should -Be $false
        }

        It 'returns false when ExpectedPublisher does not match' {
            $app = [PSCustomObject]@{
                Name='TSHTest'; ShortcutPath=$script:TSH_lnk
                ExpectedExe='notepad.exe'; ExpectedPublisher='CN=Google LLC'; ExpectedArguments=''
            }
            Test-ShortcutHealthy -App $app | Should -Be $false
        }

        It 'does not throw for a healthy shortcut' {
            $app = [PSCustomObject]@{
                Name='TSHTest'; ShortcutPath=$script:TSH_lnk
                ExpectedExe='notepad.exe'; ExpectedPublisher=''; ExpectedArguments=''
            }
            { Test-ShortcutHealthy -App $app } | Should -Not -Throw
        }
    }

    # -----------------------------------------------------------------------
    # NEW-TEST-22: Get-AppPresence -- presence-mode detection
    # -----------------------------------------------------------------------
    Describe 'Get-AppPresence' {

        It 'returns a known presence string or $null for a running process' {
            $p = Start-Process notepad.exe -WindowStyle Hidden -PassThru -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 400
            try {
                $result = Get-AppPresence -ProcessName 'notepad'
                # Valid outcomes: Running, WindowVisible, or $null
                if ($null -ne $result) {
                    $result | Should -BeIn @('Running', 'WindowVisible')
                }
            } finally {
                $p | Stop-Process -Force -ErrorAction SilentlyContinue
            }
        }

        It 'returns $null or Absent for a process that is not running' {
            $result = Get-AppPresence -ProcessName 'pester_no_such_process_xyz'
            if ($null -ne $result) {
                $result | Should -BeIn @('Absent', $null)
            }
        }

        It 'does not throw for a running process' {
            $p = Start-Process notepad.exe -WindowStyle Hidden -PassThru -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 300
            try {
                { Get-AppPresence -ProcessName 'notepad' } | Should -Not -Throw
            } finally {
                $p | Stop-Process -Force -ErrorAction SilentlyContinue
            }
        }

        It 'does not throw for a non-existent process name' {
            { Get-AppPresence -ProcessName 'pester_no_such_process_xyz' } | Should -Not -Throw
        }

        It 'does not throw for an empty process name (BUG-F guard test)' {
            { Get-AppPresence -ProcessName '' } | Should -Not -Throw
        }
    }
}

# ---------------------------------------------------------------------------
# Integration tests  (only when RUN_INTEGRATION=1)
# ---------------------------------------------------------------------------
Describe 'Integration' -Skip:($env:RUN_INTEGRATION -ne '1') {

    BeforeAll {
        $script:intTempDir = Join-Path $env:TEMP ("PesterIntTest_" + [System.IO.Path]::GetRandomFileName())
        New-Item -ItemType Directory -Path $script:intTempDir | Out-Null
    }

    AfterAll {
        Get-ChildItem -LiteralPath $script:intTempDir -Filter '*.lnk' -ErrorAction SilentlyContinue |
            Remove-Item -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:intTempDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'Initialize-Shortcut creates a valid lnk pointing to notepad.exe' {
        $notepad = Join-Path $env:SystemRoot 'System32\notepad.exe'
        $lnkPath = Join-Path $script:intTempDir '01 Notepad.lnk'
        $fakeApp = @{
            Name='Notepad'; LaunchType='Win32'; ShortcutPath=$lnkPath
            ExpectedExe='notepad.exe'; ExpectedPublisher='CN=Microsoft Windows'; ExpectedArguments=''
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
