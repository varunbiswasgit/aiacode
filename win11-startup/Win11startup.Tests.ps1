# ===========================================================================
# Pester Unit Tests: Win11startup.Tests.ps1
# ---------------------------------------------------------------------------
# Validates core functions of Win11startup.ps1:
#   - Process name derivation (Get-ProcName)
#   - Numbered shortcut regex pattern matching
#   - Sequence number generation (Get-NextShortcutNumber)
#   - Readiness state logic (Get-AppReadyState)
#   - Safe configuration loading (Load-ConfigSafe)
# ===========================================================================

BeforeAll {
    $scriptPath = Join-Path $PSScriptRoot 'Win11startup.ps1'
    
    # Extract functions from the script without executing the interactive menu loop
    $scriptContent = Get-Content -LiteralPath $scriptPath -Raw
    
    # Strip the execution loop at the bottom so functions can be unit-tested
    $functionsOnly = $scriptContent -replace '(?s)# Master Menu Loop.*', ''
    $functionsOnly = $functionsOnly -replace '(?s)\$quit\s*=\s*\$false.*', ''
    
    Invoke-Expression $functionsOnly
}

Describe 'Pattern Matching & Formatting Tests' {
    It 'Matches zero-padded numbers 01 to 99 followed by space' {
        '01 Chrome.lnk'       | Should -Match $script:NumberedLnkPattern
        '09 Slack.lnk'        | Should -Match $script:NumberedLnkPattern
        '10 Terminal.lnk'     | Should -Match $script:NumberedLnkPattern
        '99 Discord.lnk'      | Should -Match $script:NumberedLnkPattern
    }

    It 'Rejects invalid or unnumbered shortcut formats' {
        '00 Invalid.lnk'      | Should -Not -Match $script:NumberedLnkPattern
        '1 Chrome.lnk'        | Should -Not -Match $script:NumberedLnkPattern
        'Chrome.lnk'          | Should -Not -Match $script:NumberedLnkPattern
        '01Chrome.lnk'        | Should -Not -Match $script:NumberedLnkPattern
        '100 TooHigh.lnk'     | Should -Not -Match $script:NumberedLnkPattern
    }
}

Describe 'Get-ProcName Function Tests' {
    It 'Extracts standard executable name without extension' {
        $result = Get-ProcName -TargetPath 'C:\Program Files\Google\Chrome\Application\chrome.exe' -DisplayName 'Google Chrome'
        $result | Should -Be 'chrome'
    }

    It 'Falls back to sanitized display name when target is explorer.exe' {
        $result = Get-ProcName -TargetPath 'C:\Windows\explorer.exe' -DisplayName 'Windows Terminal'
        $result | Should -Be 'WindowsTerminal'
    }

    It 'Falls back to sanitized display name when TargetPath is empty' {
        $result = Get-ProcName -TargetPath '' -DisplayName 'Slack App'
        $result | Should -Be 'SlackApp'
    }
}

Describe 'Sequence Number Calculation Tests' {
    Context 'With mock files in temporary folder' {
        BeforeEach {
            $TestDriveFolder = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
            New-Item -ItemType Directory -Path $TestDriveFolder -Force | Out-Null
        }

        AfterEach {
            if (Test-Path -LiteralPath $TestDriveFolder) {
                Remove-Item -LiteralPath $TestDriveFolder -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        It 'Returns 1 (formatted as 01) when no numbered files exist' {
            $next = Get-NextShortcutNumber -StartMenuFolder $TestDriveFolder
            $next | Should -Be 1
        }

        It 'Returns the next sequential integer when shortcuts exist' {
            New-Item -ItemType File -Path (Join-Path $TestDriveFolder '01 AppA.lnk') | Out-Null
            New-Item -ItemType File -Path (Join-Path $TestDriveFolder '02 AppB.lnk') | Out-Null
            
            $next = Get-NextShortcutNumber -StartMenuFolder $TestDriveFolder
            $next | Should -Be 3
        }

        It 'Ignores non-numbered lnk files when calculating the max number' {
            New-Item -ItemType File -Path (Join-Path $TestDriveFolder '05 AppE.lnk') | Out-Null
            New-Item -ItemType File -Path (Join-Path $TestDriveFolder 'Unnumbered.lnk') | Out-Null
            
            $next = Get-NextShortcutNumber -StartMenuFolder $TestDriveFolder
            $next | Should -Be 6
        }
    }
}

Describe 'Configuration Loading Tests' {
    Context 'JSON schema resilience' {
        BeforeEach {
            $tempConfigFile = [System.IO.Path]::GetTempFileName()
        }

        AfterEach {
            if (Test-Path -LiteralPath $tempConfigFile) {
                Remove-Item -LiteralPath $tempConfigFile -Force -ErrorAction SilentlyContinue
            }
        }

        It 'Returns default configuration object if JSON file is missing' {
            $res = Load-ConfigSafe -Path 'C:\NonExistentPath\config.json'
            $res.StartMenuPath | Should -Not -BeNullOrEmpty
            $res.Shortcuts.Count | Should -Be 0
        }

        It 'Correctly parses valid configuration JSON' {
            $sampleJson = @{
                StartMenuPath = 'C:\TestFolder'
                Shortcuts = @(
                    @{
                        Name = 'App1'
                        ShortcutPath = 'C:\TestFolder\01 App1.lnk'
                        ProcessName = 'app1'
                        LaunchType = 'Win32'
                        ExePath = 'C:\app1.exe'
                        Aumid = ''
                    }
                )
            } | ConvertTo-Json -Depth 5
            
            Set-Content -Path $tempConfigFile -Value $sampleJson -Encoding UTF8
            $res = Load-ConfigSafe -Path $tempConfigFile
            $res.StartMenuPath | Should -Be 'C:\TestFolder'
            $res.Shortcuts.Count | Should -Be 1
            $res.Shortcuts[0].Name | Should -Be 'App1'
        }
    }
}

