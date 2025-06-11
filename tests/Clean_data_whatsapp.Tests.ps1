$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$scriptPath = Join-Path $here '..' 'Clean_data_whatsapp.ps1'

describe 'Clean_data_whatsapp.ps1' {
    beforeAll {
        $testDir = Join-Path $here 'tmp'
        if (Test-Path $testDir) { Remove-Item $testDir -Recurse -Force }
        New-Item -ItemType Directory -Path $testDir | Out-Null
    }
    afterAll {
        if (Test-Path $testDir) { Remove-Item $testDir -Recurse -Force }
    }

    it 'parses lines with and without AM/PM correctly' {
        $input = @"
[1/12/23, 9:45:00 PM] John Doe: Hello!
[1/12/23, 09:00:00] System: Server started
Continuation line
[1/12/23, 9:46:30 AM] Jane Smith: Morning!
"@
        $inputFile = Join-Path $testDir 'chat.txt'
        $outputFile = Join-Path $testDir 'output.txt'
        Set-Content -Path $inputFile -Value $input -Encoding UTF8
        $responses = @($inputFile, $outputFile)
        $index = 0
        Mock -CommandName Read-Host -MockWith { $responses[$index++] }
        . $scriptPath
        $output = Get-Content $outputFile -Encoding UTF8
        $expected = @(
            "1/12/23`t9:45:00 PM`tJohn Doe`tHello!",
            "1/12/23`t9:46:30 AM`tJane Smith`tMorning!"
        )
        $output | Should -BeExactly $expected
    }
}

