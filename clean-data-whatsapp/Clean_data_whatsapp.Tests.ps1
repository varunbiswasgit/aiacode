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

    function Invoke-Script($responses) {
        $index = 0
        Mock -CommandName Read-Host -MockWith { $responses[$script:index++] }
        . $scriptPath
    }

    it 'parses lines with and without AM/PM correctly (TSV)' {
        $input = @"
[1/12/23, 9:45:00 PM] John Doe: Hello!
[1/12/23, 09:00:00] System: Server started
Continuation line
[1/12/23, 9:46:30 AM] Jane Smith: Morning!
"@
        $inputFile  = Join-Path $testDir 'chat.txt'
        $outputFile = Join-Path $testDir 'output.txt'
        Set-Content -Path $inputFile -Value $input -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'T', '', '', '')
        $output = Get-Content $outputFile -Encoding UTF8
        $expected = @(
            "1/12/23`t9:45:00 PM`tJohn Doe`tHello!",
            "1/12/23`t9:00:00`tSystem`tServer started Continuation line",
            "1/12/23`t9:46:30 AM`tJane Smith`tMorning!"
        )
        $output | Should -BeExactly $expected
    }

    it 'strips surrounding quotes from file paths via Clean-Path' {
        $inputFile  = Join-Path $testDir 'chat_quotes.txt'
        $outputFile = Join-Path $testDir 'output_quotes.txt'
        Set-Content -Path $inputFile -Value '[1/5/23, 8:00:00 AM] Alice: Hi there!' -Encoding UTF8
        $quotedInput  = '"' + $inputFile  + '"'
        $quotedOutput = '"' + $outputFile + '"'
        Invoke-Script @($quotedInput, $quotedOutput, 'T', '', '', '')
        $outputFile | Should -Exist
        (Get-Content $outputFile -Encoding UTF8) | Should -Be "1/5/23`t8:00:00 AM`tAlice`tHi there!"
    }

    it 'removes LRM and replaces non-standard Unicode spaces' {
        $inputFile  = Join-Path $testDir 'chat_unicode.txt'
        $outputFile = Join-Path $testDir 'output_unicode.txt'
        $rawLine = "[1/6/23, 9:00:00`u{202F}AM] Bob: Good day"
        Set-Content -Path $inputFile -Value $rawLine -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'T', '', '', '')
        (Get-Content $outputFile -Encoding UTF8) | Should -Be "1/6/23`t9:00:00 AM`tBob`tGood day"
    }

    it 'preserves colons inside the message body in Column D' {
        $inputFile  = Join-Path $testDir 'chat_colon.txt'
        $outputFile = Join-Path $testDir 'output_colon.txt'
        Set-Content -Path $inputFile -Value '[1/7/23, 10:00:00 AM] John: Note: see below for details' -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'T', '', '', '')
        (Get-Content $outputFile -Encoding UTF8) | Should -Be "1/7/23`t10:00:00 AM`tJohn`tNote: see below for details"
    }

    it 'exits gracefully when the input file does not exist' {
        $missingFile = Join-Path $testDir 'does_not_exist.txt'
        $outputFile  = Join-Path $testDir 'output_missing.txt'
        { Invoke-Script @($missingFile, $outputFile, 'T', '', '', '') } | Should -Not -Throw
        $outputFile | Should -Not -Exist
    }

    it 'creates an empty output file when input file has no content' {
        $inputFile  = Join-Path $testDir 'chat_empty.txt'
        $outputFile = Join-Path $testDir 'output_empty.txt'
        Set-Content -Path $inputFile -Value '' -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'T', '', '', '')
        $outputFile | Should -Exist
        (Get-Content $outputFile -Encoding UTF8) | Should -BeNullOrEmpty
    }

    it 'appends continuation lines to the preceding message in Column D' {
        $inputFile  = Join-Path $testDir 'chat_continuation.txt'
        $outputFile = Join-Path $testDir 'output_continuation.txt'
        $input = @"
[1/8/23, 11:00:00 AM] Alice: First line of message
Second line of message
Third line of message
[1/8/23, 11:01:00 AM] Bob: Reply
"@
        Set-Content -Path $inputFile -Value $input -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'T', '', '', '')
        $output = Get-Content $outputFile -Encoding UTF8
        $output[0] | Should -Be "1/8/23`t11:00:00 AM`tAlice`tFirst line of message Second line of message Third line of message"
        $output[1] | Should -Be "1/8/23`t11:01:00 AM`tBob`tReply"
    }

    it 'parses DD/MM/YYYY timestamp format and normalises to M/D/YY' {
        $inputFile  = Join-Path $testDir 'chat_ddmmyyyy.txt'
        $outputFile = Join-Path $testDir 'output_ddmmyyyy.txt'
        Set-Content -Path $inputFile -Value '[25/12/2023, 10:00:00 AM] Alice: Happy Christmas!' -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'T', '', '', '')
        (Get-Content $outputFile -Encoding UTF8) | Should -Be "12/25/23`t10:00:00 AM`tAlice`tHappy Christmas!"
    }

    it 'parses YYYY-MM-DD ISO 8601 timestamp format and normalises to M/D/YY' {
        $inputFile  = Join-Path $testDir 'chat_iso.txt'
        $outputFile = Join-Path $testDir 'output_iso.txt'
        Set-Content -Path $inputFile -Value '[2023-06-15, 14:30:00] Bob: ISO format message' -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'T', '', '', '')
        (Get-Content $outputFile -Encoding UTF8) | Should -Be "6/15/23`t14:30:00`tBob`tISO format message"
    }

    it 'produces comma-separated output when CSV format is chosen' {
        $inputFile  = Join-Path $testDir 'chat_csv.txt'
        $outputFile = Join-Path $testDir 'output_csv.txt'
        Set-Content -Path $inputFile -Value '[1/9/23, 9:00:00 AM] Alice: Hello world' -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'C', '', '', '')
        (Get-Content $outputFile -Encoding UTF8) | Should -Be '1/9/23,9:00:00 AM,Alice,Hello world'
    }

    it 'wraps CSV fields in quotes when the message contains a comma' {
        $inputFile  = Join-Path $testDir 'chat_csv_comma.txt'
        $outputFile = Join-Path $testDir 'output_csv_comma.txt'
        Set-Content -Path $inputFile -Value '[1/10/23, 9:00:00 AM] Alice: Yes, please' -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'C', '', '', '')
        (Get-Content $outputFile -Encoding UTF8) | Should -Be '1/10/23,9:00:00 AM,Alice,"Yes, please"'
    }

    it 'filters output to only include messages from the specified sender' {
        $inputFile  = Join-Path $testDir 'chat_filter.txt'
        $outputFile = Join-Path $testDir 'output_filter.txt'
        $input = @"
[1/11/23, 8:00:00 AM] Alice: Good morning
[1/11/23, 8:01:00 AM] Bob: Hey Alice
[1/11/23, 8:02:00 AM] Alice: How are you?
"@
        Set-Content -Path $inputFile -Value $input -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'T', 'Alice', '', '')
        $output = Get-Content $outputFile -Encoding UTF8
        $output.Count | Should -Be 2
        $output[0] | Should -Match 'Alice'
        $output[1] | Should -Match 'Alice'
    }

    it 'filters output to only include messages within the specified date range' {
        $inputFile  = Join-Path $testDir 'chat_daterange.txt'
        $outputFile = Join-Path $testDir 'output_daterange.txt'
        $input = @"
[1/1/23, 9:00:00 AM] Alice: January message
[6/15/23, 9:00:00 AM] Bob: June message
[12/31/23, 9:00:00 AM] Alice: December message
"@
        Set-Content -Path $inputFile -Value $input -Encoding UTF8
        Invoke-Script @($inputFile, $outputFile, 'T', '', '1/1/2023', '6/30/2023')
        $output = Get-Content $outputFile -Encoding UTF8
        $output.Count | Should -Be 2
        $output[0] | Should -Match 'January message'
        $output[1] | Should -Match 'June message'
    }
}
