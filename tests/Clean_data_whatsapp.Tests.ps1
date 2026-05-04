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

    # --- Original test ---
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

    # --- Gap 1: Clean-Path strips quotes from file paths ---
    it 'strips surrounding quotes from file paths via Clean-Path' {
        $inputFile  = Join-Path $testDir 'chat_quotes.txt'
        $outputFile = Join-Path $testDir 'output_quotes.txt'
        $input = '[1/5/23, 8:00:00 AM] Alice: Hi there!'
        Set-Content -Path $inputFile -Value $input -Encoding UTF8

        # Wrap both paths in extra quotes, as a user drag-drop might produce
        $quotedInput  = '"' + $inputFile  + '"'
        $quotedOutput = '"' + $outputFile + '"'
        $responses = @($quotedInput, $quotedOutput)
        $index = 0
        Mock -CommandName Read-Host -MockWith { $responses[$index++] }
        . $scriptPath

        $outputFile | Should -Exist
        $output = Get-Content $outputFile -Encoding UTF8
        $output | Should -Be "1/5/23`t8:00:00 AM`tAlice`tHi there!"
    }

    # --- Gap 2: LRM and non-standard Unicode spaces are removed ---
    it 'removes LRM and replaces non-standard Unicode spaces' {
        $inputFile  = Join-Path $testDir 'chat_unicode.txt'
        $outputFile = Join-Path $testDir 'output_unicode.txt'

        # LRM (\u200E) inside timestamp area, narrow no-break space (\u202F) between time and AM/PM
        $rawLine = "[1/6/23, 9:00:00`u{202F}AM] Bob: Good day"
        Set-Content -Path $inputFile -Value $rawLine -Encoding UTF8

        $responses = @($inputFile, $outputFile)
        $index = 0
        Mock -CommandName Read-Host -MockWith { $responses[$index++] }
        . $scriptPath

        $output = Get-Content $outputFile -Encoding UTF8
        # After Unicode cleanup the line should parse cleanly into four columns
        $output | Should -Be "1/6/23`t9:00:00 AM`tBob`tGood day"
    }

    # --- Gap 3: Colon inside message body is preserved in Column D ---
    it 'preserves colons inside the message body in Column D' {
        $inputFile  = Join-Path $testDir 'chat_colon.txt'
        $outputFile = Join-Path $testDir 'output_colon.txt'
        $input = '[1/7/23, 10:00:00 AM] John: Note: see below for details'
        Set-Content -Path $inputFile -Value $input -Encoding UTF8

        $responses = @($inputFile, $outputFile)
        $index = 0
        Mock -CommandName Read-Host -MockWith { $responses[$index++] }
        . $scriptPath

        $output = Get-Content $outputFile -Encoding UTF8
        $output | Should -Be "1/7/23`t10:00:00 AM`tJohn`tNote: see below for details"
    }

    # --- Gap 4: Non-existent input file exits gracefully ---
    it 'exits gracefully when the input file does not exist' {
        $missingFile = Join-Path $testDir 'does_not_exist.txt'
        $outputFile  = Join-Path $testDir 'output_missing.txt'

        $responses = @($missingFile, $outputFile)
        $index = 0
        Mock -CommandName Read-Host -MockWith { $responses[$index++] }

        # Script calls exit; catch the terminating error rather than letting it propagate
        { . $scriptPath } | Should -Not -Throw
        $outputFile | Should -Not -Exist
    }

    # --- Gap 5: Empty input file produces an empty output file ---
    it 'creates an empty output file when input file has no content' {
        $inputFile  = Join-Path $testDir 'chat_empty.txt'
        $outputFile = Join-Path $testDir 'output_empty.txt'
        Set-Content -Path $inputFile -Value '' -Encoding UTF8

        $responses = @($inputFile, $outputFile)
        $index = 0
        Mock -CommandName Read-Host -MockWith { $responses[$index++] }
        . $scriptPath

        $outputFile | Should -Exist
        $output = Get-Content $outputFile -Encoding UTF8
        $output | Should -BeNullOrEmpty
    }

    # --- Gap 6: Multi-line continuation is merged into Column D ---
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

        $responses = @($inputFile, $outputFile)
        $index = 0
        Mock -CommandName Read-Host -MockWith { $responses[$index++] }
        . $scriptPath

        $output = Get-Content $outputFile -Encoding UTF8
        $output[0] | Should -Be "1/8/23`t11:00:00 AM`tAlice`tFirst line of message Second line of message Third line of message"
        $output[1] | Should -Be "1/8/23`t11:01:00 AM`tBob`tReply"
    }
}
