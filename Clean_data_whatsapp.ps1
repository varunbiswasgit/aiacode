# Function to clean up any extra quotes
function Clean-Path($path) {
    return $path -replace '"', ''  # Remove any quotes from the path
}

# Prompt the user for the input file location
$inputFile = Read-Host "Please enter the full path of the input text file"
$inputFile = Clean-Path $inputFile  # Remove any extra quotes

# Check if the input file exists
if (-not (Test-Path $inputFile)) {
    Write-Host "Error: The input file '$inputFile' does not exist. Exiting..."
    exit
}

# Prompt the user for the output file location
$outputFile = Read-Host "Please enter the full path of the output text file"
$outputFile = Clean-Path $outputFile  # Remove any extra quotes

# Regular expression to match date and time enclosed in square brackets, e.g., [M/D/YY, H:MM:SS AM/PM]
$pattern = "^\[\d{1,2}/\d{1,2}/\d{2,4}, \d{1,2}:\d{2}:\d{2} [AP]M\]"

# Initialize a list to store corrected records
$fixedLines = @()

# Read the input file line by line
Get-Content $inputFile -Encoding UTF8 | ForEach-Object {
    # Remove LRM characters and replace problematic Unicode characters with spaces
    $_ = $_ -replace '\u200E', ''  # Remove Left-to-Right Mark (LRM) entirely
    $_ = $_ -replace '\u202F|\u00A0|\u2007', ' '  # Replace other problematic Unicode spaces with regular spaces

    # Check if the line starts with a date and time enclosed in square brackets
    if ($_ -match $pattern) {
        # Extract Column A (Date) and Column B (Time) from inside the square brackets
        $dateTime = $_ -replace "\[(.*?)\].*", '$1'
        $date, $time = $dateTime -split ', '  # Split date and time by comma
        
        # Extract Column C (Name) by capturing text between ] and :
        $columnC = ($_ -replace ".*\] (.*?):.*", '$1').Trim()  # Trim leading/trailing spaces
        
        # Extract Column D (Message) by capturing text after the colon
        $columnD = ($_ -replace ".*: (.*)", '$1').Trim()  # Trim leading/trailing spaces

        # Concatenate and store the result for output (CSV-like format with tab separation)
        $fixedLines += "$date`t$time`t$columnC`t$columnD"
    } elseif ($fixedLines.Count -gt 0) {
        # Ensure there's already a record before appending a continuation line
        $fixedLines[-1] += " " + $_.Trim()  # Append to the last message (Column D), also Trim
    }
}

# Write the corrected records to the output file with newline between records
$fixedLines -join "`n" | Set-Content $outputFile

# Confirmation message
Write-Host "Data processing complete. Output has been saved to $outputFile"
