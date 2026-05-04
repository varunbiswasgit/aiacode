# Function to clean up any extra quotes
function Clean-Path($path) {
    return $path -replace '"', ''  # Remove any quotes from the path
}

# Normalise any supported timestamp format to M/D/YY so the rest of the
# pipeline works with a single canonical form.
# Supported inputs:
#   M/D/YY or M/D/YYYY   (WhatsApp default)
#   DD/MM/YYYY           (European)
#   YYYY-MM-DD           (ISO 8601)
function Normalize-Date($raw) {
    if ($raw -match '^(\d{4})-(\d{1,2})-(\d{1,2})$') {
        # ISO 8601: YYYY-MM-DD -> M/D/YY
        $y = $Matches[1]; $m = $Matches[2]; $d = $Matches[3]
        return "$([int]$m)/$([int]$d)/$($y.Substring(2))"
    }
    if ($raw -match '^(\d{1,2})/(\d{1,2})/(\d{4})$') {
        # DD/MM/YYYY (day first, 4-digit year) -> M/D/YY
        $d = $Matches[1]; $m = $Matches[2]; $y = $Matches[3]
        return "$([int]$m)/$([int]$d)/$($y.Substring(2))"
    }
    # M/D/YY or M/D/YYYY already — return as-is (strip leading zeros for consistency)
    if ($raw -match '^(\d{1,2})/(\d{1,2})/(\d{2,4})$') {
        return "$([int]$Matches[1])/$([int]$Matches[2])/$($Matches[3])"
    }
    return $raw  # Unrecognised format — pass through unchanged
}

# ── Prompts ──────────────────────────────────────────────────────────────────

$inputFile = Read-Host "Please enter the full path of the input text file"
$inputFile = Clean-Path $inputFile

if (-not (Test-Path $inputFile)) {
    Write-Host "Error: The input file '$inputFile' does not exist. Exiting..."
    exit
}

$outputFile = Read-Host "Please enter the full path of the output file"
$outputFile = Clean-Path $outputFile

# Output format
$formatChoice = Read-Host "Output format — enter T for Tab-separated (default) or C for CSV"
if ($formatChoice -match '^[Cc]') {
    $outputFormat = 'CSV'
    $sep = ','
} else {
    $outputFormat = 'TSV'
    $sep = "`t"
}

# Optional sender filter
$senderFilter = Read-Host "Filter by sender name (leave blank for all senders)"
$senderFilter = $senderFilter.Trim()

# Optional date range filter
$dateFromStr = Read-Host "Filter FROM date (MM/DD/YYYY, leave blank for no start limit)"
$dateToStr   = Read-Host "Filter TO date   (MM/DD/YYYY, leave blank for no end limit)"

$dateFrom = $null
$dateTo   = $null
if ($dateFromStr.Trim() -ne '') {
    try { $dateFrom = [datetime]::ParseExact($dateFromStr.Trim(), 'M/d/yyyy', $null) }
    catch { Write-Host "Warning: Could not parse FROM date '$dateFromStr'. Date filter ignored." }
}
if ($dateToStr.Trim() -ne '') {
    try { $dateTo = [datetime]::ParseExact($dateToStr.Trim(), 'M/d/yyyy', $null) }
    catch { Write-Host "Warning: Could not parse TO date '$dateToStr'. Date filter ignored." }
}

# ── Pattern (supports M/D/YY, DD/MM/YYYY, YYYY-MM-DD) ───────────────────────
# Match any of the three date formats followed by comma+space+time+optional AM/PM
$pattern = "^\[(?:\d{4}-\d{1,2}-\d{1,2}|\d{1,2}/\d{1,2}/\d{2,4}), \d{1,2}:\d{2}:\d{2}(?: [AP]M)?\]"

# ── Parse ────────────────────────────────────────────────────────────────────
$records = @()  # Each element is a hashtable: Date, Time, Sender, Message

Get-Content $inputFile -Encoding UTF8 | ForEach-Object {
    $_ = $_ -replace '\u200E', ''
    $_ = $_ -replace '\u202F|\u00A0|\u2007', ' '

    if ($_ -match $pattern) {
        $dateTime = $_ -replace "\[(.*?)\].*", '$1'
        $rawDate, $time = $dateTime -split ', '
        $date     = Normalize-Date $rawDate
        $sender   = ($_ -replace ".*\] (.*?):.*", '$1').Trim()
        $message  = ($_ -replace "^\[[^\]]+\] [^:]+: ?", '').Trim()

        $records += @{ Date = $date; Time = $time; Sender = $sender; Message = $message }
    } elseif ($records.Count -gt 0) {
        $records[-1].Message += ' ' + $_.Trim()
    }
}

# ── Filter ───────────────────────────────────────────────────────────────────
$filtered = $records | Where-Object {
    $keep = $true

    # Sender filter
    if ($senderFilter -ne '' -and $_.Sender -notlike "*$senderFilter*") { $keep = $false }

    # Date range filter
    if ($keep -and ($dateFrom -or $dateTo)) {
        try {
            $msgDate = [datetime]::ParseExact($_.Date, 'M/d/yy',   $null)
        } catch {
            try { $msgDate = [datetime]::ParseExact($_.Date, 'M/d/yyyy', $null) }
            catch { $msgDate = $null }
        }
        if ($msgDate) {
            if ($dateFrom -and $msgDate -lt $dateFrom) { $keep = $false }
            if ($dateTo   -and $msgDate -gt $dateTo)   { $keep = $false }
        }
    }
    $keep
}

# ── Format & Write ───────────────────────────────────────────────────────────
function Format-Field($value, $format) {
    if ($format -eq 'CSV') {
        # Wrap in quotes if the value contains a comma, quote, or newline
        if ($value -match '[,"\n]') {
            return '"' + ($value -replace '"', '""') + '"'
        }
    }
    return $value
}

$outputLines = $filtered | ForEach-Object {
    $d = Format-Field $_.Date    $outputFormat
    $t = Format-Field $_.Time    $outputFormat
    $s = Format-Field $_.Sender  $outputFormat
    $m = Format-Field $_.Message $outputFormat
    "$d$sep$t$sep$s$sep$m"
}

$outputLines -join "`n" | Set-Content $outputFile -Encoding UTF8

# ── Summary Report ───────────────────────────────────────────────────────────
$totalMessages  = $filtered.Count
$uniqueSenders  = ($filtered | ForEach-Object { $_.Sender } | Sort-Object -Unique).Count
$allSenders     = ($filtered | ForEach-Object { $_.Sender } | Sort-Object -Unique) -join ', '

# Determine earliest and latest dates from filtered records
$parsedDates = $filtered | ForEach-Object {
    try { [datetime]::ParseExact($_.Date, 'M/d/yy',   $null) } catch {
    try { [datetime]::ParseExact($_.Date, 'M/d/yyyy', $null) } catch { $null } }
} | Where-Object { $_ -ne $null } | Sort-Object

if ($parsedDates.Count -gt 0) {
    $earliest = $parsedDates[0].ToString('MM/dd/yyyy')
    $latest   = $parsedDates[-1].ToString('MM/dd/yyyy')
} else {
    $earliest = 'N/A'
    $latest   = 'N/A'
}

Write-Host ""
Write-Host "========================================"
Write-Host "           PROCESSING SUMMARY"
Write-Host "========================================"
Write-Host "Output format   : $outputFormat"
Write-Host "Output file     : $outputFile"
Write-Host "Total messages  : $totalMessages"
Write-Host "Unique senders  : $uniqueSenders"
Write-Host "Senders         : $allSenders"
Write-Host "Date range      : $earliest  ->  $latest"
Write-Host "========================================"
Write-Host ""
