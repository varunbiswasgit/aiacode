<#
.SYNOPSIS
    Outlook Keyword Search — PowerShell Engine

.DESCRIPTION
    Searches Outlook email bodies for one keyword (single mode) or a list of
    keywords from an Excel file (batch mode). Runs out-of-process so Outlook
    remains fully responsive during the search.

    Fixes applied:
      - Data start row is auto-detected by scanning the keyword column for the
        first non-empty cell, so headers on any row work correctly.
      - lastCol detection is scoped to the keyword data range only, preventing
        stray content in columns C, D, etc. from skewing output column placement.

.PARAMETER Mode
    S = Single keyword search
    B = Batch search from Excel file

.PARAMETER Keyword
    Keyword or phrase to search (Single mode only)

.PARAMETER FilePath
    Full path to Excel file containing keywords (Batch mode only)

.PARAMETER Column
    Column letter in the Excel file that contains keywords (Batch mode only)

.AUTHOR
    Varun Biswas
.REPO
    varunbiswasgit/aiacode
#>

param(
    [Parameter(Mandatory)][ValidateSet('S','B')][string]$Mode,
    [string]$Keyword   = '',
    [string]$FilePath  = '',
    [string]$Column    = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$SKIP_FOLDER_NAMES = @(
    'calendar','contacts','tasks','notes',
    'junk email','deleted items','rss feeds',
    'outbox','drafts','sync issues','conflicts',
    'local failures','server failures','recoverable items'
)

function Show-Toast {
    param([string]$Title, [string]$Message)
    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null
        $template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent(
            [Windows.UI.Notifications.ToastTemplateType]::ToastText02)
        $template.SelectSingleNode('//text[@id=1]').InnerText = $Title
        $template.SelectSingleNode('//text[@id=2]').InnerText = $Message
        $toast = [Windows.UI.Notifications.ToastNotification]::new($template)
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Outlook Keyword Search').Show($toast)
    } catch {}
}

$LogPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'OutlookKeywordSearch.log'

function Write-Log {
    param([string]$Text)
    Add-Content -Path $LogPath -Value ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Text)
}

function ConvertTo-ColumnNumber {
    param([string]$Letter)
    $Letter = $Letter.ToUpper().Trim()
    if ($Letter -notmatch '^[A-Z]+$') { return 0 }
    $result = 0
    foreach ($ch in $Letter.ToCharArray()) {
        $result = $result * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $result
}

function Get-SenderText {
    param($MailItem)
    $name  = try { $MailItem.SenderName }         catch { '' }
    $email = try { $MailItem.SenderEmailAddress } catch { '' }
    if ($name -and $email)  { return "$name <$email>" }
    if ($email)             { return $email }
    if ($name)              { return $name }
    return ''
}

function Test-SkipFolder {
    param($Folder)
    if ($SKIP_FOLDER_NAMES -contains $Folder.Name.ToLower()) { return $true }
    try { if ($Folder.DefaultItemType -ne 0) { return $true } } catch {}
    return $false
}

function Search-FolderRecursive {
    param($Folder, [string]$Keyword, [ref]$BestItem, [ref]$BestPath)

    if (Test-SkipFolder $Folder) { return }

    try {
        $items = $Folder.Items
        $items.Sort('[ReceivedTime]', $false)

        foreach ($itm in $items) {
            if ($itm.Class -ne 43) { continue }
            if ($null -ne $BestItem.Value) {
                if ($itm.ReceivedTime -gt $BestItem.Value.ReceivedTime) { break }
            }
            $body = try { $itm.Body } catch { '' }
            if ($body -and $body.IndexOf($Keyword, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                if ($null -eq $BestItem.Value -or $itm.ReceivedTime -lt $BestItem.Value.ReceivedTime) {
                    $BestItem.Value = $itm
                    $BestPath.Value = $Folder.FolderPath
                }
                break
            }
        }
    } catch {
        Write-Log "Skipped folder '$($Folder.Name)': $_"
    }

    foreach ($sub in $Folder.Folders) {
        Search-FolderRecursive -Folder $sub -Keyword $Keyword -BestItem $BestItem -BestPath $BestPath
    }
}

function Search-AllFolders {
    param([string]$Keyword)
    $bestItem = [ref]$null
    $bestPath = [ref]''
    foreach ($store in $outlook.Session.Stores) {
        $root = $store.GetRootFolder()
        Search-FolderRecursive -Folder $root -Keyword $Keyword -BestItem $bestItem -BestPath $bestPath
    }
    return @{ Item = $bestItem.Value; Path = $bestPath.Value }
}

Write-Log "=== Search started | Mode=$Mode ==="
try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-Log "FATAL: Cannot connect to Outlook COM: $_"
    Show-Toast 'Outlook Keyword Search' 'ERROR: Could not connect to Outlook. Is Outlook running?'
    exit 1
}

if ($Mode -eq 'S') {
    if (-not $Keyword) {
        Write-Log 'ERROR: No keyword supplied for single mode.'
        Show-Toast 'Outlook Keyword Search' 'ERROR: No keyword provided.'
        exit 1
    }
    Write-Log "Single mode | Keyword: $Keyword"
    $result = Search-AllFolders -Keyword $Keyword
    if ($null -eq $result.Item) {
        $msg = "No email found for keyword: $Keyword"
        Write-Log $msg
        Show-Toast 'Outlook Keyword Search — Not Found' $msg
    } else {
        $itm    = $result.Item
        $rcvd   = $itm.ReceivedTime.ToString('yyyy-MM-dd HH:mm:ss')
        $msg    = "$rcvd | $($itm.Subject) | $(Get-SenderText $itm) | $($result.Path)"
        Write-Log "FOUND: $msg"
        Show-Toast "Keyword Found: $Keyword" $msg
    }
}

if ($Mode -eq 'B') {
    if (-not (Test-Path $FilePath)) {
        Write-Log "ERROR: File not found: $FilePath"
        Show-Toast 'Outlook Keyword Search' "ERROR: File not found: $FilePath"
        exit 1
    }
    $colNum = ConvertTo-ColumnNumber $Column
    if ($colNum -le 0) {
        Write-Log "ERROR: Invalid column reference: $Column"
        Show-Toast 'Outlook Keyword Search' "ERROR: Invalid column: $Column"
        exit 1
    }

    $xlApp = New-Object -ComObject Excel.Application
    $xlApp.Visible = $false
    $xlApp.DisplayAlerts = $false
    $wb = $xlApp.Workbooks.Open($FilePath)
    $ws = $wb.Worksheets.Item(1)

    # ----------------------------------------------------------
    # FIX 1: Auto-detect first data row in the keyword column.
    #         Scan from row 1 downward; treat the first non-empty
    #         cell as the header row for output labels, and start
    #         processing keywords from the row after it.
    # ----------------------------------------------------------
    $firstDataRow = 0
    for ($scanRow = 1; $scanRow -le 1000; $scanRow++) {
        $cellVal = $ws.Cells.Item($scanRow, $colNum).Value2
        if ($null -ne $cellVal -and $cellVal.ToString().Trim().Length -gt 0) {
            $firstDataRow = $scanRow
            break
        }
    }

    if ($firstDataRow -eq 0) {
        $wb.Close($false)
        $xlApp.Quit()
        Write-Log "ERROR: No data found in column $Column"
        Show-Toast 'Outlook Keyword Search' "ERROR: No data found in column $Column"
        exit 1
    }

    $lastRow = $ws.Cells($ws.Rows.Count, $colNum).End(-4162).Row

    if ($lastRow -lt ($firstDataRow + 1)) {
        $wb.Close($false)
        $xlApp.Quit()
        Write-Log "ERROR: No keyword rows found below row $firstDataRow"
        Show-Toast 'Outlook Keyword Search' "ERROR: No keywords found below row $firstDataRow"
        exit 1
    }

    # ----------------------------------------------------------
    # FIX 2: Scope lastCol search to the keyword data rows only.
    #         This prevents headers or data in other unrelated
    #         columns (e.g. C, D) from inflating lastCol and
    #         causing output columns to land in the wrong place.
    # ----------------------------------------------------------
    $dataRange = $ws.Range(
        $ws.Cells.Item($firstDataRow, 1),
        $ws.Cells.Item($lastRow, $ws.Columns.Count)
    )
    $findCell = $dataRange.Find(
        '*',
        $dataRange.Cells.Item(1, 1),
        [Type]::Missing, [Type]::Missing,
        2,   # xlByColumns
        2    # xlPrevious
    )
    $lastCol = if ($findCell) { $findCell.Column } else { $colNum }

    $resultCol = $lastCol + 1
    $senderCol = $lastCol + 2
    $statusCol = $lastCol + 3

    # Write output headers aligned to the first data row
    $ws.Cells.Item($firstDataRow, $resultCol).Value2 = 'Match Email'
    $ws.Cells.Item($firstDataRow, $senderCol).Value2 = 'Sender'
    $ws.Cells.Item($firstDataRow, $statusCol).Value2 = 'Status'

    $processed = 0
    $found = 0

    for ($row = $firstDataRow + 1; $row -le $lastRow; $row++) {
        $kw = $ws.Cells.Item($row, $colNum).Value2
        $kw = if ($kw) { $kw.ToString().Trim() } else { '' }
        if ($kw.Length -eq 0) {
            $ws.Cells.Item($row, $statusCol).Value2 = 'Blank Keyword'
            continue
        }
        Write-Log "Searching keyword (row $row / $lastRow): $kw"
        $result = Search-AllFolders -Keyword $kw
        if ($null -eq $result.Item) {
            $ws.Cells.Item($row, $resultCol).Value2 = ''
            $ws.Cells.Item($row, $senderCol).Value2 = ''
            $ws.Cells.Item($row, $statusCol).Value2 = 'Not Found'
        } else {
            $itm  = $result.Item
            $rcvd = $itm.ReceivedTime.ToString('yyyy-MM-dd HH:mm:ss')
            $ws.Cells.Item($row, $resultCol).Value2 = "$rcvd | $($itm.Subject) | $($result.Path)"
            $ws.Cells.Item($row, $senderCol).Value2 = Get-SenderText $itm
            $ws.Cells.Item($row, $statusCol).Value2 = 'Found'
            $found++
            Write-Log "FOUND row $row: $rcvd | $($itm.Subject)"
        }
        $processed++
        [System.Threading.Thread]::Sleep(0)
    }

    $wb.Save()
    $wb.Close($false)
    $xlApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlApp) | Out-Null

    $summary = "Batch complete. First data row: $firstDataRow | Processed: $processed | Found: $found"
    Write-Log $summary
    Show-Toast 'Outlook Keyword Search — Batch Complete' "Processed: $processed | Found: $found"
}

Write-Log '=== Search ended ==='
