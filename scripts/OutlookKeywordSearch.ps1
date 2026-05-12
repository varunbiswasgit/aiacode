<#
.SYNOPSIS
    Outlook Keyword Search — PowerShell Engine (v2)

.DESCRIPTION
    Searches Outlook email bodies for one keyword (single mode) or a list of
    keywords from an Excel file (batch mode). Runs out-of-process so Outlook
    remains fully responsive during the search.

    Optimizations implemented:
      3. Skip non-mail folders (Calendar, Contacts, Tasks, Junk, Deleted, etc.)
      4. Early exit per folder once current item is newer than best match found
      5. Yields to OS between keywords in batch mode (non-blocking iteration)
      6. Runs as a separate PS process — Outlook UI never blocked

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

# ---------------------------------------------------------------
# CONSTANTS — folders to skip entirely (optimization 3)
# ---------------------------------------------------------------
$SKIP_FOLDER_NAMES = @(
    'calendar','contacts','tasks','notes',
    'junk email','deleted items','rss feeds',
    'outbox','drafts','sync issues','conflicts',
    'local failures','server failures','recoverable items'
)

# ---------------------------------------------------------------
# TOAST NOTIFICATION HELPER
# ---------------------------------------------------------------
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
    } catch {
        # Fallback: write to log only — toast may not be available on all OS versions
    }
}

# ---------------------------------------------------------------
# LOG HELPER
# ---------------------------------------------------------------
$LogPath = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'OutlookKeywordSearch.log'

function Write-Log {
    param([string]$Text)
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Text
    Add-Content -Path $LogPath -Value $line
}

# ---------------------------------------------------------------
# COLUMN LETTER TO NUMBER
# ---------------------------------------------------------------
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

# ---------------------------------------------------------------
# SENDER TEXT
# ---------------------------------------------------------------
function Get-SenderText {
    param($MailItem)
    $name  = try { $MailItem.SenderName }         catch { '' }
    $email = try { $MailItem.SenderEmailAddress } catch { '' }
    if ($name -and $email)  { return "$name <$email>" }
    if ($email)             { return $email }
    if ($name)              { return $name }
    return ''
}

# ---------------------------------------------------------------
# SHOULD SKIP FOLDER (optimization 3)
# ---------------------------------------------------------------
function Test-SkipFolder {
    param($Folder)
    # Skip by name
    if ($SKIP_FOLDER_NAMES -contains $Folder.Name.ToLower()) { return $true }
    # Skip if folder does not hold mail items (0 = olMailItem)
    try { if ($Folder.DefaultItemType -ne 0) { return $true } } catch {}
    return $false
}

# ---------------------------------------------------------------
# SEARCH ONE FOLDER RECURSIVELY
# Returns the oldest MailItem match, or $null
# ---------------------------------------------------------------
function Search-FolderRecursive {
    param(
        $Folder,
        [string]$Keyword,
        [ref]$BestItem,
        [ref]$BestPath
    )

    if (Test-SkipFolder $Folder) { return }

    try {
        $items = $Folder.Items
        $items.Sort('[ReceivedTime]', $false)   # ascending = oldest first

        foreach ($itm in $items) {
            if ($itm.Class -ne 43) { continue }   # 43 = olMail

            # Optimization 4: early exit — if item is already newer than best
            # match, all remaining items in this ascending-sorted folder are
            # also newer. Skip the rest of this folder.
            if ($null -ne $BestItem.Value) {
                if ($itm.ReceivedTime -gt $BestItem.Value.ReceivedTime) { break }
            }

            $body = try { $itm.Body } catch { '' }
            if ($body -and $body.IndexOf($Keyword, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                if ($null -eq $BestItem.Value -or $itm.ReceivedTime -lt $BestItem.Value.ReceivedTime) {
                    $BestItem.Value = $itm
                    $BestPath.Value = $Folder.FolderPath
                }
                break   # oldest in this folder found — no need to continue
            }
        }
    } catch {
        Write-Log "Skipped folder '$($Folder.Name)': $_"
    }

    # Recurse into subfolders
    foreach ($sub in $Folder.Folders) {
        Search-FolderRecursive -Folder $sub -Keyword $Keyword -BestItem $BestItem -BestPath $BestPath
    }
}

# ---------------------------------------------------------------
# SEARCH ALL STORES
# ---------------------------------------------------------------
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

# ---------------------------------------------------------------
# CONNECT TO OUTLOOK
# ---------------------------------------------------------------
Write-Log "=== Search started | Mode=$Mode ==="
try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-Log "FATAL: Cannot connect to Outlook COM: $_"
    Show-Toast 'Outlook Keyword Search' 'ERROR: Could not connect to Outlook. Is Outlook running?'
    exit 1
}

# ---------------------------------------------------------------
# SINGLE MODE
# ---------------------------------------------------------------
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
        $sender = Get-SenderText $itm
        $rcvd   = $itm.ReceivedTime.ToString('yyyy-MM-dd HH:mm:ss')
        $subj   = $itm.Subject
        $path   = $result.Path
        $msg    = "$rcvd | $subj | $sender | $path"
        Write-Log "FOUND: $msg"
        Show-Toast "Keyword Found: $Keyword" $msg
    }
}

# ---------------------------------------------------------------
# BATCH MODE
# ---------------------------------------------------------------
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

    # Open Excel via COM
    $xlApp = New-Object -ComObject Excel.Application
    $xlApp.Visible = $false
    $xlApp.DisplayAlerts = $false
    $wb = $xlApp.Workbooks.Open($FilePath)
    $ws = $wb.Worksheets.Item(1)

    # Find last row in keyword column
    $lastRow = $ws.Cells($ws.Rows.Count, $colNum).End(-4162).Row   # -4162 = xlUp

    # Find last used column
    $findCell = $ws.Cells.Find('*', $ws.Cells.Item(1,1), [Type]::Missing, [Type]::Missing, 2, 2)
    $lastCol  = if ($findCell) { $findCell.Column } else { $colNum }

    $resultCol = $lastCol + 1
    $senderCol = $lastCol + 2
    $statusCol = $lastCol + 3

    $ws.Cells.Item(1, $resultCol).Value2 = 'Match Email'
    $ws.Cells.Item(1, $senderCol).Value2 = 'Sender'
    $ws.Cells.Item(1, $statusCol).Value2 = 'Status'

    $processed = 0
    $found     = 0

    for ($row = 2; $row -le $lastRow; $row++) {
        $kw = $ws.Cells.Item($row, $colNum).Value2
        $kw = if ($kw) { $kw.ToString().Trim() } else { '' }

        if ($kw.Length -eq 0) {
            $ws.Cells.Item($row, $resultCol).Value2 = ''
            $ws.Cells.Item($row, $senderCol).Value2 = ''
            $ws.Cells.Item($row, $statusCol).Value2 = 'Blank Keyword'
            continue
        }

        Write-Log "Searching keyword ($row/$lastRow): $kw"
        $result = Search-AllFolders -Keyword $kw

        if ($null -eq $result.Item) {
            $ws.Cells.Item($row, $resultCol).Value2 = ''
            $ws.Cells.Item($row, $senderCol).Value2 = ''
            $ws.Cells.Item($row, $statusCol).Value2 = 'Not Found'
        } else {
            $itm    = $result.Item
            $rcvd   = $itm.ReceivedTime.ToString('yyyy-MM-dd HH:mm:ss')
            $subj   = $itm.Subject
            $path   = $result.Path
            $sender = Get-SenderText $itm
            $ws.Cells.Item($row, $resultCol).Value2 = "$rcvd | $subj | $path"
            $ws.Cells.Item($row, $senderCol).Value2 = $sender
            $ws.Cells.Item($row, $statusCol).Value2 = 'Found'
            $found++
            Write-Log "FOUND row $row: $rcvd | $subj"
        }

        $processed++

        # Optimization 5: yield to OS between keywords
        [System.Threading.Thread]::Sleep(0)
    }

    $wb.Save()
    $wb.Close($false)
    $xlApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xlApp) | Out-Null

    $summary = "Batch complete. Processed: $processed | Found: $found | File: $FilePath"
    Write-Log $summary
    Show-Toast 'Outlook Keyword Search — Batch Complete' "Processed: $processed | Found: $found"
}

Write-Log '=== Search ended ==='
