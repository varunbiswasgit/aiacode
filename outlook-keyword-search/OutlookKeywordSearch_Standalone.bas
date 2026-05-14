Attribute VB_Name = "OutlookKeywordSearch_Standalone"
Option Explicit

' ============================================================
' Outlook Keyword Search — Standalone VBA (no PowerShell)
' Author  : Varun Biswas
' Repo    : varunbiswasgit/aiacode
' Purpose : Pure VBA macro. No external dependencies.
'           Runs entirely inside Outlook.
'           Single mode: result shown in MsgBox + Immediate Window.
'           Batch mode: reads keywords from Excel, appends
'           Match Email, Sender, and Status columns to same file.
' Install : Alt+F11 in Outlook > Insert Module > paste this file.
'           Run via Tools > Macros > RunKeywordSearch.
' Fix     : Data start row is auto-detected (not hardcoded to row 2).
'           lastCol is scoped to the keyword column data range only,
'           preventing stray content in other columns from skewing
'           the output column position.
' ============================================================

Public Sub RunKeywordSearch()
    Dim modeChoice As String

    modeChoice = Trim(InputBox( _
        "Choose mode:" & vbCrLf & _
        "S = Single keyword search" & vbCrLf & _
        "B = Batch mode from Excel", _
        "Outlook Keyword Search"))

    If Len(modeChoice) = 0 Then Exit Sub

    Select Case UCase(modeChoice)
        Case "S" : RunSingleKeywordSearch
        Case "B" : RunBatchKeywordSearch
        Case Else
            MsgBox "Invalid mode. Please enter S or B.", vbExclamation
    End Select
End Sub

' ------------------------------------------------------------
' SINGLE MODE
' ------------------------------------------------------------
Private Sub RunSingleKeywordSearch()
    Dim keyword As String
    keyword = Trim(InputBox("Enter keyword or phrase to search in email BODY:", "Single Keyword Search"))
    If Len(keyword) = 0 Then
        MsgBox "No keyword entered.", vbExclamation
        Exit Sub
    End If

    Dim result As String
    Dim bestItem As Object
    Dim bestPath As String
    SearchAllFolders keyword, bestItem, bestPath

    If bestItem Is Nothing Then
        result = "No email found for keyword: " & keyword
    Else
        result = "Keyword  : " & keyword & vbCrLf & _
                 "Received : " & Format(bestItem.ReceivedTime, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
                 "Subject  : " & bestItem.Subject & vbCrLf & _
                 "Sender   : " & SenderText(bestItem) & vbCrLf & _
                 "Folder   : " & bestPath
    End If

    MsgBox result, vbInformation, "Keyword Search Result"
    Debug.Print String(60, "-")
    Debug.Print result
    Debug.Print String(60, "-")
End Sub

' ------------------------------------------------------------
' BATCH MODE
' ------------------------------------------------------------
Private Sub RunBatchKeywordSearch()
    Dim filePath As String
    Dim colRef   As String

    filePath = Trim(InputBox("Enter full Excel file path (e.g. C:\Users\You\keywords.xlsx):", "Batch Mode — File Path"))
    If Len(filePath) = 0 Then
        MsgBox "No file path entered.", vbExclamation
        Exit Sub
    End If
    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbCritical
        Exit Sub
    End If

    colRef = UCase(Trim(InputBox("Enter column letter containing keywords (e.g. A):", "Batch Mode — Keyword Column")))
    If Len(colRef) = 0 Then
        MsgBox "No column entered.", vbExclamation
        Exit Sub
    End If
    If Not colRef Like "[A-Z]" And Not colRef Like "[A-Z][A-Z]" Then
        MsgBox "Invalid Excel column reference: " & colRef, vbCritical
        Exit Sub
    End If

    ' Open Excel via late binding
    Dim xlApp As Object, wb As Object, ws As Object
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set wb = xlApp.Workbooks.Open(filePath)
    Set ws = wb.Worksheets(1)

    Dim colNum As Long
    colNum = ColLetterToNumber(colRef)

    Dim firstDataRow As Long
    Dim scanRow As Long
    firstDataRow = 0
    For scanRow = 1 To 1000
        If Trim(CStr(IIf(IsNull(ws.Cells(scanRow, colNum).Value), "", ws.Cells(scanRow, colNum).Value))) <> "" Then
            Dim cellVal As String
            cellVal = Trim(CStr(ws.Cells(scanRow, colNum).Value))
            If Not IsNumeric(cellVal) Then
                firstDataRow = scanRow
                Exit For
            End If
        End If
    Next scanRow

    If firstDataRow = 0 Then
        wb.Close False
        xlApp.Quit
        MsgBox "No data found in column " & colRef & ".", vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colNum).End(-4162).Row

    If lastRow < firstDataRow Then
        wb.Close False
        xlApp.Quit
        MsgBox "No keyword data found below row " & firstDataRow & " in column " & colRef & ".", vbExclamation
        Exit Sub
    End If

    Dim dataRange As Object
    Dim findCell As Object
    Set dataRange = ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(lastRow, ws.Columns.Count))
    Set findCell = dataRange.Find("*", dataRange.Cells(1, 1), , , 2, 2)

    Dim lastCol As Long
    If findCell Is Nothing Then
        lastCol = colNum
    Else
        lastCol = findCell.Column
    End If

    Dim resultCol As Long : resultCol = lastCol + 1
    Dim senderCol As Long : senderCol = lastCol + 2
    Dim statusCol As Long : statusCol = lastCol + 3

    ws.Cells(firstDataRow, resultCol).Value = "Match Email"
    ws.Cells(firstDataRow, senderCol).Value = "Sender"
    ws.Cells(firstDataRow, statusCol).Value = "Status"

    Dim row As Long
    Dim kw As String
    Dim bestItem As Object
    Dim bestPath As String
    Dim processed As Long : processed = 0
    Dim found As Long    : found = 0

    For row = firstDataRow + 1 To lastRow
        kw = Trim(CStr(IIf(IsNull(ws.Cells(row, colNum).Value), "", ws.Cells(row, colNum).Value)))

        If Len(kw) = 0 Then
            ws.Cells(row, resultCol).Value = ""
            ws.Cells(row, senderCol).Value = ""
            ws.Cells(row, statusCol).Value = "Blank Keyword"
        Else
            Set bestItem = Nothing
            bestPath = ""
            SearchAllFolders kw, bestItem, bestPath

            If bestItem Is Nothing Then
                ws.Cells(row, resultCol).Value = ""
                ws.Cells(row, senderCol).Value = ""
                ws.Cells(row, statusCol).Value = "Not Found"
            Else
                ws.Cells(row, resultCol).Value = Format(bestItem.ReceivedTime, "yyyy-mm-dd hh:nn:ss") & _
                                                  " | " & bestItem.Subject & " | " & bestPath
                ws.Cells(row, senderCol).Value = SenderText(bestItem)
                ws.Cells(row, statusCol).Value = "Found"
                found = found + 1
            End If
            processed = processed + 1
        End If

        DoEvents
    Next row

    wb.Save
    wb.Close False
    xlApp.Quit
    Set ws = Nothing : Set wb = Nothing : Set xlApp = Nothing

    MsgBox "Batch complete." & vbCrLf & _
           "Data started at row : " & firstDataRow & vbCrLf & _
           "Processed           : " & processed & vbCrLf & _
           "Found               : " & found, vbInformation, "Batch Search Complete"
End Sub

' ------------------------------------------------------------
' SEARCH — all stores, folders, subfolders
' ------------------------------------------------------------
Private Sub SearchAllFolders(keyword As String, ByRef bestItem As Object, ByRef bestPath As String)
    Dim store As Object
    For Each store In Application.Session.Stores
        SearchFolderRecursive store.GetRootFolder(), keyword, bestItem, bestPath
    Next store
End Sub

Private Sub SearchFolderRecursive(folder As Object, keyword As String, ByRef bestItem As Object, ByRef bestPath As String)
    If SkipFolder(folder) Then Exit Sub

    On Error Resume Next
    Dim items As Object
    Set items = folder.Items
    items.Sort "[ReceivedTime]", False

    Dim itm As Object
    For Each itm In items
        If itm.Class = 43 Then
            If Not bestItem Is Nothing Then
                If itm.ReceivedTime > bestItem.ReceivedTime Then Exit For
            End If
            Dim body As String
            body = ""
            body = itm.body
            If InStr(1, body, keyword, vbTextCompare) > 0 Then
                If bestItem Is Nothing Then
                    Set bestItem = itm
                    bestPath = folder.FolderPath
                ElseIf itm.ReceivedTime < bestItem.ReceivedTime Then
                    Set bestItem = itm
                    bestPath = folder.FolderPath
                End If
                Exit For
            End If
        End If
    Next itm
    On Error GoTo 0

    Dim sub As Object
    For Each sub In folder.Folders
        SearchFolderRecursive sub, keyword, bestItem, bestPath
    Next sub
End Sub

' ------------------------------------------------------------
' HELPERS
' ------------------------------------------------------------
Private Function SkipFolder(folder As Object) As Boolean
    Dim skipNames As Variant
    skipNames = Array("calendar", "contacts", "tasks", "notes", _
                      "junk email", "deleted items", "rss feeds", _
                      "outbox", "drafts", "sync issues", "conflicts", _
                      "local failures", "server failures", "recoverable items")
    Dim n As String : n = LCase(folder.Name)
    Dim i As Integer
    For i = 0 To UBound(skipNames)
        If n = skipNames(i) Then SkipFolder = True : Exit Function
    Next i
    On Error Resume Next
    If folder.DefaultItemType <> 0 Then SkipFolder = True
    On Error GoTo 0
End Function

Private Function SenderText(itm As Object) As String
    Dim nm As String : nm = ""
    Dim em As String : em = ""
    On Error Resume Next
    nm = itm.SenderName
    em = itm.SenderEmailAddress
    On Error GoTo 0
    If Len(nm) > 0 And Len(em) > 0 Then
        SenderText = nm & " <" & em & ">"
    ElseIf Len(em) > 0 Then
        SenderText = em
    Else
        SenderText = nm
    End If
End Function

Private Function ColLetterToNumber(col As String) As Long
    Dim i As Integer, result As Long
    result = 0
    For i = 1 To Len(col)
        result = result * 26 + (Asc(UCase(Mid(col, i, 1))) - Asc("A") + 1)
    Next i
    ColLetterToNumber = result
End Function
