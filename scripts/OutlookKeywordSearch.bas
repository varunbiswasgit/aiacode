Attribute VB_Name = "OutlookKeywordSearch"
Option Explicit

' ============================================================
' Outlook Keyword Search Macro
' Author  : Varun Biswas
' Repo    : varunbiswasgit/aiacode
' Purpose : Search Outlook email bodies for keywords.
'           Supports Single mode (MsgBox + Debug.Print) and
'           Batch mode (reads Excel, appends result columns).
' Scope   : All Outlook stores, folders, and subfolders.
' Match   : Returns OLDEST matching email by ReceivedTime.
' Field   : Email BODY only (case-insensitive).
' ============================================================

Private gFoundMail       As Outlook.MailItem
Private gFoundFolderPath As String

' ------------------------------------------------------------
' ENTRY POINT
' ------------------------------------------------------------
Public Sub RunKeywordSearch()
    Dim modeChoice As String
    Dim keyword    As String
    Dim filePath   As String
    Dim colRef     As String

    modeChoice = Trim(InputBox( _
        "Choose mode:" & vbCrLf & _
        "S = Single keyword search" & vbCrLf & _
        "B = Batch mode from Excel", _
        "Outlook Keyword Search"))

    If Len(modeChoice) = 0 Then Exit Sub

    Select Case UCase(modeChoice)

        Case "S"
            keyword = Trim(InputBox("Enter keyword or phrase to search in email BODY:", "Single Keyword Search"))
            If Len(keyword) = 0 Then
                MsgBox "No keyword entered.", vbExclamation
                Exit Sub
            End If
            RunSingleKeywordSearch keyword

        Case "B"
            filePath = Trim(InputBox("Enter full Excel file path (example: C:\Users\You\keywords.xlsx):", "Batch Mode - File Path"))
            If Len(filePath) = 0 Then
                MsgBox "No file path entered.", vbExclamation
                Exit Sub
            End If

            colRef = Trim(InputBox("Enter column letter containing keywords (example: A):", "Batch Mode - Keyword Column"))
            If Len(colRef) = 0 Then
                MsgBox "No column entered.", vbExclamation
                Exit Sub
            End If

            RunBatchKeywordSearch filePath, colRef

        Case Else
            MsgBox "Invalid mode. Please enter S or B.", vbExclamation

    End Select
End Sub

' ------------------------------------------------------------
' SINGLE KEYWORD MODE
' ------------------------------------------------------------
Private Sub RunSingleKeywordSearch(ByVal keyword As String)
    Dim resultText As String

    Set gFoundMail = Nothing
    gFoundFolderPath = vbNullString

    SearchAllFoldersForOldestBodyMatch keyword

    If gFoundMail Is Nothing Then
        resultText = "No email found for keyword: " & keyword
    Else
        resultText = "Oldest email found" & vbCrLf & _
                     "Keyword  : " & keyword & vbCrLf & _
                     "Received : " & Format(gFoundMail.ReceivedTime, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
                     "Subject  : " & NzText(gFoundMail.Subject) & vbCrLf & _
                     "Sender   : " & GetSenderText(gFoundMail) & vbCrLf & _
                     "Folder   : " & gFoundFolderPath
    End If

    Debug.Print String(80, "-")
    Debug.Print resultText
    Debug.Print String(80, "-")

    MsgBox resultText, vbInformation, "Single Keyword Search Result"
End Sub

' ------------------------------------------------------------
' BATCH MODE
' ------------------------------------------------------------
Private Sub RunBatchKeywordSearch(ByVal filePath As String, ByVal keywordColumn As String)
    Dim xlApp         As Object
    Dim wb            As Object
    Dim ws            As Object
    Dim lastRow       As Long
    Dim lastCol       As Long
    Dim keywordColNum As Long
    Dim resultCol     As Long
    Dim senderCol     As Long
    Dim statusCol     As Long
    Dim rowNum        As Long
    Dim keyword       As String
    Dim processedCount As Long
    Dim foundCount    As Long
    Dim findCell      As Object

    On Error GoTo ErrHandler

    If Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbCritical
        Exit Sub
    End If

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set wb = xlApp.Workbooks.Open(filePath)
    Set ws = wb.Worksheets(1)

    keywordColNum = ColumnLetterToNumber(keywordColumn)
    If keywordColNum <= 0 Then
        MsgBox "Invalid Excel column reference: " & keywordColumn, vbExclamation
        GoTo SafeExit
    End If

    ' Last used row in keyword column
    lastRow = ws.Cells(ws.Rows.Count, keywordColNum).End(-4162).Row   ' -4162 = xlUp

    ' Last used column across entire sheet using Find (more robust than End)
    Set findCell = ws.Cells.Find("*", ws.Cells(1, 1), , , 2, 2)      ' xlByColumns, xlPrevious
    If findCell Is Nothing Then
        lastCol = keywordColNum
    Else
        lastCol = findCell.Column
    End If

    ' Append new output columns automatically after last used column
    resultCol = lastCol + 1
    senderCol = lastCol + 2
    statusCol = lastCol + 3

    ws.Cells(1, resultCol).Value = "Match Email"
    ws.Cells(1, senderCol).Value = "Sender"
    ws.Cells(1, statusCol).Value = "Status"

    processedCount = 0
    foundCount = 0

    For rowNum = 2 To lastRow
        keyword = Trim(CStr(ws.Cells(rowNum, keywordColNum).Value))

        If Len(keyword) = 0 Then
            ws.Cells(rowNum, resultCol).Value = ""
            ws.Cells(rowNum, senderCol).Value = ""
            ws.Cells(rowNum, statusCol).Value = "Blank Keyword"
        Else
            Set gFoundMail = Nothing
            gFoundFolderPath = vbNullString

            SearchAllFoldersForOldestBodyMatch keyword

            If gFoundMail Is Nothing Then
                ws.Cells(rowNum, resultCol).Value = ""
                ws.Cells(rowNum, senderCol).Value = ""
                ws.Cells(rowNum, statusCol).Value = "Not Found"
            Else
                ws.Cells(rowNum, resultCol).Value = _
                    Format(gFoundMail.ReceivedTime, "yyyy-mm-dd hh:nn:ss") & " | " & _
                    NzText(gFoundMail.Subject) & " | " & _
                    gFoundFolderPath

                ws.Cells(rowNum, senderCol).Value = GetSenderText(gFoundMail)
                ws.Cells(rowNum, statusCol).Value = "Found"
                foundCount = foundCount + 1
            End If

            processedCount = processedCount + 1
        End If
    Next rowNum

    wb.Save

    MsgBox "Batch search complete." & vbCrLf & _
           "Keywords processed : " & processedCount & vbCrLf & _
           "Matches found      : " & foundCount & vbCrLf & _
           "File saved         : " & filePath, _
           vbInformation, "Batch Search Complete"

SafeExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=True
    If Not xlApp Is Nothing Then xlApp.Quit
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Batch Search Error"
    Resume SafeExit
End Sub

' ------------------------------------------------------------
' SEARCH: ALL STORES AND FOLDERS (recursive)
' ------------------------------------------------------------
Private Sub SearchAllFoldersForOldestBodyMatch(ByVal keyword As String)
    Dim oStore     As Outlook.Store
    Dim rootFolder As Outlook.Folder

    For Each oStore In Application.Session.Stores
        Set rootFolder = oStore.GetRootFolder
        SearchFolderRecursive rootFolder, keyword
    Next oStore
End Sub

Private Sub SearchFolderRecursive(ByVal oFolder As Outlook.Folder, ByVal keyword As String)
    Dim items     As Outlook.Items
    Dim itm       As Object
    Dim i         As Long
    Dim subFolder As Outlook.Folder

    On Error Resume Next
    Set items = oFolder.Items

    If Not items Is Nothing Then
        items.Sort "[ReceivedTime]", False   ' False = ascending = oldest first

        For i = 1 To items.Count
            Set itm = items(i)
            If TypeName(itm) = "MailItem" Then
                If BodyContainsKeyword(itm, keyword) Then
                    If gFoundMail Is Nothing Then
                        Set gFoundMail = itm
                        gFoundFolderPath = oFolder.FolderPath
                    ElseIf itm.ReceivedTime < gFoundMail.ReceivedTime Then
                        Set gFoundMail = itm
                        gFoundFolderPath = oFolder.FolderPath
                    End If
                End If
            End If
        Next i
    End If
    On Error GoTo 0

    For Each subFolder In oFolder.Folders
        SearchFolderRecursive subFolder, keyword
    Next subFolder
End Sub

' ------------------------------------------------------------
' BODY KEYWORD MATCH (case-insensitive)
' ------------------------------------------------------------
Private Function BodyContainsKeyword(ByVal itm As Outlook.MailItem, ByVal keyword As String) As Boolean
    Dim bodyText As String
    On Error Resume Next
    bodyText = itm.Body
    On Error GoTo 0

    If Len(bodyText) = 0 Then
        BodyContainsKeyword = False
    Else
        BodyContainsKeyword = (InStr(1, bodyText, keyword, vbTextCompare) > 0)
    End If
End Function

' ------------------------------------------------------------
' SENDER: display name <email>
' ------------------------------------------------------------
Private Function GetSenderText(ByVal itm As Outlook.MailItem) As String
    Dim sName  As String
    Dim sEmail As String
    On Error Resume Next
    sName  = itm.SenderName
    sEmail = itm.SenderEmailAddress
    On Error GoTo 0

    If Len(Trim(sName)) > 0 And Len(Trim(sEmail)) > 0 Then
        GetSenderText = sName & " <" & sEmail & ">"
    ElseIf Len(Trim(sEmail)) > 0 Then
        GetSenderText = sEmail
    ElseIf Len(Trim(sName)) > 0 Then
        GetSenderText = sName
    Else
        GetSenderText = ""
    End If
End Function

' ------------------------------------------------------------
' UTILITIES
' ------------------------------------------------------------
Private Function ColumnLetterToNumber(ByVal colLetter As String) As Long
    Dim i      As Long
    Dim result As Long
    Dim ch     As String

    colLetter = UCase(Trim(colLetter))
    If Len(colLetter) = 0 Then
        ColumnLetterToNumber = 0
        Exit Function
    End If

    For i = 1 To Len(colLetter)
        ch = Mid(colLetter, i, 1)
        If ch < "A" Or ch > "Z" Then
            ColumnLetterToNumber = 0
            Exit Function
        End If
        result = result * 26 + (Asc(ch) - Asc("A") + 1)
    Next i

    ColumnLetterToNumber = result
End Function

Private Function NzText(ByVal v As Variant) As String
    If IsNull(v) Then
        NzText = ""
    Else
        NzText = Trim(CStr(v))
    End If
End Function
