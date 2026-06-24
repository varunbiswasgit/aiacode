Option Explicit
Private Const xlWBATWorksheet As Long = -4167

Public Sub BuildManagerResponseTracker()
    Dim olApp As Object, olNs As Object, olFldr As Object
    Dim xlApp As Object, srcWb As Object, outWb As Object, srcWs As Object, outWs As Object
    Dim excelPath As String, excludeSender As String, colInput As String, outFile As String
    Dim colNum As Long, lastRow As Long, r As Long, outRow As Long
    Dim dictManagers As Object, dictLatest As Object, managerEmail As Variant
    Dim items As Object, itm As Object, mail As Object
    Dim senderAddr As String, attName As String
    Dim latestInfo As Object, saveFolder As String

    Set olApp = CreateObject("Outlook.Application")
    Set olNs = olApp.GetNamespace("MAPI")
    Set olFldr = olNs.PickFolder
    If olFldr Is Nothing Then Exit Sub

    excludeSender = LCase$(Trim$(InputBox("Enter the sender email address to exclude:", "Exclude Sender")))
    If Len(excludeSender) = 0 Then Exit Sub

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    excelPath = PickExcelFile(xlApp)
    If Len(excelPath) = 0 Then GoTo CleanExit

    Set srcWb = xlApp.Workbooks.Open(excelPath, False, True)
    Set srcWs = srcWb.Worksheets(1)
    colInput = InputBox("Enter the column letter containing manager emails (e.g. B):", "Manager Email Column")
    If Len(Trim$(colInput)) = 0 Then GoTo CleanExit
    colNum = ColumnToNumber(colInput)
    If colNum < 1 Then GoTo CleanExit

    lastRow = GetLastUsedRow(srcWs, colNum)
    Set dictManagers = CreateObject("Scripting.Dictionary")
    Set dictLatest = CreateObject("Scripting.Dictionary")

    For r = 2 To lastRow
        managerEmail = LCase$(Trim$(CStr(srcWs.Cells(r, colNum).Value)))
        If Len(managerEmail) > 0 Then
            If Not dictManagers.Exists(CStr(managerEmail)) Then dictManagers.Add CStr(managerEmail), True
        End If
    Next r

    Set items = olFldr.Items
    items.Sort "[ReceivedTime]", True

    For Each itm In items
        If TypeName(itm) = "MailItem" Then
            Set mail = itm
            senderAddr = LCase$(GetSenderSmtpAddress(mail))
            If Len(senderAddr) > 0 And senderAddr <> excludeSender Then
                If dictManagers.Exists(senderAddr) Then
                    If Not dictLatest.Exists(senderAddr) Then
                        Set latestInfo = CreateObject("Scripting.Dictionary")
                        latestInfo.Add "Time", mail.ReceivedTime
                        latestInfo.Add "Subject", mail.Subject
                        latestInfo.Add "HasAtt", False
                        latestInfo.Add "AttName", ""
                        latestInfo.Add "ExcelSeen", False
                        If HasExcelAttachment(mail, attName) Then
                            latestInfo("HasAtt") = True
                            latestInfo("AttName") = attName
                            latestInfo("ExcelSeen") = True
                        End If
                        dictLatest.Add senderAddr, latestInfo
                    Else
                        If HasExcelAttachment(mail, attName) Then
                            dictLatest(senderAddr)("ExcelSeen") = True
                            If Len(dictLatest(senderAddr)("AttName")) = 0 Then dictLatest(senderAddr)("AttName") = attName
                        End If
                    End If
                End If
            End If
        End If
    Next itm

    saveFolder = GetSafeFolderPath(srcWb)
    If Len(saveFolder) = 0 Then saveFolder = CurDir
    outFile = saveFolder & "\Manager_Response_Tracke.xlsx"

    If IsFileOpen(outFile) Then
        MsgBox "Please close Manager_Response_Tracke.xlsx before running the macro.", vbExclamation
        GoTo CleanExit
    End If

    Set outWb = xlApp.Workbooks.Add(xlWBATWorksheet)
    Set outWs = outWb.Worksheets(1)

    outWs.Range("A1:H1").Value = Array("Manager Email", "Response Received", "Latest Email Time", "Latest Email Subject", "Last Email has Attachment?", "Excel Seen In Thread", "Newest Attachment Name", "Clarification Required")

    outRow = 2
    For Each managerEmail In dictManagers.Keys
        outWs.Cells(outRow, 1).Value = CStr(managerEmail)
        If dictLatest.Exists(CStr(managerEmail)) Then
            outWs.Cells(outRow, 2).Value = "Yes"
            outWs.Cells(outRow, 3).Value = dictLatest(CStr(managerEmail))("Time")
            outWs.Cells(outRow, 4).Value = dictLatest(CStr(managerEmail))("Subject")
            outWs.Cells(outRow, 5).Value = IIf(dictLatest(CStr(managerEmail))("HasAtt"), "Yes", "No")
            outWs.Cells(outRow, 6).Value = IIf(dictLatest(CStr(managerEmail))("ExcelSeen"), "Yes", "No")
            outWs.Cells(outRow, 7).Value = dictLatest(CStr(managerEmail))("AttName")
            outWs.Cells(outRow, 8).Value = IIf(dictLatest(CStr(managerEmail))("HasAtt"), "No", "Yes")
        Else
            outWs.Cells(outRow, 2).Value = "No"
            outWs.Cells(outRow, 5).Value = "No"
            outWs.Cells(outRow, 6).Value = "No"
            outWs.Cells(outRow, 8).Value = "No"
        End If
        outRow = outRow + 1
    Next managerEmail

    outWs.Columns.AutoFit
    outWb.SaveAs outFile, 51
    outWb.Close True

CleanExit:
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
End Sub

Private Function PickExcelFile(ByVal xlApp As Object) As String
    Dim fd As Object
    Set fd = xlApp.Application.FileDialog(3)
    fd.Title = "Select Manager List Excel File"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then PickExcelFile = fd.SelectedItems(1)
End Function

Private Function ColumnToNumber(ByVal colInput As String) As Long
    Dim i As Long, result As Long, ch As String
    colInput = UCase$(Trim$(colInput))
    For i = 1 To Len(colInput)
        ch = Mid$(colInput, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
        result = result * 26 + (Asc(ch) - 64)
    Next i
    ColumnToNumber = result
End Function

Private Function GetLastUsedRow(ByVal ws As Object, ByVal colNum As Long) As Long
    Dim rng As Object
    On Error Resume Next
    Set rng = ws.Columns(colNum).Find(What:="*", LookIn:=-4163, LookAt:=2, SearchOrder:=1, SearchDirection:=2, MatchCase:=False)
    If rng Is Nothing Then
        GetLastUsedRow = 1
    Else
        GetLastUsedRow = rng.Row
    End If
End Function

Private Function GetSafeFolderPath(ByVal wb As Object) As String
    On Error Resume Next
    Dim p As String
    p = wb.Path
    If Len(p) > 0 Then
        GetSafeFolderPath = p
    Else
        GetSafeFolderPath = ""
    End If
End Function

Private Function IsFileOpen(ByVal filePath As String) As Boolean
    Dim f As Integer
    On Error GoTo Locked
    If Len(Dir(filePath)) = 0 Then Exit Function
    f = FreeFile
    Open filePath For Binary Access Read Write Lock Read Write As #f
    Close #f
    Exit Function
Locked:
    IsFileOpen = True
    On Error Resume Next
    Close #f
End Function

Private Function GetSenderSmtpAddress(ByVal mail As Object) As String
    On Error Resume Next
    Dim addr As String, pa As Object
    addr = mail.SenderEmailAddress
    If LCase$(mail.SenderEmailType) = "ex" Then
        Set pa = mail.PropertyAccessor
        addr = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001E")
    End If
    GetSenderSmtpAddress = addr
End Function

Private Function HasExcelAttachment(ByVal mail As Object, ByRef firstName As String) As Boolean
    Dim i As Long, fn As String
    firstName = ""
    For i = 1 To mail.Attachments.Count
        fn = LCase$(mail.Attachments.Item(i).FileName)
        If Right$(fn, 5) = ".xlsx" Then
            HasExcelAttachment = True
            firstName = mail.Attachments.Item(i).FileName
            Exit Function
        End If
    Next i
End Function
