Option Explicit

Private Const MAX_COL_WIDTH As Double = 55
Private Const SAP_DATE_MARKER As String = "Date"
Private Const SAP_SELECTION_MARKER As String = "Selection No."
Private Const DEFAULT_TABLE_STYLE As String = "TableStyleMedium2"

Private Type TRunOptions
    DoSplit As Boolean
    UseSAPMode As Boolean
    DoDedupe As Boolean
    DoTable As Boolean
    DeleteBlankColsIgnoringHeader As Boolean
    HasHeaderRow As Boolean
    MarkerText As String
End Type

Public Sub RunUnifiedDataFormatter_v3()

    Dim ws As Worksheet
    Dim mainChoice As Long
    Dim opt As TRunOptions

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    If Not TryGetLongInput( _
        "Choose an option:" & vbCrLf & vbCrLf & _
        "1 = Simple formatting (active sheet only)" & vbCrLf & _
        "2 = Advanced formatting (active sheet only)" & vbCrLf & _
        "3 = Advanced formatting + optional split (active sheet only)" & vbCrLf & _
        "4 = SAP output processing (active sheet only)", _
        "Unified Data Formatter", mainChoice) Then Exit Sub

    Select Case mainChoice
        Case 1
            BeginAppState
            On Error GoTo ErrHandlerActiveSheet
            ProcessSheetCore ws, opt
            EndAppState
            MsgBox "Simple formatting completed.", vbInformation, "Done"
            Exit Sub

        Case 2
            BeginAppState
            On Error GoTo ErrHandlerActiveSheet
            ProcessSheetCore ws, opt
            EndAppState
            MsgBox "Advanced formatting completed.", vbInformation, "Done"
            Exit Sub

        Case 3
            If Not TryGetYesNoCancel("Do you want to split a column before formatting?", "Optional Column Split", opt.DoSplit) Then Exit Sub
            BeginAppState
            On Error GoTo ErrHandlerActiveSheet
            ProcessSheetCore ws, opt
            EndAppState
            MsgBox "Advanced formatting completed.", vbInformation, "Done"
            Exit Sub

        Case 4
            If Not GetSAPRunOptions(opt) Then Exit Sub
            BeginAppState
            On Error GoTo ErrHandlerActiveSheet
            ProcessSheetCore ws, opt
            EndAppState
            MsgBox "SAP processing completed.", vbInformation, "Done"
            Exit Sub

        Case Else
            MsgBox "Please enter 1, 2, 3, or 4.", vbExclamation, "Invalid Choice"
            Exit Sub
    End Select

ErrHandlerActiveSheet:
    EndAppState
    MsgBox "Error: " & Err.Description, vbExclamation, "Macro Stopped"

End Sub

Private Function GetSAPRunOptions(ByRef opt As TRunOptions) As Boolean

    If Not TryGetYesNoCancel("Run SAP output mode on the active sheet?", "SAP Mode", opt.UseSAPMode) Then Exit Function
    If Not opt.UseSAPMode Then
        GetSAPRunOptions = True
        Exit Function
    End If

    If Not TryGetSAPMarkerChoice(opt.MarkerText) Then Exit Function
    If Not TryGetYesNoCancel("Does the final dataset have a header row?", "Header Row", opt.HasHeaderRow) Then Exit Function
    If Not TryGetYesNoCancel("Remove duplicate rows?", "Duplicate Removal", opt.DoDedupe) Then Exit Function
    If Not TryGetYesNoCancel("Split a column before final formatting?", "Optional Column Split", opt.DoSplit) Then Exit Function
    If Not TryGetYesNoCancel( _
        "Delete blank columns ignoring row 1 as a header row?" & vbCrLf & vbCrLf & _
        "Yes = Ignore row 1 when checking blank columns" & vbCrLf & _
        "No = Check the entire column", _
        "Blank Column Rule", opt.DeleteBlankColsIgnoringHeader) Then Exit Function
    If Not TryGetYesNoCancel("Convert the final result to an Excel table?", "Convert To Table", opt.DoTable) Then Exit Function

    GetSAPRunOptions = True

End Function

Private Sub ProcessSheetCore(ByVal ws As Worksheet, ByRef opt As TRunOptions)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    If ws Is Nothing Then Exit Sub
    If IsSheetEmpty(ws) Then Exit Sub

    If opt.UseSAPMode And Len(opt.MarkerText) > 0 Then
        CropSheetBeforeMarker ws, opt.MarkerText
        If IsSheetEmpty(ws) Then Exit Sub
    End If

    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedColumn(ws)
    If lastRow = 0 Or lastCol = 0 Then Exit Sub

    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    CleanTextInRange dataRange
    NormalizeColumns ws, lastRow, lastCol

    If opt.DoDedupe Then
        RemoveDuplicateRows ws, opt.HasHeaderRow
        lastRow = GetLastUsedRow(ws)
        lastCol = GetLastUsedColumn(ws)
    End If

    If opt.DoSplit Then
        SplitColumnInteractive ws, lastRow
        lastRow = GetLastUsedRow(ws)
        lastCol = GetLastUsedColumn(ws)
    End If

    DeleteBlankRows ws
    DeleteBlankColumns ws, opt.DeleteBlankColsIgnoringHeader

    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedColumn(ws)
    If lastRow = 0 Or lastCol = 0 Then Exit Sub

    ApplyStandardFormatting ws, lastRow, lastCol
    If opt.DoTable Then ConvertUsedRangeToTable ws, opt.HasHeaderRow

End Sub

Private Sub CropSheetBeforeMarker(ByVal ws As Worksheet, ByVal markerText As String)

    Dim targetCell As Range
    Dim deleteToCol As Long
    Dim deleteToRow As Long

    Set targetCell = ws.UsedRange.Find(What:=markerText, LookIn:=xlValues, LookAt:=xlWhole)
    If targetCell Is Nothing Then
        If MsgBox("'" & markerText & "' not found. Continue without marker trimming?", vbYesNo + vbQuestion, "Marker Not Found") = vbNo Then
            Err.Raise vbObjectError + 1100, , "Required marker not found: " & markerText
        End If
        Exit Sub
    End If

    deleteToRow = targetCell.Row - 1
    deleteToCol = targetCell.Column - 1
    If deleteToCol >= 1 Then ws.Range(ws.Columns(1), ws.Columns(deleteToCol)).Delete
    If deleteToRow >= 1 Then ws.Range(ws.Rows(1), ws.Rows(deleteToRow)).Delete

End Sub

Private Sub CleanTextInRange(ByVal rng As Range)

    Dim cell As Range

    For Each cell In rng.Cells
        If Not IsError(cell.Value) Then
            If VarType(cell.Value) = vbString Then cell.Value = CleanAndTrim(cell.Value)
        End If
    Next cell

End Sub

Private Sub NormalizeColumns(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)

    Dim colIndex As Long

    If lastRow = 0 Or lastCol = 0 Then Exit Sub

    For colIndex = 1 To lastCol
        ws.Range(ws.Cells(1, colIndex), ws.Cells(lastRow, colIndex)).TextToColumns _
            Destination:=ws.Cells(1, colIndex), _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=False, _
            Semicolon:=False, _
            Comma:=False, _
            Space:=False, _
            Other:=False, _
            TrailingMinusNumbers:=True
    Next colIndex

End Sub

Private Sub RemoveDuplicateRows(ByVal ws As Worksheet, ByVal hasHeaderRow As Boolean)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim data As Variant
    Dim result() As Variant
    Dim tempOut() As Variant
    Dim dict As Object
    Dim startRow As Long
    Dim r As Long, c As Long, outRow As Long
    Dim key As String

    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedColumn(ws)
    If lastRow <= 1 Or lastCol = 0 Then Exit Sub

    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    data = dataRange.Value2
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    ReDim tempOut(1 To lastRow, 1 To lastCol)
    startRow = IIf(hasHeaderRow, 2, 1)
    outRow = 0

    If hasHeaderRow Then
        outRow = 1
        For c = 1 To lastCol
            tempOut(outRow, c) = data(1, c)
        Next c
    End If

    For r = startRow To lastRow
        key = vbNullString
        For c = 1 To lastCol
            key = key & Chr$(30) & NzToString(data(r, c))
        Next c
        If Not dict.Exists(key) Then
            dict.Add key, True
            outRow = outRow + 1
            For c = 1 To lastCol
                tempOut(outRow, c) = data(r, c)
            Next c
        End If
    Next r

    ReDim result(1 To outRow, 1 To lastCol)
    For r = 1 To outRow
        For c = 1 To lastCol
            result(r, c) = tempOut(r, c)
        Next c
    Next r

    dataRange.ClearContents
    ws.Range(ws.Cells(1, 1), ws.Cells(outRow, lastCol)).Value = result
    If outRow < lastRow Then ws.Range(ws.Cells(outRow + 1, 1), ws.Cells(lastRow, lastCol)).Clear

End Sub

Private Sub DeleteBlankRows(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim r As Long

    lastRow = GetLastUsedRow(ws)
    If lastRow = 0 Then Exit Sub

    For r = lastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then ws.Rows(r).Delete
    Next r

End Sub

Private Sub DeleteBlankColumns(ByVal ws As Worksheet, ByVal ignoreHeaderRow As Boolean)

    Dim lastCol As Long
    Dim lastRow As Long
    Dim c As Long

    lastCol = GetLastUsedColumn(ws)
    lastRow = GetLastUsedRow(ws)
    If lastCol = 0 Or lastRow = 0 Then Exit Sub

    For c = lastCol To 1 Step -1
        If ignoreHeaderRow And lastRow >= 2 Then
            If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c))) = 0 Then ws.Columns(c).Delete
        Else
            If Application.WorksheetFunction.CountA(ws.Columns(c)) = 0 Then ws.Columns(c).Delete
        End If
    Next c

End Sub

Private Sub ApplyStandardFormatting(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)

    Dim c As Long
    Dim formatRange As Range

    If lastRow = 0 Or lastCol = 0 Then Exit Sub

    Set formatRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    formatRange.EntireColumn.AutoFit

    For c = 1 To lastCol
        If ws.Columns(c).ColumnWidth > MAX_COL_WIDTH Then ws.Columns(c).ColumnWidth = MAX_COL_WIDTH
    Next c

    formatRange.WrapText = True
    formatRange.VerticalAlignment = xlVAlignCenter
    formatRange.EntireRow.AutoFit

End Sub

Private Sub SplitColumnInteractive(ByVal ws As Worksheet, ByVal lastRow As Long)

    Dim splitCol As Long
    Dim delimiterChoice As Long
    Dim otherDelimiter As String
    Dim targetCol As Range
    Dim useTab As Boolean, useSemicolon As Boolean, useComma As Boolean, useSpace As Boolean, useOther As Boolean

    If Not TryGetLongInput("Enter the column number to split (example: 3 for column C).", "Column to Split", splitCol) Then Exit Sub
    If splitCol < 1 Or splitCol > 16384 Then
        MsgBox "Invalid column number.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    If Not TryGetLongInput( _
        "Choose delimiter:" & vbCrLf & vbCrLf & _
        "1 = Tab" & vbCrLf & _
        "2 = Semicolon" & vbCrLf & _
        "3 = Comma" & vbCrLf & _
        "4 = Space" & vbCrLf & _
        "5 = Other character", _
        "Split Delimiter", delimiterChoice) Then Exit Sub

    ResetDelimiterFlags useTab, useSemicolon, useComma, useSpace, useOther

    Select Case delimiterChoice
        Case 1: useTab = True
        Case 2: useSemicolon = True
        Case 3: useComma = True
        Case 4: useSpace = True
        Case 5
            If Not TryGetTextInput("Enter the single character delimiter.", "Other Delimiter", otherDelimiter) Then Exit Sub
            If Len(otherDelimiter) = 0 Then
                MsgBox "Delimiter cannot be blank.", vbExclamation, "Invalid Input"
                Exit Sub
            End If
            useOther = True
            otherDelimiter = Left$(otherDelimiter, 1)
        Case Else
            MsgBox "Invalid delimiter choice.", vbExclamation, "Invalid Input"
            Exit Sub
    End Select

    Set targetCol = ws.Range(ws.Cells(1, splitCol), ws.Cells(lastRow, splitCol))
    targetCol.TextToColumns _
        Destination:=targetCol.Cells(1, 1), _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=useTab, _
        Semicolon:=useSemicolon, _
        Comma:=useComma, _
        Space:=useSpace, _
        Other:=useOther, _
        OtherChar:=IIf(useOther, otherDelimiter, False), _
        TrailingMinusNumbers:=True

End Sub

Private Sub ConvertUsedRangeToTable(ByVal ws As Worksheet, ByVal hasHeaderRow As Boolean)

    Dim dataRange As Range
    Dim tbl As ListObject
    Dim headerSetting As XlYesNoGuess

    If GetLastUsedRow(ws) = 0 Or GetLastUsedColumn(ws) = 0 Then Exit Sub
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(GetLastUsedRow(ws), GetLastUsedColumn(ws)))
    If dataRange.Rows.Count = 0 Or dataRange.Columns.Count = 0 Then Exit Sub
    If IsInTable(dataRange.Cells(1, 1)) Then Exit Sub

    headerSetting = IIf(hasHeaderRow, xlYes, xlNo)
    Set tbl = ws.ListObjects.Add(xlSrcRange, dataRange, , headerSetting)
    tbl.TableStyle = DEFAULT_TABLE_STYLE

End Sub

Private Function TryGetLongInput(ByVal promptText As String, ByVal titleText As String, ByRef resultValue As Long) As Boolean

    Dim v As Variant

    v = Application.InputBox(Prompt:=promptText, Title:=titleText, Type:=1)
    If VarType(v) = vbBoolean Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    resultValue = CLng(v)
    TryGetLongInput = True

End Function

Private Function TryGetTextInput(ByVal promptText As String, ByVal titleText As String, ByRef resultValue As String) As Boolean

    Dim v As Variant

    v = Application.InputBox(Prompt:=promptText, Title:=titleText, Type:=2)
    If VarType(v) = vbBoolean Then Exit Function
    resultValue = CStr(v)
    TryGetTextInput = True

End Function

Private Function TryGetYesNoCancel(ByVal promptText As String, ByVal titleText As String, ByRef resultValue As Boolean) As Boolean

    Dim response As VbMsgBoxResult

    response = MsgBox(promptText, vbYesNoCancel + vbQuestion, titleText)
    If response = vbCancel Then Exit Function
    resultValue = (response = vbYes)
    TryGetYesNoCancel = True

End Function

Private Function TryGetSAPMarkerChoice(ByRef markerText As String) As Boolean

    Dim response As VbMsgBoxResult

    response = MsgBox( _
        "Which SAP marker should be used for trimming the leading rows/columns?" & vbCrLf & vbCrLf & _
        "Yes = 'Selection No.'" & vbCrLf & _
        "No = 'Date'" & vbCrLf & _
        "Cancel = Skip SAP marker trimming", _
        vbYesNoCancel + vbQuestion, _
        "SAP Marker")

    Select Case response
        Case vbYes:    markerText = SAP_SELECTION_MARKER
        Case vbNo:     markerText = SAP_DATE_MARKER
        Case vbCancel: markerText = vbNullString
    End Select

    TryGetSAPMarkerChoice = True

End Function

Private Function IsInTable(ByVal cell As Range) As Boolean

    On Error Resume Next
    IsInTable = Not cell.ListObject Is Nothing
    On Error GoTo 0

End Function

Private Function IsSheetEmpty(ByVal ws As Worksheet) As Boolean

    IsSheetEmpty = (GetLastUsedRow(ws) = 0 Or GetLastUsedColumn(ws) = 0)

End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = lastCell.Row
    End If

End Function

Private Function GetLastUsedColumn(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        GetLastUsedColumn = 0
    Else
        GetLastUsedColumn = lastCell.Column
    End If

End Function

Private Function CleanAndTrim(ByVal val As Variant) As Variant

    Dim tempValue As String

    If IsError(val) Then
        CleanAndTrim = val
        Exit Function
    End If
    If VarType(val) <> vbString Then
        CleanAndTrim = val
        Exit Function
    End If

    tempValue = CStr(val)
    tempValue = Replace(tempValue, Chr$(160), " ")
    tempValue = Replace(tempValue, Chr$(9), " ")
    tempValue = Application.WorksheetFunction.Trim(tempValue)
    CleanAndTrim = tempValue

End Function

Private Function NzToString(ByVal v As Variant) As String

    If IsError(v) Or IsNull(v) Then
        NzToString = vbNullString
    Else
        NzToString = CStr(v)
    End If

End Function

Private Sub ResetDelimiterFlags(ByRef useTab As Boolean, ByRef useSemicolon As Boolean, ByRef useComma As Boolean, ByRef useSpace As Boolean, ByRef useOther As Boolean)

    useTab = False
    useSemicolon = False
    useComma = False
    useSpace = False
    useOther = False

End Sub

Private Sub BeginAppState()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Sub

Private Sub EndAppState()

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
