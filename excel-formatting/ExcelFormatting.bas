Option Explicit

Private Const MAX_COL_WIDTH As Double = 55
Private Const DEFAULT_TABLE_STYLE As String = "TableStyleMedium2"

Private Type TRunOptions
    UseCropSelection As Boolean
    CropStartRow As Long
    CropStartCol As Long
    DoDedupe As Boolean
    DoTable As Boolean
    DeleteBlankColsIgnoringHeader As Boolean
    HasHeaderRow As Boolean
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
        "3 = Generic formatting with data cropping (active sheet only)", _
        "Unified Data Formatter", mainChoice) Then Exit Sub

    Select Case mainChoice
        Case 1
            ' Simple formatting only
            opt.UseCropSelection = False
            opt.HasHeaderRow = False
            opt.DoDedupe = False
            opt.DeleteBlankColsIgnoringHeader = False
            opt.DoTable = False

            BeginAppState
            On Error GoTo ErrHandlerActiveSheet
            ProcessSheetCore ws, opt
            EndAppState
            MsgBox "Simple formatting completed.", vbInformation, "Done"
            Exit Sub

        Case 2
            ' Advanced formatting = everything from option 3 except cropping
            opt.UseCropSelection = False
            opt.HasHeaderRow = True
            opt.DoDedupe = True
            opt.DeleteBlankColsIgnoringHeader = True
            opt.DoTable = True

            BeginAppState
            On Error GoTo ErrHandlerActiveSheet
            ProcessSheetCore ws, opt
            EndAppState
            MsgBox "Advanced formatting completed.", vbInformation, "Done"
            Exit Sub

        Case 3
            ' Generic formatting with cropping
            If Not GetGenericCropRunOptions(ws, opt) Then Exit Sub

            BeginAppState
            On Error GoTo ErrHandlerActiveSheet
            ProcessSheetCore ws, opt
            EndAppState
            MsgBox "Generic formatting completed.", vbInformation, "Done"
            Exit Sub

        Case Else
            MsgBox "Please enter 1, 2, or 3.", vbExclamation, "Invalid Choice"
            Exit Sub
    End Select

ErrHandlerActiveSheet:
    EndAppState
    MsgBox "Error: " & Err.Description, vbExclamation, "Macro Stopped"

End Sub

Private Function GetGenericCropRunOptions(ByVal ws As Worksheet, ByRef opt As TRunOptions) As Boolean

    Dim startCell As Range

    Set startCell = GetCroppingStartCell(ws)
    If startCell Is Nothing Then Exit Function

    opt.UseCropSelection = True
    opt.CropStartRow = startCell.Row
    opt.CropStartCol = startCell.Column

    ' Same settings as option 2
    opt.HasHeaderRow = True
    opt.DoDedupe = True
    opt.DeleteBlankColsIgnoringHeader = True
    opt.DoTable = True

    GetGenericCropRunOptions = True

End Function

Private Function GetCroppingStartCell(ByVal ws As Worksheet) As Range

    Dim pickedRange As Range

    On Error Resume Next
    Set pickedRange = Application.InputBox( _
        Prompt:="Select the first cell where the actual data starts." & vbCrLf & vbCrLf & _
                "Everything above and to the left of this cell will be removed.", _
        Title:="Select Data Start Cell", _
        Type:=8)
    On Error GoTo 0

    If pickedRange Is Nothing Then Exit Function

    Set pickedRange = pickedRange.Cells(1, 1)

    If Not pickedRange.Worksheet Is ws Then
        MsgBox "Please select a cell on the active sheet.", vbExclamation, "Invalid Selection"
        Exit Function
    End If

    Set GetCroppingStartCell = pickedRange

End Function

Private Sub ProcessSheetCore(ByVal ws As Worksheet, ByRef opt As TRunOptions)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    If ws Is Nothing Then Exit Sub
    If IsSheetEmpty(ws) Then Exit Sub

    If opt.UseCropSelection Then
        CropSheetToStartCell ws, opt.CropStartRow, opt.CropStartCol
        If IsSheetEmpty(ws) Then Exit Sub
    End If

    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedColumn(ws)
    If lastRow = 0 Or lastCol = 0 Then Exit Sub

    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
	CleanTextInRange dataRange
	' IMPORTANT: clean header again explicitly
	If opt.HasHeaderRow Then
		CleanTextInRange ws.Rows(1)
	End If
	If opt.DoDedupe Then
		RemoveDuplicateRows ws, opt.HasHeaderRow
	End If
	' Normalize AFTER dedupe (prevents header/data mismatch issues)
	NormalizeColumns ws, GetLastUsedRow(ws), GetLastUsedColumn(ws)

    DeleteBlankRows ws
    DeleteBlankColumns ws, opt.DeleteBlankColsIgnoringHeader

    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedColumn(ws)
    If lastRow = 0 Or lastCol = 0 Then Exit Sub

    ApplyStandardFormatting ws, lastRow, lastCol

    If opt.DoTable Then
        ConvertUsedRangeToTable ws, opt.HasHeaderRow
    End If

End Sub

Private Sub CropSheetToStartCell(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)

    Dim deleteToRow As Long
    Dim deleteToCol As Long

    deleteToRow = startRow - 1
    deleteToCol = startCol - 1

    If deleteToCol >= 1 Then
        ws.Range(ws.Columns(1), ws.Columns(deleteToCol)).Delete
    End If

    If deleteToRow >= 1 Then
        ws.Range(ws.Rows(1), ws.Rows(deleteToRow)).Delete
    End If

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

    If outRow < lastRow Then
        ws.Range(ws.Cells(outRow + 1, 1), ws.Cells(lastRow, lastCol)).Clear
    End If

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
    Dim r As Long
    Dim hasData As Boolean
    Dim cellVal As String

    lastCol = GetLastUsedColumn(ws)
    lastRow = GetLastUsedRow(ws)
    If lastCol = 0 Or lastRow = 0 Then Exit Sub

    For c = lastCol To 1 Step -1

        hasData = False

        For r = IIf(ignoreHeaderRow, 2, 1) To lastRow
            
            If Not IsError(ws.Cells(r, c).Value) Then
                cellVal = Trim$(Replace(ws.Cells(r, c).Value, Chr(160), ""))
                
                If Len(cellVal) > 0 Then
                    hasData = True
                    Exit For
                End If
            End If

        Next r

        If Not hasData Then
            ws.Columns(c).Delete
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
