Option Explicit

Private Const MAX_COL_WIDTH As Double = 55
Private Const DEFAULT_TABLE_STYLE As String = "TableStyleMedium2"

Private Type TRunOptions
    OptionLevel   As Long     ' 1 = Simple, 2 = Advanced, 3 = Keyword/Cell crop
    UseKeywordMode As Boolean ' True = keyword anchor; False = user selects range
    MarkerText    As String   ' Used when UseKeywordMode = True
End Type

' ============================================================
' ENTRY POINT
' ============================================================

Public Sub RunUnifiedDataFormatter_v3()

    Dim ws As Worksheet
    Dim mainChoice As Long
    Dim opt As TRunOptions

    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub

    If Not TryGetLongInput( _
        "Choose an option:" & vbCrLf & vbCrLf & _
        "1 = Simple formatting  (cap/autofit columns, remove duplicates, autofit rows)" & vbCrLf & _
        "2 = Advanced formatting  (Option 1 + text-to-columns, delete blank columns, convert to table)" & vbCrLf & _
        "3 = Crop and format  (Option 2 + crop data range by keyword or cell selection)", _
        "Unified Data Formatter", mainChoice) Then Exit Sub

    opt.OptionLevel = mainChoice

    Select Case mainChoice
        Case 1
            RunOption ws, opt, "Simple formatting completed."
        Case 2
            RunOption ws, opt, "Advanced formatting completed."
        Case 3
            If Not GetCropOptions(opt) Then Exit Sub
            RunOption ws, opt, "Crop and format completed."
        Case Else
            MsgBox "Please enter 1, 2, or 3.", vbExclamation, "Invalid Choice"
    End Select

End Sub

' ============================================================
' SHARED RUNNER
' ============================================================

Private Sub RunOption(ByVal ws As Worksheet, ByRef opt As TRunOptions, ByVal doneMessage As String)

    BeginAppState
    On Error GoTo ErrHandler
    ProcessSheetCore ws, opt
    EndAppState
    MsgBox doneMessage, vbInformation, "Done"
    Exit Sub

ErrHandler:
    EndAppState
    MsgBox "Error: " & Err.Description, vbExclamation, "Macro Stopped"

End Sub

' ============================================================
' OPTION 3 INPUT
' ============================================================

Private Function GetCropOptions(ByRef opt As TRunOptions) As Boolean

    Dim cropChoice As Long

    If Not TryGetLongInput( _
        "Choose crop method:" & vbCrLf & vbCrLf & _
        "1 = Keyword anchor (enter the exact text of the first header cell)" & vbCrLf & _
        "2 = Cell selection (select the top-left header cell when prompted)", _
        "Crop Method", cropChoice) Then Exit Function

    Select Case cropChoice
        Case 1
            If Not TryGetMarkerKeyword(opt.MarkerText) Then Exit Function
            opt.UseKeywordMode = True
        Case 2
            opt.UseKeywordMode = False
        Case Else
            MsgBox "Please enter 1 or 2.", vbExclamation, "Invalid Choice"
            Exit Function
    End Select

    GetCropOptions = True

End Function

' ============================================================
' CORE PIPELINE
' ============================================================

Private Sub ProcessSheetCore(ByVal ws As Worksheet, ByRef opt As TRunOptions)

    Dim lastRow As Long
    Dim lastCol As Long

    If ws Is Nothing Then Exit Sub
    If IsSheetEmpty(ws) Then Exit Sub

    ' --- Option 3: crop first ---
    If opt.OptionLevel = 3 Then
        If opt.UseKeywordMode Then
            CropByKeyword ws, opt.MarkerText
        Else
            CropByCellSelection ws
        End If
        If IsSheetEmpty(ws) Then Exit Sub
    End If

    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedColumn(ws)
    If lastRow = 0 Or lastCol = 0 Then Exit Sub

    ' --- All options: clean text ---
    CleanTextInRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' --- Options 2 and 3: text-to-columns (auto type, no delimiters) ---
    If opt.OptionLevel >= 2 Then
        NormalizeColumns ws, lastRow, lastCol
    End If

    ' --- All options: delete blank rows ---
    DeleteBlankRows ws

    ' --- Options 2 and 3: delete blank columns ---
    If opt.OptionLevel >= 2 Then
        DeleteBlankColumns ws, False
    End If

    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedColumn(ws)
    If lastRow = 0 Or lastCol = 0 Then Exit Sub

    ' --- All options: cap columns at MAX_COL_WIDTH, then autofit columns ---
    CapAndAutofitColumns ws, lastRow, lastCol

    ' --- All options: remove duplicate rows ---
    RemoveDuplicateRows ws, True

    ' --- All options: autofit row heights ---
    ws.Range(ws.Cells(1, 1), ws.Cells(GetLastUsedRow(ws), GetLastUsedColumn(ws))) _
        .EntireRow.AutoFit

    ' --- Options 2 and 3: convert to table ---
    If opt.OptionLevel >= 2 Then
        ConvertUsedRangeToTable ws, True
    End If

End Sub

' ============================================================
' OPTION 3: CROP BY KEYWORD
' ============================================================

Private Sub CropByKeyword(ByVal ws As Worksheet, ByVal markerText As String)

    Dim anchorCell As Range

    Set anchorCell = ws.UsedRange.Find( _
        What:=markerText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If anchorCell Is Nothing Then
        If MsgBox("'" & markerText & "' not found. Continue without cropping?", _
                  vbYesNo + vbQuestion, "Keyword Not Found") = vbNo Then
            Err.Raise vbObjectError + 1100, , "Required keyword not found: " & markerText
        End If
        Exit Sub
    End If

    CropToAnchor ws, anchorCell

End Sub

' ============================================================
' OPTION 3: CROP BY CELL SELECTION
' ============================================================

Private Sub CropByCellSelection(ByVal ws As Worksheet)

    Dim anchorCell As Range
    Dim selectedRange As Range

    On Error Resume Next
    Set selectedRange = Application.InputBox( _
        Prompt:="Select the top-left header cell of the table, then click OK." & vbCrLf & _
                "Cancel = abort", _
        Title:="Select Table Anchor Cell", Type:=8)
    On Error GoTo 0

    If selectedRange Is Nothing Then
        Err.Raise vbObjectError + 1101, , "No anchor cell selected."
    End If

    Set anchorCell = selectedRange.Cells(1, 1)
    CropToAnchor ws, anchorCell

End Sub

' ============================================================
' SHARED CROP LOGIC
' ============================================================

Private Sub CropToAnchor(ByVal ws As Worksheet, ByVal anchorCell As Range)

    Dim headerRow As Long
    Dim anchorCol As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim c As Long
    Dim r As Long
    Dim maxDataCol As Long
    Dim maxDataRow As Long
    Dim syntheticCounter As Long

    headerRow  = anchorCell.Row
    anchorCol  = anchorCell.Column
    lastCol    = GetLastUsedColumn(ws)
    lastRow    = GetLastUsedRow(ws)
    maxDataCol = anchorCol
    maxDataRow = headerRow

    For c = anchorCol To lastCol
        For r = headerRow To lastRow
            If LenB(CStr(ws.Cells(r, c).Value2)) > 0 Then
                If c > maxDataCol Then maxDataCol = c
                If r > maxDataRow Then maxDataRow = r
            End If
        Next r
    Next c

    syntheticCounter = 1
    For c = anchorCol To maxDataCol
        If LenB(CStr(ws.Cells(headerRow, c).Value2)) = 0 Then
            ws.Cells(headerRow, c).Value = "Column" & syntheticCounter
            syntheticCounter = syntheticCounter + 1
        End If
    Next c

    If maxDataRow < lastRow Then _
        ws.Range(ws.Cells(maxDataRow + 1, 1), ws.Cells(lastRow, lastCol)).Clear
    If maxDataCol < lastCol Then _
        ws.Range(ws.Cells(1, maxDataCol + 1), ws.Cells(lastRow, lastCol)).Clear

    If anchorCol > 1 Then ws.Range(ws.Columns(1), ws.Columns(anchorCol - 1)).Delete
    If headerRow > 1 Then ws.Range(ws.Rows(1), ws.Rows(headerRow - 1)).Delete

End Sub

' ============================================================
' SHARED HELPERS
' ============================================================

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
            Tab:=False, Semicolon:=False, Comma:=False, Space:=False, Other:=False, _
            TrailingMinusNumbers:=True
    Next colIndex

End Sub

Private Sub DeleteBlankRows(ByVal ws As Worksheet)

    Dim r As Long
    Dim lastRow As Long

    lastRow = GetLastUsedRow(ws)
    If lastRow = 0 Then Exit Sub

    For r = lastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then ws.Rows(r).Delete
    Next r

End Sub

Private Sub DeleteBlankColumns(ByVal ws As Worksheet, ByVal ignoreHeaderRow As Boolean)

    Dim c As Long
    Dim lastCol As Long
    Dim lastRow As Long

    lastCol = GetLastUsedColumn(ws)
    lastRow = GetLastUsedRow(ws)
    If lastCol = 0 Or lastRow = 0 Then Exit Sub

    For c = lastCol To 1 Step -1
        If ignoreHeaderRow And lastRow >= 2 Then
            If Application.WorksheetFunction.CountA( _
               ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c))) = 0 Then ws.Columns(c).Delete
        Else
            If Application.WorksheetFunction.CountA(ws.Columns(c)) = 0 Then ws.Columns(c).Delete
        End If
    Next c

End Sub

Private Sub CapAndAutofitColumns(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)

    Dim c As Long
    Dim formatRange As Range

    If lastRow = 0 Or lastCol = 0 Then Exit Sub

    Set formatRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    For c = 1 To lastCol
        If ws.Columns(c).ColumnWidth > MAX_COL_WIDTH Then ws.Columns(c).ColumnWidth = MAX_COL_WIDTH
    Next c

    formatRange.WrapText = True
    formatRange.VerticalAlignment = xlVAlignCenter
    formatRange.EntireColumn.AutoFit

    For c = 1 To lastCol
        If ws.Columns(c).ColumnWidth > MAX_COL_WIDTH Then ws.Columns(c).ColumnWidth = MAX_COL_WIDTH
    Next c

End Sub

Private Sub RemoveDuplicateRows(ByVal ws As Worksheet, ByVal hasHeaderRow As Boolean)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim data As Variant
    Dim tempOut() As Variant
    Dim result() As Variant
    Dim dict As Object
    Dim startRow As Long
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
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
    If outRow < lastRow Then _
        ws.Range(ws.Cells(outRow + 1, 1), ws.Cells(lastRow, lastCol)).Clear

End Sub

Private Sub ConvertUsedRangeToTable(ByVal ws As Worksheet, ByVal hasHeaderRow As Boolean)

    Dim dataRange As Range
    Dim tbl As ListObject

    If IsSheetEmpty(ws) Then Exit Sub
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(GetLastUsedRow(ws), GetLastUsedColumn(ws)))
    If IsInTable(dataRange.Cells(1, 1)) Then Exit Sub

    Set tbl = ws.ListObjects.Add(xlSrcRange, dataRange, , IIf(hasHeaderRow, xlYes, xlNo))
    tbl.TableStyle = DEFAULT_TABLE_STYLE

End Sub

' ============================================================
' INPUT / UI HELPERS
' ============================================================

Private Function TryGetLongInput(ByVal promptText As String, ByVal titleText As String, _
                                  ByRef resultValue As Long) As Boolean

    Dim v As Variant

    v = Application.InputBox(Prompt:=promptText, Title:=titleText, Type:=1)
    If VarType(v) = vbBoolean Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    resultValue = CLng(v)
    TryGetLongInput = True

End Function

Private Function TryGetMarkerKeyword(ByRef markerText As String) As Boolean

    Dim v As Variant

    v = Application.InputBox( _
        Prompt:="Enter the exact text of the first header cell of the table." & vbCrLf & _
                "Match is case-insensitive but must equal the whole cell value." & vbCrLf & vbCrLf & _
                "Cancel = abort", _
        Title:="Table Anchor Keyword", Type:=2)

    If VarType(v) = vbBoolean Then Exit Function
    markerText = Trim$(CStr(v))
    TryGetMarkerKeyword = True

End Function

Private Function TryGetYesNoCancel(ByVal promptText As String, ByVal titleText As String, _
                                    ByRef resultValue As Boolean) As Boolean

    Dim response As VbMsgBoxResult

    response = MsgBox(promptText, vbYesNoCancel + vbQuestion, titleText)
    If response = vbCancel Then Exit Function
    resultValue = (response = vbYes)
    TryGetYesNoCancel = True

End Function

' ============================================================
' UTILITY FUNCTIONS
' ============================================================

Private Function IsSheetEmpty(ByVal ws As Worksheet) As Boolean
    IsSheetEmpty = (GetLastUsedRow(ws) = 0 Or GetLastUsedColumn(ws) = 0)
End Function

Private Function IsInTable(ByVal cell As Range) As Boolean
    On Error Resume Next
    IsInTable = Not cell.ListObject Is Nothing
    On Error GoTo 0
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If lastCell Is Nothing Then GetLastUsedRow = 0 Else GetLastUsedRow = lastCell.Row
End Function

Private Function GetLastUsedColumn(ByVal ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If lastCell Is Nothing Then GetLastUsedColumn = 0 Else GetLastUsedColumn = lastCell.Column
End Function

Private Function CleanAndTrim(ByVal val As Variant) As Variant
    Dim s As String
    If IsError(val) Or VarType(val) <> vbString Then
        CleanAndTrim = val
        Exit Function
    End If
    s = CStr(val)
    s = Replace(s, Chr$(160), " ")
    s = Replace(s, Chr$(9), " ")
    CleanAndTrim = Application.WorksheetFunction.Trim(s)
End Function

Private Function NzToString(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Then NzToString = vbNullString Else NzToString = CStr(v)
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
