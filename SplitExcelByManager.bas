' =============================================================================
' SplitExcelByManager  v3.0
' Splits the active sheet into one workbook per unique manager.
' Synchronized feature parity with split_excel_by_manager.py v3.0
' =============================================================================
Sub SplitExcelByManager()

    ' --- Configuration (edit these two constants to match your data) ----------
    Const MANAGER_COL_NAME  As String = "Manager"   ' Header of the manager column
    Const OUTPUT_SUBFOLDER  As String = "manager_reports"  ' Subfolder under workbook path
    ' --------------------------------------------------------------------------

    Dim ws          As Worksheet
    Dim managerCol  As Long
    Dim lastRow     As Long
    Dim managers    As Collection
    Dim cell        As Range
    Dim newWb       As Workbook
    Dim manager     As Variant
    Dim safeName    As String
    Dim outputDir   As String
    Dim successCount As Long
    Dim totalCount  As Long
    Dim hdr         As Range

    ' Guard: workbook must be saved before running
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first before running this macro.", vbExclamation
        Exit Sub
    End If

    Set ws = ActiveSheet

    ' Locate the manager column by header name (not hardcoded position)
    managerCol = 0
    For Each hdr In ws.Rows(1).Cells
        If Trim(hdr.Value) = MANAGER_COL_NAME Then
            managerCol = hdr.Column
            Exit For
        End If
    Next hdr

    If managerCol = 0 Then
        MsgBox "Error: Column '" & MANAGER_COL_NAME & "' not found in row 1." & vbCrLf & _
               "Please update the MANAGER_COL_NAME constant at the top of the macro.", vbCritical
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, managerCol).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "Error: No data rows found (sheet appears empty below the header).", vbCritical
        Exit Sub
    End If

    ' Build output directory
    outputDir = ThisWorkbook.Path & Application.PathSeparator & OUTPUT_SUBFOLDER
    On Error Resume Next
    MkDir outputDir
    On Error GoTo 0

    ' Clear any existing AutoFilter before applying a new one
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' Collect unique, non-blank manager names
    Set managers = New Collection
    On Error Resume Next
    For Each cell In ws.Range(ws.Cells(2, managerCol), ws.Cells(lastRow, managerCol))
        If Not IsEmpty(cell.Value) And Trim(CStr(cell.Value)) <> "" Then
            managers.Add cell.Value, CStr(cell.Value)
        End If
    Next cell
    On Error GoTo 0

    If managers.Count = 0 Then
        MsgBox "Error: Column '" & MANAGER_COL_NAME & "' contains no valid (non-blank) data.", vbCritical
        ws.AutoFilterMode = False
        Exit Sub
    End If

    successCount = 0
    totalCount = managers.Count

    ' Split and save a workbook for each unique manager
    For Each manager In managers

        ' --- Sanitize: replace file-system-invalid characters with underscore ---
        safeName = CStr(manager)
        safeName = Application.WorksheetFunction.Substitute(safeName, "/",  "_")
        safeName = Application.WorksheetFunction.Substitute(safeName, "\", "_")
        safeName = Application.WorksheetFunction.Substitute(safeName, ":",  "_")
        safeName = Application.WorksheetFunction.Substitute(safeName, "*",  "_")
        safeName = Application.WorksheetFunction.Substitute(safeName, "?",  "_")
        safeName = Application.WorksheetFunction.Substitute(safeName, "<",  "_")
        safeName = Application.WorksheetFunction.Substitute(safeName, ">",  "_")
        safeName = Application.WorksheetFunction.Substitute(safeName, "|",  "_")
        safeName = Trim(safeName)

        ' --- Reserved Windows device names (matches Python list) ----------------
        Select Case LCase(safeName)
            Case "con", "prn", "aux", "nul", _
                 "com1", "com2", "com3", "com4", "com5", _
                 "com6", "com7", "com8", "com9", _
                 "lpt1", "lpt2", "lpt3", "lpt4", "lpt5", _
                 "lpt6", "lpt7", "lpt8", "lpt9"
                safeName = "Unknown_Manager"
        End Select

        ' --- Length cap: 200 chars (matches Python) ----------------------------
        If Len(safeName) > 200 Then safeName = Left(safeName, 200)

        Dim outPath As String
        outPath = outputDir & Application.PathSeparator & safeName & "_report.xlsx"

        ' --- Per-manager error handling ----------------------------------------
        On Error GoTo ManagerError

        ws.Rows(1).AutoFilter Field:=managerCol, Criteria1:=manager
        Set newWb = Workbooks.Add
        ws.UsedRange.SpecialCells(xlCellTypeVisible).Copy newWb.Sheets(1).Range("A1")

        ' Auto-fit with explicit width cap (min 8, max 50) — matches Python
        Dim col As Range
        Dim maxLen As Long
        Dim colWidth As Double
        Dim colIdx As Long
        colIdx = 0
        For Each col In newWb.Sheets(1).UsedRange.Columns
            colIdx = colIdx + 1
            maxLen = 0
            Dim c As Range
            For Each c In col.Cells
                If Len(CStr(c.Value)) > maxLen Then maxLen = Len(CStr(c.Value))
            Next c
            colWidth = maxLen + 2
            If colWidth < 8  Then colWidth = 8
            If colWidth > 50 Then colWidth = 50
            newWb.Sheets(1).Columns(colIdx).ColumnWidth = colWidth
        Next col

        newWb.SaveAs outPath, FileFormat:=xlOpenXMLWorkbook
        newWb.Close False

        successCount = successCount + 1
        GoTo NextManager

ManagerError:
        ' Log the error and continue to the next manager
        Dim errMsg As String
        errMsg = Err.Description
        On Error GoTo 0
        If Not newWb Is Nothing Then
            On Error Resume Next
            newWb.Close False
            On Error GoTo 0
        End If
        ' Silently skip this manager (matches Python behaviour)
        GoTo NextManager

NextManager:
        On Error GoTo 0
    Next manager

    ws.AutoFilterMode = False
    MsgBox "Done! " & successCount & "/" & totalCount & " report(s) saved to:" & _
           vbCrLf & outputDir, vbInformation
End Sub
