Sub SplitExcelByManager()
    Dim ws As Worksheet
    Dim managerCol As Long
    Dim lastRow As Long
    Dim managers As Collection
    Dim cell As Range
    Dim newWb As Workbook
    Dim manager As Variant
    Dim safeName As String

    ' Guard: workbook must be saved before running
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first before running this macro.", vbExclamation
        Exit Sub
    End If

    Set ws = ActiveSheet
    managerCol = 1 ' Manager is in column A
    lastRow = ws.Cells(ws.Rows.Count, managerCol).End(xlUp).Row

    ' Clear any existing AutoFilter before applying a new one
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' Collect unique manager names using Collection key deduplication
    Set managers = New Collection
    On Error Resume Next
    For Each cell In ws.Range(ws.Cells(2, managerCol), ws.Cells(lastRow, managerCol))
        If Not IsEmpty(cell.Value) And Trim(CStr(cell.Value)) <> "" Then
            managers.Add cell.Value, CStr(cell.Value)
        End If
    Next cell
    On Error GoTo 0

    ' Split and save a workbook for each unique manager
    For Each manager In managers
        ' Sanitize manager name: replace characters invalid in file names
        safeName = manager
        safeName = Application.WorksheetFunction.Substitute(safeName, "/", "-")
        safeName = Application.WorksheetFunction.Substitute(safeName, "\", "-")
        safeName = Application.WorksheetFunction.Substitute(safeName, ":", "-")
        safeName = Application.WorksheetFunction.Substitute(safeName, "*", "-")
        safeName = Application.WorksheetFunction.Substitute(safeName, "?", "-")
        safeName = Application.WorksheetFunction.Substitute(safeName, "<", "-")
        safeName = Application.WorksheetFunction.Substitute(safeName, ">", "-")
        safeName = Application.WorksheetFunction.Substitute(safeName, "|", "-")
        safeName = Trim(safeName)

        ws.Rows(1).AutoFilter Field:=managerCol, Criteria1:=manager
        Set newWb = Workbooks.Add
        ws.UsedRange.SpecialCells(xlCellTypeVisible).Copy newWb.Sheets(1).Range("A1")
        newWb.Sheets(1).Columns.AutoFit

        ' Save to the same folder as the source workbook
        newWb.SaveAs ThisWorkbook.Path & Application.PathSeparator & safeName & "_Report.xlsx", FileFormat:=xlOpenXMLWorkbook
        newWb.Close False
    Next manager

    ws.AutoFilterMode = False
    MsgBox "Done! " & managers.Count & " report(s) saved to:" & vbCrLf & ThisWorkbook.Path, vbInformation
End Sub
