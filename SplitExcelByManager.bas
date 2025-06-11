Sub SplitExcelByManager()
    Dim ws As Worksheet
    Dim managerCol As Long
    Dim lastRow As Long
    Dim managers As Collection
    Dim cell As Range
    Dim newWb As Workbook
    Dim manager As Variant

    Set ws = ActiveSheet
    managerCol = 1 'Assume manager is in column A
    lastRow = ws.Cells(ws.Rows.Count, managerCol).End(xlUp).Row

    Set managers = New Collection
    On Error Resume Next
    For Each cell In ws.Range(ws.Cells(2, managerCol), ws.Cells(lastRow, managerCol))
        managers.Add cell.Value, CStr(cell.Value)
    Next cell
    On Error GoTo 0

    For Each manager In managers
        ws.Rows(1).AutoFilter Field:=managerCol, Criteria1:=manager
        Set newWb = Workbooks.Add
        ws.UsedRange.SpecialCells(xlCellTypeVisible).Copy newWb.Sheets(1).Range("A1")
        newWb.Sheets(1).Columns.AutoFit
        newWb.SaveAs manager & "_Report.xlsx"
        newWb.Close False
    Next manager

    ws.AutoFilterMode = False
End Sub
