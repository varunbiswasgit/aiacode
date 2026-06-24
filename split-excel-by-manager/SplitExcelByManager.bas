Sub SplitExcelByManager_UserInputs()

    Dim ws As Worksheet, managerHeader As String, outputDir As String
    Dim managerCol As Long, lastRow As Long, lastCol As Long
    Dim managers As Object, cell As Range, manager As Variant
    Dim newWb As Workbook, safeName As String, outPath As String
    Dim fd As FileDialog

    Set ws = ActiveSheet

    Dim headerCell As Range

    Set headerCell = Application.InputBox( _
    Prompt:="Select the manager column header cell:", Title:="Manager Column", Type:=8)

    If headerCell Is Nothing Then Exit Sub
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select folder to save manager files"
    If fd.Show <> -1 Then Exit Sub
    outputDir = fd.SelectedItems(1)

    managerCol = headerCell.Column
    lastRow = ws.Cells(ws.Rows.Count, managerCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set managers = CreateObject("Scripting.Dictionary")

    For Each cell In ws.Range(ws.Cells(2, managerCol), ws.Cells(lastRow, managerCol))
        If Trim(cell.Value) <> "" Then managers(Trim(cell.Value)) = 1
    Next cell

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    For Each manager In managers.Keys
        safeName = CleanFileName(CStr(manager))
        outPath = outputDir & Application.PathSeparator & safeName & "_report.xlsx"

        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter _
            Field:=managerCol, Criteria1:=manager

        Set newWb = Workbooks.Add

         ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)) _
        .SpecialCells(xlCellTypeVisible).Copy
    
        With newWb.Sheets(1).Range("A1")
            .PasteSpecial Paste:=xlPasteAll
            .PasteSpecial Paste:=xlPasteColumnWidths
        End With

       With newWb.Sheets(1).UsedRange
            .WrapText = False
            .Columns.AutoFit
            .Rows.AutoFit
        End With
        
        Application.DisplayAlerts = False
        newWb.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        newWb.Close False
    Next manager

    ws.AutoFilterMode = False
    Application.CutCopyMode = False

    MsgBox "Done. Created " & managers.Count & " files in:" & vbCrLf & outputDir

End Sub

Function CleanFileName(s As String) As String
    Dim badChars, ch
    badChars = Array("/", "\", ":", "*", "?", """", "<", ">", "|")

    For Each ch In badChars
        s = Replace(s, ch, "_")
    Next ch

    s = Trim(s)
    If Len(s) = 0 Then s = "Unknown_Manager"
    If Len(s) > 150 Then s = Left(s, 150)

    CleanFileName = s
End Function

