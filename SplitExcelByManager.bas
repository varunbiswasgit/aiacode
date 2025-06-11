Option Explicit

Sub SplitExcelByManager()
    Dim src As Worksheet
    Dim headerRow As Range
    Dim managerHeader As String
    Dim managerColIndex As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    Dim cell As Range
    Dim managers As Object
    Dim mgr As Variant
    Dim savePath As String

    Set src = ActiveSheet
    Set headerRow = src.Rows(1)

    managerHeader = InputBox("Enter the manager column header:", _
                             "Manager Column", "Manager")

    managerColIndex = 0
    For Each cell In headerRow.Cells
        If Trim(cell.Value) = managerHeader Then
            managerColIndex = cell.Column
            Exit For
        End If
        If cell.Value = "" Then Exit For
    Next cell

    If managerColIndex = 0 Then
        MsgBox "Column '" & managerHeader & "' not found.", vbCritical
        Exit Sub
    End If

    lastRow = src.Cells(src.Rows.Count, managerColIndex).End(xlUp).Row
    lastCol = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    Set dataRange = src.Range(src.Cells(1, 1), src.Cells(lastRow, lastCol))

    Set managers = CreateObject("Scripting.Dictionary")
    For Each cell In src.Range(src.Cells(2, managerColIndex), src.Cells(lastRow, managerColIndex))
        If Trim(cell.Value) <> "" Then managers(cell.Value) = 1
    Next cell

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder to save split workbooks"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            MsgBox "No folder selected. Operation canceled.", vbExclamation
            Exit Sub
        End If
        savePath = .SelectedItems(1)
    End With

    For Each mgr In managers.Keys
        src.AutoFilterMode = False
        dataRange.AutoFilter Field:=managerColIndex, Criteria1:=mgr

        Dim newWb As Workbook
        Set newWb = Workbooks.Add(xlWBATWorksheet)
        dataRange.SpecialCells(xlCellTypeVisible).Copy newWb.Sheets(1).Range("A1")

        newWb.SaveAs Filename:=savePath & "\" & SanitizeFileName(CStr(mgr)) & ".xlsx", _
                     FileFormat:=xlOpenXMLWorkbook
        newWb.Close SaveChanges:=False
    Next mgr

    src.AutoFilterMode = False
    MsgBox managers.Count & " file(s) saved to " & savePath, vbInformation
End Sub

Function SanitizeFileName(name As String) As String
    Dim invalidChars As Variant
    Dim ch As Variant
    invalidChars = Array("\\", "/", ":", "*", "?", "\""", "<", ">", "|")
    For Each ch In invalidChars
        name = Replace(name, ch, "_")
    Next ch
    name = Trim(name)
    If Len(name) = 0 Then name = "Unknown_Manager"
    SanitizeFileName = name
End Function
