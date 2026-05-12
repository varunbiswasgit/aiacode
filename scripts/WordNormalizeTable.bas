Attribute VB_Name = "WordNormalizeTable"
Option Explicit

' ============================
' WordNormalizeTable.bas
' Mirrors manual Table Properties:
'   Table  : 100% width, left align, no text wrap
'   Row    : uncheck Specify Height
'   Column : uncheck Preferred Width
'   Cell   : uncheck Preferred Width, vertical align center
' Covers all tables including header/footer ranges.
' ============================

Sub NormalizeTables_Light()
    Dim doc As Document
    Dim tbl As Table
    Dim rw As Row
    Dim col As Column
    Dim cel As Cell

    Set doc = ActiveDocument

    For Each tbl In doc.Tables

        ' --- TABLE: 100% width, left alignment, no text wrapping ---
        With tbl
            .PreferredWidthType = 3     ' wdPreferPercent
            .PreferredWidth = 100
            .Alignment = 0              ' wdAlignRowLeft
            .AllowAutoFit = True
            .TextWrapping = 0           ' wdTableTextWrappingNone
        End With

        ' --- ROW: uncheck Specify Height ---
        For Each rw In tbl.Rows
            rw.HeightRule = 0           ' wdRowHeightAuto (unchecks Specify Height)
        Next rw

        ' --- COLUMN: uncheck Preferred Width ---
        For Each col In tbl.Columns
            col.PreferredWidthType = 0  ' wdPreferNone (unchecks Preferred Width)
        Next col

        ' --- CELL: uncheck Preferred Width, vertical align center ---
        For Each cel In tbl.Range.Cells
            cel.PreferredWidthType = 0  ' wdPreferNone (unchecks Preferred Width)
            cel.VerticalAlignment = 1   ' wdCellAlignVerticalCenter
        Next cel

    Next tbl

    MsgBox "Done. " & doc.Tables.Count & " table(s) normalized.", vbInformation
End Sub
