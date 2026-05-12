'Attribute VB_Name = "WordNormalizeTable"
Option Explicit

' ============================
' WordNormalizeTable.bas
' Light, fast Word table normalization macro.
' Use as the default daily-driver.
' ============================

Sub NormalizeTables_Light()
    Dim doc As Document
    Dim tbl As Table
    Dim row As row
    Dim cel As Cell
    Dim para As Paragraph

    Set doc = ActiveDocument

    For Each tbl In doc.Tables
        ' --- Table-level: full width, no fixed size lock ---
        With tbl
            .PreferredWidthType = 3     ' wdPreferPercent = 3
            .PreferredWidth = 100
            .AllowAutoFit = True
        End With

        ' --- Row-level: clear height constraints ---
        For Each row In tbl.Rows
            row.HeightRule = 0          ' wdRowHeightAuto = 0
            row.Height = 0
            row.AllowBreakAcrossPages = True
        Next row

        ' --- Cell-level: clear width constraints, apply Arial 10 ---
        For Each cel In tbl.Range.Cells
            cel.Width = 0
            cel.PreferredWidthType = 0  ' wdPreferAuto = 0

            For Each para In cel.Range.Paragraphs
                With para.Range.Font
                    .Name = "Arial"
                    .Size = 10
                End With
            Next para
        Next cel
    Next tbl

    MsgBox "NormalizeTables_Light complete. " & doc.Tables.Count & " table(s) processed.", vbInformation
End Sub
