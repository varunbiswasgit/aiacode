Attribute VB_Name = "NormalizeTable"
Option Explicit

' ============================
' NormalizeTable.bas
' Light, fast Word table normalization macro.
' Use as the default daily-driver.
' For stubborn tables with images or nested content,
' use the full StandardizeTables_TwoPass_AllStories macro.
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
            .PreferredWidthType = wdPreferPercent
            .PreferredWidth = 100
            .AllowAutoFit = True
        End With

        ' --- Row-level: clear height constraints ---
        For Each row In tbl.Rows
            row.HeightRule = wdRowHeightAuto
            row.Height = 0
            row.AllowBreakAcrossPages = True
        Next row

        ' --- Cell-level: clear width constraints, apply Arial 10 ---
        For Each cel In tbl.Range.Cells
            cel.Width = 0
            cel.PreferredWidthType = wdPreferAuto

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

' ============================
' StandardizeTables_TwoPass_AllStories
' Heavier two-pass macro for documents with images or nested tables.
' Iterates all stories (body, headers, footers, text boxes).
' ============================

Sub StandardizeTables_TwoPass_AllStories()
    Dim doc As Document
    Dim stry As Range
    Dim tbl As Table
    Dim row As row
    Dim cel As Cell
    Dim para As Paragraph
    Dim tableCount As Long

    Set doc = ActiveDocument
    tableCount = 0

    For Each stry In doc.StoryRanges
        Do
            For Each tbl In stry.Tables
                ' Pass 1: unlock dimensions
                tbl.PreferredWidthType = wdPreferPercent
                tbl.PreferredWidth = 100
                tbl.AllowAutoFit = True

                For Each row In tbl.Rows
                    row.HeightRule = wdRowHeightAuto
                    row.Height = 0
                    row.AllowBreakAcrossPages = True
                Next row

                For Each cel In tbl.Range.Cells
                    cel.Width = 0
                    cel.PreferredWidthType = wdPreferAuto
                Next cel

                ' Pass 2: apply font
                For Each cel In tbl.Range.Cells
                    For Each para In cel.Range.Paragraphs
                        With para.Range.Font
                            .Name = "Arial"
                            .Size = 10
                        End With
                    Next para
                Next cel

                tableCount = tableCount + 1
            Next tbl

            Set stry = stry.NextStoryRange
        Loop While Not stry Is Nothing
    Next stry

    MsgBox "StandardizeTables_TwoPass_AllStories complete. " & tableCount & " table(s) processed.", vbInformation
End Sub
