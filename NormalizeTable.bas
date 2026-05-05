Attribute VB_Name = "NormalizeTable"
Option Explicit

' =========================
' NormalizeTable.bas
' Light, fast Word table normalization macro.
' Use as the default daily-driver.
' For stubborn tables with images or nested content,
' use the full StandardizeTables_TwoPass_AllStories macro.
' =========================

Private Const TARGET_FONT_NAME As String = "Arial"
Private Const TARGET_FONT_SIZE As Single = 10

Public Sub NormalizeTables_Light()
    Dim doc As Document
    Dim srStory As Range, sr As Range, tbl As Table
    Dim startT As Double

    Set doc = ActiveDocument
    startT = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    Application.StatusBar = "Normalizing tables..."

    For Each srStory In doc.StoryRanges
        Set sr = srStory
        Do While Not sr Is Nothing
            If sr.Tables.Count > 0 Then
                For Each tbl In sr.Tables
                    ApplyNormalization tbl
                Next tbl
            End If
            Set sr = sr.NextStoryRange
        Loop
    Next srStory

    Application.ScreenUpdating = True
    Application.DisplayAlerts = wdAlertsAll
    Application.StatusBar = False

    MsgBox "Done in " & Format(Timer - startT, "0.00") & " sec.", vbInformation, "Normalize Tables"
End Sub

' =========================
' Core: apply normalization to one table
' =========================
Private Sub ApplyNormalization(ByVal tbl As Table)
    On Error GoTo SafeExit

    ' Layout
    tbl.Rows.WrapAroundText = False
    tbl.Rows.Alignment = wdAlignRowLeft
    tbl.Rows.LeftIndent = 0

    ' Clear row heights
    ClearRowHeights tbl

    ' Clear column preferred widths
    ClearColumnPreferredWidths tbl

    ' AutoFit to window/margins
    tbl.AllowAutoFit = True
    tbl.AutoFitBehavior wdAutoFitWindow

    ' Set table to 100% width and lock
    tbl.PreferredWidthType = wdPreferredWidthPercent
    tbl.PreferredWidth = 100
    tbl.AllowAutoFit = False

    ' Font and alignment
    tbl.Range.Font.Name = TARGET_FONT_NAME
    tbl.Range.Font.Size = TARGET_FONT_SIZE
    tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft

SafeExit:
    If Err.Number <> 0 Then
        Debug.Print "!! NormalizeTables_Light error: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
End Sub

' =========================
' Helpers
' =========================
Private Sub ClearRowHeights(ByVal tbl As Table)
    Dim rw As Row
    On Error Resume Next
    For Each rw In tbl.Rows
        rw.HeightRule = wdRowHeightAuto
        rw.Height = 0
    Next rw
    On Error GoTo 0
End Sub

Private Sub ClearColumnPreferredWidths(ByVal tbl As Table)
    Dim col As Column
    On Error Resume Next
    For Each col In tbl.Columns
        col.PreferredWidthType = wdPreferredWidthAuto
        col.PreferredWidth = 0
    Next col
    On Error GoTo 0
End Sub
