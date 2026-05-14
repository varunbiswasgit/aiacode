Private Sub FormatTable(tbl As Table)

    Dim cl As Cell
    Dim rw As Row

    '--- Normalize structure FIRST (this is what UI does silently) ---
    On Error Resume Next
    tbl.AllowAutoFit = True
    tbl.AutoFitBehavior wdAutoFitContent
    On Error GoTo 0

    '--- Table ---
    With tbl
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        .Rows.Alignment = wdAlignRowLeft
    End With

    '--- Rows ---
    For Each rw In tbl.Rows
        rw.HeightRule = wdRowHeightAuto
    Next rw

    '--- Cells ONLY (avoid Columns collection entirely) ---
    For Each cl In tbl.Range.Cells
        On Error Resume Next
        cl.PreferredWidthType = wdPreferredWidthAuto
        cl.VerticalAlignment = wdCellAlignVerticalCenter
        On Error GoTo 0
    Next cl

End Sub
