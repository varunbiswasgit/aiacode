Sub ResizeImagesAndCleanDocument()
    Dim i As Long
    Dim minWidth As Single
    Dim maxWidth As Single
    Dim docRange As Range
    
    minWidth = InchesToPoints(3) ' Minimum width (inches in points)
    maxWidth = InchesToPoints(6.3) ' Maximum width (inches in points)

    With ActiveDocument
        ' Loop through all inline shapes in the document
        For i = 1 To .InlineShapes.Count
            With .InlineShapes(i)
                If .Type = wdInlineShapePicture Then
                    ' Resize only if the width exceeds the minimum width
                    If .Width > minWidth Then
                        .LockAspectRatio = msoTrue
                        .Width = maxWidth
						' Apply border to all images
						.Line.Weight = 1.2
						.Line.Style = msoLineSingle
						.Line.ForeColor.RGB = RGB(68, 114, 198)
                    End If
                End If
            End With
        Next i
        
        ' Remove multiple empty lines (more than 2 consecutive)
        Set docRange = .Content
        With docRange.Find
            .ClearFormatting
            .Text = "^13{3,}" ' Finds three or more consecutive empty lines
            .Replacement.Text = "^p^p" ' Replaces with two empty lines
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
    End With
End Sub

