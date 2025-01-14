Sub ResizeImagesAndCleanDocument()
    Dim i As Long
    Dim minWidth As Single
    Dim maxWidth As Single
    Dim docRange As Range
    Dim borderWidth As Single
    Dim borderColorR As Integer
    Dim borderColorG As Integer
    Dim borderColorB As Integer

    ' Prompt user for inputs
    minWidth = InchesToPoints(InputBox("Enter the minimum width (in inches):", "Minimum Width", 3))
    maxWidth = InchesToPoints(InputBox("Enter the maximum width (in inches):", "Maximum Width", 6.3))
    borderWidth = InputBox("Enter the border width (in points):", "Border Width", 1.2)
    borderColorR = InputBox("Enter the red component of the border color (0-255):", "Border Color - Red", 68)
    borderColorG = InputBox("Enter the green component of the border color (0-255):", "Border Color - Green", 114)
    borderColorB = InputBox("Enter the blue component of the border color (0-255):", "Border Color - Blue", 198)

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
                        .Line.Weight = borderWidth
                        .Line.Style = msoLineSingle
                        .Line.ForeColor.RGB = RGB(borderColorR, borderColorG, borderColorB)
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
