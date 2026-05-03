Sub ResizeImagesAndCleanDocument()
    Dim i As Long
    Dim minWidth As Single
    Dim maxWidth As Single
    Dim docRange As Range
    Dim borderWidth As Single
    Dim borderColorR As Integer
    Dim borderColorG As Integer
    Dim borderColorB As Integer
    Dim inputValid As Boolean

    ' Validate Minimum Width
    Do
        minWidth = InchesToPoints(InputBox("Enter the minimum width (in inches):", "Minimum Width", 3))
        inputValid = minWidth > 0
        If Not inputValid Then MsgBox "Please enter a valid positive number for Minimum Width.", vbExclamation
    Loop Until inputValid

    ' Validate Maximum Width
    Do
        maxWidth = InchesToPoints(InputBox("Enter the maximum width (in inches):", "Maximum Width", 6.3))
        inputValid = maxWidth > minWidth
        If Not inputValid Then MsgBox "Maximum Width must be greater than Minimum Width.", vbExclamation
    Loop Until inputValid

    ' Validate Border Width
    Do
        borderWidth = InputBox("Enter the border width (in points):", "Border Width", 1.2)
        inputValid = borderWidth > 0
        If Not inputValid Then MsgBox "Please enter a valid positive number for Border Width.", vbExclamation
    Loop Until inputValid

    ' Validate Border Color - Red
    Do
        borderColorR = InputBox("Enter the red component of the border color (0-255):", "Border Color - Red", 68)
        inputValid = borderColorR >= 0 And borderColorR <= 255
        If Not inputValid Then MsgBox "Please enter a valid number between 0 and 255 for the Red component.", vbExclamation
    Loop Until inputValid

    ' Validate Border Color - Green
    Do
        borderColorG = InputBox("Enter the green component of the border color (0-255):", "Border Color - Green", 114)
        inputValid = borderColorG >= 0 And borderColorG <= 255
        If Not inputValid Then MsgBox "Please enter a valid number between 0 and 255 for the Green component.", vbExclamation
    Loop Until inputValid

    ' Validate Border Color - Blue
    Do
        borderColorB = InputBox("Enter the blue component of the border color (0-255):", "Border Color - Blue", 198)
        inputValid = borderColorB >= 0 And borderColorB <= 255
        If Not inputValid Then MsgBox "Please enter a valid number between 0 and 255 for the Blue component.", vbExclamation
    Loop Until inputValid

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