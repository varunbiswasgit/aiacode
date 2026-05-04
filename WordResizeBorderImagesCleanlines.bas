Sub ResizeImagesAndCleanDocument()
    Dim i As Long
    Dim minWidth As Single
    Dim maxWidth As Single
    Dim docRange As Range
    Dim borderWidth As Single
    Dim borderColorR As Integer
    Dim borderColorG As Integer
    Dim borderColorB As Integer
    Dim s As String
    Dim ok As Boolean

    Do
        s = InputBox("Enter the minimum width (in inches):", "Minimum Width", 3)
        If s = "" Then Exit Sub
        ok = IsNumeric(s) And CDbl(s) > 0
        If ok Then
            minWidth = InchesToPoints(CDbl(s))
        Else
            MsgBox "Please enter a valid positive number for Minimum Width.", vbExclamation
        End If
    Loop Until ok

    Do
        s = InputBox("Enter the maximum width (in inches):", "Maximum Width", 6.3)
        If s = "" Then Exit Sub
        ok = IsNumeric(s) And CDbl(s) > 0 And InchesToPoints(CDbl(s)) > minWidth
        If ok Then
            maxWidth = InchesToPoints(CDbl(s))
        Else
            MsgBox "Maximum Width must be a valid number greater than Minimum Width.", vbExclamation
        End If
    Loop Until ok

    Do
        s = InputBox("Enter the border width (in points):", "Border Width", 1.2)
        If s = "" Then Exit Sub
        ok = IsNumeric(s) And CDbl(s) > 0
        If ok Then
            borderWidth = CSng(s)
        Else
            MsgBox "Please enter a valid positive number for Border Width.", vbExclamation
        End If
    Loop Until ok

    Do
        s = InputBox("Enter the red component of the border color (0-255):", "Border Color - Red", 68)
        If s = "" Then Exit Sub
        ok = IsNumeric(s) And CLng(s) >= 0 And CLng(s) <= 255
        If ok Then
            borderColorR = CLng(s)
        Else
            MsgBox "Please enter a valid number between 0 and 255 for the Red component.", vbExclamation
        End If
    Loop Until ok

    Do
        s = InputBox("Enter the green component of the border color (0-255):", "Border Color - Green", 114)
        If s = "" Then Exit Sub
        ok = IsNumeric(s) And CLng(s) >= 0 And CLng(s) <= 255
        If ok Then
            borderColorG = CLng(s)
        Else
            MsgBox "Please enter a valid number between 0 and 255 for the Green component.", vbExclamation
        End If
    Loop Until ok

    Do
        s = InputBox("Enter the blue component of the border color (0-255):", "Border Color - Blue", 198)
        If s = "" Then Exit Sub
        ok = IsNumeric(s) And CLng(s) >= 0 And CLng(s) <= 255
        If ok Then
            borderColorB = CLng(s)
        Else
            MsgBox "Please enter a valid number between 0 and 255 for the Blue component.", vbExclamation
        End If
    Loop Until ok

    With ActiveDocument
        For i = 1 To .InlineShapes.Count
            With .InlineShapes(i)
                If .Type = wdInlineShapePicture Then
                    If .Width > minWidth Then
                        .LockAspectRatio = msoTrue
                        .Width = maxWidth
                        .Line.Weight = borderWidth
                        .Line.Style = msoLineSingle
                        .Line.ForeColor.RGB = RGB(borderColorR, borderColorG, borderColorB)
                    End If
                End If
            End With
        Next i

        Set docRange = .Content
        With docRange.Find
            .ClearFormatting
            .Text = "^13{3,}"
            .Replacement.Text = "^p^p"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With

        MsgBox "Images resized and document cleaned.", vbInformation
    End With
End Sub
