Option Explicit

Sub ResizeImagesAndCleanDocument()
    Dim i As Long
    Dim minWidth As Single
    Dim maxWidth As Single
    Dim borderWidth As Single
    Dim borderColorR As Integer
    Dim borderColorG As Integer
    Dim borderColorB As Integer
    Dim borderColorRGB As Long
    Dim s As String
    Dim ok As Boolean
    Dim oPara As Paragraph
    Dim sText As String
    Dim blankCount As Long
    Dim colorMode As String

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
        s = InputBox("Choose border color input method:" & vbCrLf & _
                     "  1 = RGB (three separate values)" & vbCrLf & _
                     "  2 = Hex code (e.g. #4472C6 or 4472C6)" & vbCrLf & vbCrLf & _
                     "Enter 1 or 2:", "Border Color Mode", "1")
        If s = "" Then Exit Sub
        s = Trim$(s)
        ok = (s = "1" Or s = "2")
        If Not ok Then
            MsgBox "Please enter 1 for RGB or 2 for Hex.", vbExclamation
        End If
    Loop Until ok
    colorMode = s

    If colorMode = "1" Then

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

        borderColorRGB = RGB(borderColorR, borderColorG, borderColorB)

    Else

        Do
            s = InputBox("Enter the border color as a hex code:" & vbCrLf & _
                         "(with or without #, e.g. #4472C6 or 4472C6)", _
                         "Border Color - Hex", "#4472C6")
            If s = "" Then Exit Sub
            borderColorRGB = HexToRGB(Trim$(s))
            ok = (borderColorRGB <> -1)
            If Not ok Then
                MsgBox "Please enter a valid 6-character hex color code.", vbExclamation
            End If
        Loop Until ok

    End If

    With ActiveDocument
        For i = 1 To .InlineShapes.Count
            With .InlineShapes(i)
                If .Type = wdInlineShapePicture Then
                    If .Width > minWidth Then
                        .LockAspectRatio = msoTrue
                        .Width = maxWidth
                        .Line.Weight = borderWidth
                        .Line.Style = msoLineSingle
                        .Line.ForeColor.RGB = borderColorRGB
                    End If
                End If
            End With
        Next i

        For i = .Paragraphs.Count To 1 Step -1
            Set oPara = .Paragraphs(i)
            sText = ParaTextWithoutMark(oPara)

            If IsEffectivelyBlank(sText) Then
                If oPara.Range.ListFormat.ListType <> wdListNoNumbering Then
                    oPara.Range.ListFormat.RemoveNumbers
                    oPara.Range.Style = .Styles(wdStyleNormal)
                End If
            End If
        Next i

        blankCount = 0
        For i = .Paragraphs.Count To 1 Step -1
            Set oPara = .Paragraphs(i)
            sText = ParaTextWithoutMark(oPara)

            If IsEffectivelyBlank(sText) Then
                blankCount = blankCount + 1

                If blankCount > 1 Then
                    If Not oPara.Range.Information(wdWithInTable) Then
                        On Error Resume Next
                        oPara.Range.Delete
                        On Error GoTo 0
                    Else
                        oPara.Range.ListFormat.RemoveNumbers
                        oPara.Range.Style = .Styles(wdStyleNormal)
                    End If
                Else
                    oPara.Range.ListFormat.RemoveNumbers
                    oPara.Range.Style = .Styles(wdStyleNormal)
                End If
            Else
                blankCount = 0
            End If
        Next i

        MsgBox "Images resized and document cleaned.", vbInformation
    End With
End Sub

Private Function HexToRGB(sHex As String) As Long
    Dim r As Integer, g As Integer, b As Integer
    
    sHex = Replace(Trim$(sHex), "#", "")
    
    If Len(sHex) <> 6 Then
        HexToRGB = -1
        Exit Function
    End If
    
    On Error GoTo InvalidHex
    
    r = CInt("&H" & Left$(sHex, 2))
    g = CInt("&H" & Mid$(sHex, 3, 2))
    b = CInt("&H" & Right$(sHex, 2))
    
    HexToRGB = RGB(r, g, b)
    Exit Function

InvalidHex:
    HexToRGB = -1
End Function

Private Function ParaTextWithoutMark(oPara As Paragraph) As String
    Dim sText As String
    
    sText = oPara.Range.Text
    If Len(sText) > 0 Then
        If Asc(Right$(sText, 1)) = 13 Then
            sText = Left$(sText, Len(sText) - 1)
        End If
    End If
    
    ParaTextWithoutMark = sText
End Function

Private Function IsEffectivelyBlank(sLine As String) As Boolean
    Dim sClean As String
    
    sClean = sLine
    sClean = Replace(sClean, Chr(9), "")
    sClean = Replace(sClean, Chr(10), "")
    sClean = Replace(sClean, Chr(11), "")
    sClean = Replace(sClean, Chr(12), "")
    sClean = Replace(sClean, Chr(13), "")
    sClean = Replace(sClean, Chr(32), "")
    sClean = Replace(sClean, Chr(160), "")
    sClean = Replace(sClean, ChrW(8203), "")
    sClean = Replace(sClean, ChrW(8204), "")
    sClean = Replace(sClean, ChrW(8205), "")
    sClean = Replace(sClean, ChrW(8206), "")
    sClean = Replace(sClean, ChrW(8207), "")
    sClean = Replace(sClean, ChrW(8239), "")
    sClean = Replace(sClean, ChrW(8287), "")
    sClean = Replace(sClean, ChrW(173), "")
    sClean = Replace(sClean, ChrW(65279), "")
    
    IsEffectivelyBlank = (Len(sClean) = 0)
End Function

