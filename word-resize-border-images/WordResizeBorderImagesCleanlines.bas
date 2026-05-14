Sub ResizeImagesAndCleanDocument()
    Dim i As Long
    Dim minWidth As Single
    Dim maxWidth As Single
    Dim docRange As Range
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
        s = Trim(s)
        ok = (s = "1" Or s = "2")
        If Not ok Then MsgBox "Please enter 1 for RGB or 2 for Hex.", vbExclamation
    Loop Until ok
    colorMode = s

    If colorMode = "1" Then
        Do
            s = InputBox("Enter the red component (0-255):", "Border Color - Red", 68)
            If s = "" Then Exit Sub
            ok = IsNumeric(s) And CLng(s) >= 0 And CLng(s) <= 255
            If ok Then borderColorR = CLng(s) Else MsgBox "Enter 0-255 for Red.", vbExclamation
        Loop Until ok
        Do
            s = InputBox("Enter the green component (0-255):", "Border Color - Green", 114)
            If s = "" Then Exit Sub
            ok = IsNumeric(s) And CLng(s) >= 0 And CLng(s) <= 255
            If ok Then borderColorG = CLng(s) Else MsgBox "Enter 0-255 for Green.", vbExclamation
        Loop Until ok
        Do
            s = InputBox("Enter the blue component (0-255):", "Border Color - Blue", 198)
            If s = "" Then Exit Sub
            ok = IsNumeric(s) And CLng(s) >= 0 And CLng(s) <= 255
            If ok Then borderColorB = CLng(s) Else MsgBox "Enter 0-255 for Blue.", vbExclamation
        Loop Until ok
        borderColorRGB = RGB(borderColorR, borderColorG, borderColorB)
    Else
        Do
            s = InputBox("Enter hex color (e.g. #4472C6 or 4472C6):", "Border Color - Hex", "#4472C6")
            If s = "" Then Exit Sub
            borderColorRGB = HexToRGB(Trim(s))
            ok = (borderColorRGB <> -1)
            If Not ok Then MsgBox "Enter a valid 6-character hex code.", vbExclamation
        Loop Until ok
    End If

    With ActiveDocument

        ' FEATURE 1: Resize images and apply border
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

        ' FEATURE 2 - STEP 0: Remove ghost bullet/numbered blank paragraphs
        blankCount = 0
        For i = .Paragraphs.Count To 1 Step -1
            Set oPara = .Paragraphs(i)
            sText = oPara.Range.Text
            If Len(sText) > 0 Then
                If Asc(Right(sText, 1)) = 13 Then sText = Left(sText, Len(sText) - 1)
            End If
            If IsEffectivelyBlank(sText) Then
                If oPara.Range.ListFormat.ListType <> wdListNoNumbering Then
                    oPara.Range.Delete
                End If
            End If
        Next i

        ' FEATURE 2 - STEP 1: Collapse consecutive blank paragraphs to max 1
        blankCount = 0
        For i = .Paragraphs.Count To 1 Step -1
            Set oPara = .Paragraphs(i)
            sText = oPara.Range.Text
            If Len(sText) > 0 Then
                If Asc(Right(sText, 1)) = 13 Then sText = Left(sText, Len(sText) - 1)
            End If
            If IsEffectivelyBlank(sText) Then
                blankCount = blankCount + 1
                If blankCount > 1 Then
                    oPara.Range.Delete
                Else
                    With oPara.Range
                        .ListFormat.RemoveNumbers
                        .Style = ActiveDocument.Styles(wdStyleNormal)
                    End With
                End If
            Else
                blankCount = 0
            End If
        Next i

        MsgBox "Images resized and document cleaned.", vbInformation
    End With
End Sub

Private Function HexToRGB(sHex As String) As Long
    sHex = Replace(sHex, "#", "")
    If Len(sHex) <> 6 Then HexToRGB = -1 : Exit Function
    On Error GoTo InvalidHex
    Dim r As Integer, g As Integer, b As Integer
    r = CInt("&H" & Left(sHex, 2))
    g = CInt("&H" & Mid(sHex, 3, 2))
    b = CInt("&H" & Right(sHex, 2))
    HexToRGB = RGB(r, g, b)
    Exit Function
InvalidHex:
    HexToRGB = -1
End Function

Private Function IsEffectivelyBlank(sIn As String) As Boolean
    Dim sClean As String
    sClean = sIn
    sClean = Trim(sClean)
    sClean = Replace(sClean, Chr(160),    "")  ' Non-breaking space
    sClean = Replace(sClean, ChrW(8203),  "")  ' Zero-width space
    sClean = Replace(sClean, ChrW(8204),  "")  ' Zero-width non-joiner
    sClean = Replace(sClean, ChrW(8205),  "")  ' Zero-width joiner
    sClean = Replace(sClean, ChrW(8206),  "")  ' Left-to-right mark
    sClean = Replace(sClean, ChrW(8207),  "")  ' Right-to-left mark
    sClean = Replace(sClean, ChrW(8239),  "")  ' Narrow no-break space
    sClean = Replace(sClean, ChrW(8287),  "")  ' Medium mathematical space
    sClean = Replace(sClean, ChrW(173),   "")  ' Soft hyphen
    sClean = Replace(sClean, ChrW(65279), "")  ' BOM / Zero-width no-break space
    IsEffectivelyBlank = (Len(sClean) = 0)
End Function
