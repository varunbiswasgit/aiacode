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
    Dim oPara As Paragraph
    Dim sText As String
    Dim blankCount As Long

    ' -------------------------------------------------------------------------
    ' INPUT: Minimum image width
    ' -------------------------------------------------------------------------
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

    ' -------------------------------------------------------------------------
    ' INPUT: Maximum image width
    ' -------------------------------------------------------------------------
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

    ' -------------------------------------------------------------------------
    ' INPUT: Border width
    ' -------------------------------------------------------------------------
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

    ' -------------------------------------------------------------------------
    ' INPUT: Border color - Red
    ' -------------------------------------------------------------------------
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

    ' -------------------------------------------------------------------------
    ' INPUT: Border color - Green
    ' -------------------------------------------------------------------------
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

    ' -------------------------------------------------------------------------
    ' INPUT: Border color - Blue
    ' -------------------------------------------------------------------------
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

        ' =====================================================================
        ' FEATURE 1: Resize images wider than minWidth and apply border
        ' UNCHANGED from original
        ' =====================================================================
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

        ' =====================================================================
        ' FEATURE 2 - STEP 0: Remove pre-existing ghost bullet / numbered items
        ' A ghost is a paragraph that is blank (including Chr(160), Unicode
        ' zero-width chars, BOM, etc.) AND still carries list formatting.
        ' These are deleted outright as they have no content value.
        ' Loops backwards to avoid index shifting on deletion.
        ' =====================================================================
        blankCount = 0
        For i = .Paragraphs.Count To 1 Step -1
            Set oPara = .Paragraphs(i)
            sText = oPara.Range.Text
            ' Strip trailing paragraph mark Chr(13) before blank test
            If Len(sText) > 0 Then
                If Asc(Right(sText, 1)) = 13 Then
                    sText = Left(sText, Len(sText) - 1)
                End If
            End If
            If IsEffectivelyBlank(sText) Then
                If oPara.Range.ListFormat.ListType <> wdListNoNumbering Then
                    oPara.Range.Delete
                End If
            End If
        Next i

        ' =====================================================================
        ' FEATURE 2 - STEP 1: Collapse consecutive blank paragraphs to max 1.
        ' Handles standard Chr(13) blanks AND impostor blanks containing
        ' Chr(160), Unicode zero-width spaces, BOM, and other SAP export chars
        ' that Trim() and the ^13{3,} wildcard engine do not catch.
        ' On the first blank in a run, list formatting is stripped to prevent
        ' ghost bullet lines being created during the collapse.
        ' Extra consecutive blanks are deleted outright.
        ' Loops backwards to avoid index shifting on deletion.
        ' =====================================================================
        blankCount = 0
        For i = .Paragraphs.Count To 1 Step -1
            Set oPara = .Paragraphs(i)
            sText = oPara.Range.Text
            ' Strip trailing paragraph mark Chr(13) before blank test
            If Len(sText) > 0 Then
                If Asc(Right(sText, 1)) = 13 Then
                    sText = Left(sText, Len(sText) - 1)
                End If
            End If
            If IsEffectivelyBlank(sText) Then
                blankCount = blankCount + 1
                If blankCount > 1 Then
                    ' Extra consecutive blank - delete it
                    oPara.Range.Delete
                Else
                    ' First blank in a run - keep it but strip any list/bullet
                    ' formatting to prevent ghost bullet creation
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

' =============================================================================
' IsEffectivelyBlank
' Returns True if a string contains only whitespace-equivalent characters,
' including SAP/Unicode invisible characters that VBA Trim() does not catch:
'   Chr(160)     Non-breaking space     - common in SAP ALV/GUI exports
'   ChrW(8203)   Zero-width space
'   ChrW(8204)   Zero-width non-joiner
'   ChrW(8205)   Zero-width joiner
'   ChrW(8206)   Left-to-right mark
'   ChrW(8207)   Right-to-left mark
'   ChrW(8239)   Narrow no-break space  - SAP EU locale number formatting
'   ChrW(8287)   Medium mathematical space
'   ChrW(173)    Soft hyphen
'   ChrW(65279)  BOM / Zero-width no-break space - SAP CSV/Excel export marker
' =============================================================================
Private Function IsEffectivelyBlank(sLine As String) As Boolean
    Dim sClean As String
    sClean = sLine

    ' Standard ASCII whitespace
    sClean = Replace(sClean, Chr(9),   "")   ' Tab
    sClean = Replace(sClean, Chr(10),  "")   ' Line Feed
    sClean = Replace(sClean, Chr(11),  "")   ' Vertical tab (Word line break)
    sClean = Replace(sClean, Chr(12),  "")   ' Form feed
    sClean = Replace(sClean, Chr(13),  "")   ' Carriage Return
    sClean = Replace(sClean, Chr(32),  "")   ' Space
    sClean = Replace(sClean, Chr(160), "")   ' Non-breaking space (NBSP)

    ' SAP / Unicode invisible characters
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
