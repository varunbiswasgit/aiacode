Attribute VB_Name = "BoldListPrefixesOutlook"
Option Explicit

' =============================================================================
' BoldListPrefixesOutlook
' Version : v1
' Author  : Varun Biswas (AI-assisted)
' Purpose : Bolds the prefix of every bulleted/numbered list item up to and
'           including the first colon (:) or dash (-), whichever appears first.
'           Works in both Word and the Outlook message editor.
' =============================================================================

Sub BoldListPrefixesOutlook()
    Dim doc     As Object
    Dim para    As Object
    Dim txt     As String
    Dim posColon As Long
    Dim posDash  As Long
    Dim endPos   As Long
    Dim rng     As Object

    ' --- Resolve active document (Word or Outlook inspector) -----------------
    On Error Resume Next
    Set doc = Application.ActiveDocument
    If doc Is Nothing Then
        Set doc = Application.ActiveInspector.WordEditor
    End If
    On Error GoTo 0

    If doc Is Nothing Then
        MsgBox "No editable document found.", vbExclamation, "BoldListPrefixesOutlook"
        Exit Sub
    End If

    ' --- Iterate every paragraph in the document ----------------------------
    For Each para In doc.Paragraphs

        ' Only process paragraphs that belong to a list (bulleted or numbered)
        If para.Range.ListFormat.ListType <> 0 Then

            txt = para.Range.Text

            posColon = InStr(txt, ":")
            posDash  = InStr(txt, "-")

            ' Pick the earliest delimiter present
            If posColon > 0 And posDash > 0 Then
                endPos = IIf(posColon < posDash, posColon, posDash)
            ElseIf posColon > 0 Then
                endPos = posColon
            ElseIf posDash > 0 Then
                endPos = posDash
            Else
                endPos = 0
            End If

            ' Bold from the start of the paragraph text up to and including
            ' the delimiter (endPos is 1-based within txt; rng.Start is absolute)
            If endPos > 1 Then
                Set rng = para.Range.Duplicate
                rng.End = rng.Start + endPos - 1
                rng.Font.Bold = True
            End If

        End If

    Next para

End Sub
