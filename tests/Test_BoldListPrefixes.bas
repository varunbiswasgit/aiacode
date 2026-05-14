Attribute VB_Name = "Test_BoldListPrefixes"
Option Explicit

' =============================================================================
' Test_BoldListPrefixes.bas
' Automated VBA unit tests for BoldListPrefixes.
' Run from the same VBA project that contains BoldListPrefixes.bas.
'
' Execute:  Alt+F8 -> RunAllTests
' Output:   Immediate window (Ctrl+G) and a summary MsgBox.
' =============================================================================

' ---------------------------------------------------------------------------
' Public entry point
' ---------------------------------------------------------------------------
Public Sub RunAllTests()
    Dim passed As Long
    Dim failed As Long
    Dim log    As String

    passed = 0 : failed = 0 : log = ""

    Call RunTest_ColonDelimiter(passed, failed, log)
    Call RunTest_DashDelimiter(passed, failed, log)
    Call RunTest_ColonBeforeDash(passed, failed, log)
    Call RunTest_DashBeforeColon(passed, failed, log)
    Call RunTest_NoDelimiter(passed, failed, log)
    Call RunTest_DelimiterAtPosition1(passed, failed, log)
    Call RunTest_NonListParagraphUnchanged(passed, failed, log)

    Dim summary As String
    summary = "BoldListPrefixes Test Run" & vbCrLf & _
              "Passed : " & passed & vbCrLf & _
              "Failed : " & failed & vbCrLf & vbCrLf & log

    Debug.Print summary
    MsgBox summary, IIf(failed = 0, vbInformation, vbExclamation), "Test Results"
End Sub

' ---------------------------------------------------------------------------
' Helper: create a throw-away Word document pre-populated with test content.
' The caller is responsible for closing the document.
' ---------------------------------------------------------------------------
Private Function CreateTestDoc(listItems() As String, _
                               Optional nonListLine As String = "") As Document
    Dim doc As Document
    Dim para As Paragraph
    Dim i   As Long

    Set doc = Documents.Add
    doc.Content.Text = ""  ' clear default empty paragraph

    ' Insert an optional non-list paragraph first
    If Len(nonListLine) > 0 Then
        With doc.Content
            .InsertAfter nonListLine & vbCr
        End With
    End If

    ' Insert bulleted list items
    For i = LBound(listItems) To UBound(listItems)
        Dim rng As Range
        Set rng = doc.Content
        rng.Collapse wdCollapseEnd
        rng.InsertAfter listItems(i) & vbCr
        rng.ListFormat.ApplyListTemplateWithLevel _
            ListTemplate:=ListGalleries(wdBulletGallery).ListTemplates(1), _
            ContinuePreviousList:=False, _
            ApplyTo:=wdListApplyToWholeList, _
            DefaultListBehavior:=wdWord10ListBehavior
    Next i

    Set CreateTestDoc = doc
End Function

' ---------------------------------------------------------------------------
' Helper: check whether text in a paragraph range is bold up to a position.
' posExpectedBoldEnd is 1-based (inclusive) within the paragraph text.
' ---------------------------------------------------------------------------
Private Function IsBoldUpTo(para As Paragraph, posExpectedBoldEnd As Long) As Boolean
    Dim rng As Range
    Set rng = para.Range.Duplicate
    ' Check the character AT posExpectedBoldEnd (the delimiter itself)
    rng.Start = para.Range.Start
    rng.End   = para.Range.Start + posExpectedBoldEnd - 1
    IsBoldUpTo = (rng.Font.Bold = True)
End Function

Private Function IsNotBoldAfter(para As Paragraph, posAfter As Long) As Boolean
    Dim rng As Range
    Set rng = para.Range.Duplicate
    rng.Start = para.Range.Start + posAfter
    ' Move one character past delimiter to check the remainder
    If rng.Start >= para.Range.End - 1 Then
        IsNotBoldAfter = True   ' nothing after delimiter
        Exit Function
    End If
    rng.End = para.Range.End - 1  ' exclude trailing CR
    IsNotBoldAfter = (rng.Font.Bold = False)
End Function

' ---------------------------------------------------------------------------
' TC-01  Colon delimiter — bulleted list
' ---------------------------------------------------------------------------
Private Sub RunTest_ColonDelimiter(passed As Long, failed As Long, log As String)
    Const TEST_NAME As String = "TC-01 Colon delimiter"
    Dim items(0) As String
    items(0) = "Scope: defines the boundary"

    Dim doc As Document
    Set doc = CreateTestDoc(items)

    BoldListPrefixes   ' run the macro under test

    Dim para As Paragraph
    Set para = doc.Paragraphs(1)  ' adjust if non-list prefix paragraph exists

    ' Find the first list paragraph
    Dim lp As Paragraph
    For Each lp In doc.Paragraphs
        If lp.Range.ListFormat.ListType <> 0 Then
            Set para = lp : Exit For
        End If
    Next lp

    Dim posColon As Long
    posColon = InStr(para.Range.Text, ":")

    If IsBoldUpTo(para, posColon) And IsNotBoldAfter(para, posColon) Then
        passed = passed + 1
        log = log & "PASS  " & TEST_NAME & vbCrLf
    Else
        failed = failed + 1
        log = log & "FAIL  " & TEST_NAME & vbCrLf
    End If

    doc.Close wdDoNotSaveChanges
End Sub

' ---------------------------------------------------------------------------
' TC-02  Dash delimiter — numbered list
' ---------------------------------------------------------------------------
Private Sub RunTest_DashDelimiter(passed As Long, failed As Long, log As String)
    Const TEST_NAME As String = "TC-02 Dash delimiter"
    Dim items(0) As String
    items(0) = "Owner - accountable"

    Dim doc As Document
    Set doc = CreateTestDoc(items)
    BoldListPrefixes

    Dim lp As Paragraph
    For Each lp In doc.Paragraphs
        If lp.Range.ListFormat.ListType <> 0 Then Exit For
    Next lp

    Dim posDash As Long
    posDash = InStr(lp.Range.Text, "-")

    If IsBoldUpTo(lp, posDash) And IsNotBoldAfter(lp, posDash) Then
        passed = passed + 1
        log = log & "PASS  " & TEST_NAME & vbCrLf
    Else
        failed = failed + 1
        log = log & "FAIL  " & TEST_NAME & vbCrLf
    End If

    doc.Close wdDoNotSaveChanges
End Sub

' ---------------------------------------------------------------------------
' TC-03  Colon appears before dash — colon should win
' ---------------------------------------------------------------------------
Private Sub RunTest_ColonBeforeDash(passed As Long, failed As Long, log As String)
    Const TEST_NAME As String = "TC-03 Colon before dash"
    Dim items(0) As String
    items(0) = "Priority: high - urgent"

    Dim doc As Document
    Set doc = CreateTestDoc(items)
    BoldListPrefixes

    Dim lp As Paragraph
    For Each lp In doc.Paragraphs
        If lp.Range.ListFormat.ListType <> 0 Then Exit For
    Next lp

    Dim txt As String
    txt = lp.Range.Text
    Dim posColon As Long : posColon = InStr(txt, ":")

    If IsBoldUpTo(lp, posColon) And IsNotBoldAfter(lp, posColon) Then
        passed = passed + 1
        log = log & "PASS  " & TEST_NAME & vbCrLf
    Else
        failed = failed + 1
        log = log & "FAIL  " & TEST_NAME & vbCrLf
    End If

    doc.Close wdDoNotSaveChanges
End Sub

' ---------------------------------------------------------------------------
' TC-04  Dash appears before colon — dash should win
' ---------------------------------------------------------------------------
Private Sub RunTest_DashBeforeColon(passed As Long, failed As Long, log As String)
    Const TEST_NAME As String = "TC-04 Dash before colon"
    Dim items(0) As String
    items(0) = "Step-by-step: follow"

    Dim doc As Document
    Set doc = CreateTestDoc(items)
    BoldListPrefixes

    Dim lp As Paragraph
    For Each lp In doc.Paragraphs
        If lp.Range.ListFormat.ListType <> 0 Then Exit For
    Next lp

    Dim txt As String
    txt = lp.Range.Text
    Dim posDash As Long : posDash = InStr(txt, "-")

    If IsBoldUpTo(lp, posDash) And IsNotBoldAfter(lp, posDash) Then
        passed = passed + 1
        log = log & "PASS  " & TEST_NAME & vbCrLf
    Else
        failed = failed + 1
        log = log & "FAIL  " & TEST_NAME & vbCrLf
    End If

    doc.Close wdDoNotSaveChanges
End Sub

' ---------------------------------------------------------------------------
' TC-05  No delimiter — paragraph must be skipped (no bold applied)
' ---------------------------------------------------------------------------
Private Sub RunTest_NoDelimiter(passed As Long, failed As Long, log As String)
    Const TEST_NAME As String = "TC-05 No delimiter"
    Dim items(0) As String
    items(0) = "Plain list item no delimiter here"

    Dim doc As Document
    Set doc = CreateTestDoc(items)
    BoldListPrefixes

    Dim lp As Paragraph
    For Each lp In doc.Paragraphs
        If lp.Range.ListFormat.ListType <> 0 Then Exit For
    Next lp

    Dim rng As Range
    Set rng = lp.Range.Duplicate
    rng.End = rng.End - 1   ' exclude trailing CR

    If rng.Font.Bold = False Then
        passed = passed + 1
        log = log & "PASS  " & TEST_NAME & vbCrLf
    Else
        failed = failed + 1
        log = log & "FAIL  " & TEST_NAME & vbCrLf
    End If

    doc.Close wdDoNotSaveChanges
End Sub

' ---------------------------------------------------------------------------
' TC-06  Delimiter at position 1 — endPos > 1 guard, no bold applied
' ---------------------------------------------------------------------------
Private Sub RunTest_DelimiterAtPosition1(passed As Long, failed As Long, log As String)
    Const TEST_NAME As String = "TC-10 Delimiter at position 1"
    Dim items(0) As String
    items(0) = ": orphan colon"

    Dim doc As Document
    Set doc = CreateTestDoc(items)
    BoldListPrefixes

    Dim lp As Paragraph
    For Each lp In doc.Paragraphs
        If lp.Range.ListFormat.ListType <> 0 Then Exit For
    Next lp

    Dim rng As Range
    Set rng = lp.Range.Duplicate
    rng.End = rng.End - 1

    If rng.Font.Bold = False Then
        passed = passed + 1
        log = log & "PASS  " & TEST_NAME & vbCrLf
    Else
        failed = failed + 1
        log = log & "FAIL  " & TEST_NAME & vbCrLf
    End If

    doc.Close wdDoNotSaveChanges
End Sub

' ---------------------------------------------------------------------------
' TC-07  Non-list paragraph must remain untouched
' ---------------------------------------------------------------------------
Private Sub RunTest_NonListParagraphUnchanged(passed As Long, failed As Long, log As String)
    Const TEST_NAME As String = "TC-06 Non-list paragraph unchanged"
    Dim items(0) As String
    items(0) = "List item: relevant"   ' list item — will be bolded

    Dim doc As Document
    ' nonListLine inserted as first paragraph
    Set doc = CreateTestDoc(items, "Introduction: not a list item")
    BoldListPrefixes

    ' First paragraph should be a non-list paragraph — verify not bolded
    Dim firstPara As Paragraph
    Set firstPara = doc.Paragraphs(1)

    If firstPara.Range.ListFormat.ListType = 0 Then
        Dim rng As Range
        Set rng = firstPara.Range.Duplicate
        rng.End = rng.End - 1
        If rng.Font.Bold = False Then
            passed = passed + 1
            log = log & "PASS  " & TEST_NAME & vbCrLf
        Else
            failed = failed + 1
            log = log & "FAIL  " & TEST_NAME & " (non-list para was bolded)" & vbCrLf
        End If
    Else
        failed = failed + 1
        log = log & "FAIL  " & TEST_NAME & " (first para unexpectedly is a list item)" & vbCrLf
    End If

    doc.Close wdDoNotSaveChanges
End Sub
