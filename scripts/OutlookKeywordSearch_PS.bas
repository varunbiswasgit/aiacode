Attribute VB_Name = "OutlookKeywordSearch_PS"
Option Explicit

' ============================================================
' Outlook Keyword Search — PowerShell-Assisted VBA Launcher
' Author  : Varun Biswas
' Repo    : varunbiswasgit/aiacode
' Purpose : Thin VBA launcher that collects user inputs and
'           delegates all search work to OutlookKeywordSearch_PS.ps1
'           running as a separate PowerShell process.
'           Outlook UI remains fully responsive during search.
' Requires: OutlookKeywordSearch_PS.ps1 saved locally.
'           The macro will prompt for the script path at runtime.
' Change log:
'   2026-05-12  Retry loops added to all file/path prompts.
'               Macro exits only when user presses Cancel.
' ============================================================

Private Const PS_SCRIPT_DEFAULT As String = "C:\Users\Varun\scripts\OutlookKeywordSearch_PS.ps1"

Public Sub RunKeywordSearch()
    Dim psScriptPath As String
    Dim modeChoice   As String
    Dim keyword      As String
    Dim filePath     As String
    Dim colRef       As String
    Dim psCmd        As String
    Dim userInput    As String

    ' -------------------------------------------------------
    ' Step 1: PS script path — retry until valid or Cancelled.
    ' -------------------------------------------------------
    Do
        userInput = InputBox( _
            "Enter the full path to the PowerShell script:" & vbCrLf & _
            "(Edit or accept the default below)", _
            "PowerShell Script Path", _
            PS_SCRIPT_DEFAULT)

        ' Empty string = Cancel pressed
        If userInput = "" Then
            MsgBox "Operation cancelled.", vbInformation
            Exit Sub
        End If

        psScriptPath = Trim(userInput)

        If Dir(psScriptPath) <> "" Then
            Exit Do   ' valid path — continue
        Else
            MsgBox "PowerShell script not found:" & vbCrLf & psScriptPath & vbCrLf & vbCrLf & _
                   "Please check the path and try again, or press Cancel to exit.", _
                   vbExclamation, "Script Not Found"
        End If
    Loop

    ' -------------------------------------------------------
    ' Step 2: Mode selection — retry until S, B, or Cancelled.
    ' -------------------------------------------------------
    Do
        userInput = InputBox( _
            "Choose mode:" & vbCrLf & _
            "S = Single keyword search" & vbCrLf & _
            "B = Batch mode from Excel", _
            "Outlook Keyword Search")

        If userInput = "" Then
            MsgBox "Operation cancelled.", vbInformation
            Exit Sub
        End If

        modeChoice = UCase(Trim(userInput))

        If modeChoice = "S" Or modeChoice = "B" Then
            Exit Do
        Else
            MsgBox "Invalid choice. Please enter S or B.", vbExclamation, "Invalid Mode"
        End If
    Loop

    ' -------------------------------------------------------
    ' Step 3a: Single mode — keyword prompt.
    ' -------------------------------------------------------
    If modeChoice = "S" Then
        Do
            userInput = InputBox("Enter keyword or phrase to search in email BODY:", "Single Keyword Search")

            If userInput = "" Then
                MsgBox "Operation cancelled.", vbInformation
                Exit Sub
            End If

            keyword = Trim(userInput)

            If Len(keyword) > 0 Then
                Exit Do
            Else
                MsgBox "Keyword cannot be blank. Try again or press Cancel to exit.", vbExclamation
            End If
        Loop

        keyword = Replace(keyword, """", "'")
        psCmd = "powershell.exe -ExecutionPolicy Bypass -File """ & psScriptPath & _
                """ -Mode S -Keyword """ & keyword & """"

    ' -------------------------------------------------------
    ' Step 3b: Batch mode — Excel file path, then column.
    ' -------------------------------------------------------
    ElseIf modeChoice = "B" Then

        ' --- Excel file path retry loop ---
        Do
            userInput = InputBox( _
                "Enter full Excel file path (e.g. C:\Users\You\keywords.xlsx):", _
                "Batch Mode — File Path")

            If userInput = "" Then
                MsgBox "Operation cancelled.", vbInformation
                Exit Sub
            End If

            filePath = Trim(userInput)

            If Dir(filePath) <> "" Then
                Exit Do   ' valid file — continue
            Else
                MsgBox "File not found:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
                       "Please check the path and try again, or press Cancel to exit.", _
                       vbExclamation, "File Not Found"
            End If
        Loop

        ' --- Column letter retry loop ---
        Do
            userInput = InputBox( _
                "Enter column letter containing keywords (e.g. A):", _
                "Batch Mode — Keyword Column")

            If userInput = "" Then
                MsgBox "Operation cancelled.", vbInformation
                Exit Sub
            End If

            colRef = UCase(Trim(userInput))

            ' Validate: must be 1-2 alpha characters only (A-Z or AA-XFD)
            If colRef Like "[A-Z]" Or colRef Like "[A-Z][A-Z]" Then
                Exit Do
            Else
                MsgBox "Invalid column reference '" & colRef & "'. Enter a letter such as A or AB.", _
                       vbExclamation, "Invalid Column"
            End If
        Loop

        psCmd = "powershell.exe -ExecutionPolicy Bypass -File """ & psScriptPath & _
                """ -Mode B -FilePath """ & filePath & """ -Column """ & colRef & """"
    End If

    ' -------------------------------------------------------
    ' Step 4: Launch PS in background and confirm to user.
    ' -------------------------------------------------------
    Shell psCmd, vbHide

    MsgBox "Search started in background." & vbCrLf & vbCrLf & _
           "Outlook remains fully usable while the search runs." & vbCrLf & _
           "You will receive a Windows notification when the search completes." & vbCrLf & _
           "For batch mode, results are written to your Excel file automatically.", _
           vbInformation, "Background Search Running"
End Sub
