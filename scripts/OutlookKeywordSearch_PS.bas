Attribute VB_Name = "OutlookKeywordSearch_PS"
Option Explicit

' ============================================================
' Outlook Keyword Search — PowerShell-Assisted VBA Launcher
' Author  : Varun Biswas
' Repo    : varunbiswasgit/aiacode
' Change log:
'   2026-05-12  Retry loops added to all file/path prompts.
'               Macro exits only when user presses Cancel.
'   2026-05-12  VBA confirmation dialog now shows PowerShell
'               PID and process name so user can verify in
'               Task Manager that the job is running.
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
    Dim pid          As Long

    ' -------------------------------------------------------
    ' Step 1: PS script path — retry until valid or Cancelled.
    ' -------------------------------------------------------
    Do
        userInput = InputBox( _
            "Enter the full path to the PowerShell script:" & vbCrLf & _
            "(Edit or accept the default below)", _
            "PowerShell Script Path", _
            PS_SCRIPT_DEFAULT)

        If userInput = "" Then
            MsgBox "Operation cancelled.", vbInformation
            Exit Sub
        End If

        psScriptPath = Trim(userInput)

        If Dir(psScriptPath) <> "" Then
            Exit Do
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
                Exit Do
            Else
                MsgBox "File not found:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
                       "Please check the path and try again, or press Cancel to exit.", _
                       vbExclamation, "File Not Found"
            End If
        Loop

        Do
            userInput = InputBox( _
                "Enter column letter containing keywords (e.g. A):", _
                "Batch Mode — Keyword Column")

            If userInput = "" Then
                MsgBox "Operation cancelled.", vbInformation
                Exit Sub
            End If

            colRef = UCase(Trim(userInput))

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
    ' Step 4: Launch PS and capture PID.
    '         Shell() returns the PID of the launched process.
    '         Show it in the confirmation dialog so the user
    '         can verify the job in Task Manager.
    ' -------------------------------------------------------
    pid = Shell(psCmd, vbHide)

    If pid = 0 Then
        MsgBox "ERROR: PowerShell process could not be started." & vbCrLf & vbCrLf & _
               "Command attempted:" & vbCrLf & psCmd, _
               vbCritical, "Launch Failed"
        Exit Sub
    End If

    MsgBox "Search started in background." & vbCrLf & vbCrLf & _
           "Process : powershell.exe" & vbCrLf & _
           "PID     : " & pid & vbCrLf & vbCrLf & _
           "You can verify this in Task Manager > Details tab > PID " & pid & "." & vbCrLf & vbCrLf & _
           "Outlook remains fully usable while the search runs." & vbCrLf & _
           "A Windows notification will appear when the job completes." & vbCrLf & _
           "For batch mode, results are written to your Excel file automatically.", _
           vbInformation, "Background Search Running"
End Sub
