Attribute VB_Name = "OutlookKeywordSearch_PS"
Option Explicit

' ============================================================
' Outlook Keyword Search — PowerShell-Assisted VBA Launcher
' Author  : Varun Biswas
' Repo    : varunbiswasgit/aiacode
' Change log:
'   2026-05-12  Retry loops on all prompts; exit only on Cancel.
'   2026-05-12  Dialog shows PowerShell PID after launch.
'   2026-05-12  All key events echoed to Immediate window
'               (Ctrl+G in VBA editor) via Debug.Print so the
'               user can monitor progress without any extra tool.
' ============================================================

Private Const PS_SCRIPT_DEFAULT As String = "C:\Users\Varun\scripts\OutlookKeywordSearch_PS.ps1"

' Helper: timestamp prefix for Immediate window lines
Private Function TS() As String
    TS = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] "
End Function

Public Sub RunKeywordSearch()
    Dim psScriptPath As String
    Dim modeChoice   As String
    Dim keyword      As String
    Dim filePath     As String
    Dim colRef       As String
    Dim psCmd        As String
    Dim userInput    As String
    Dim pid          As Long

    Debug.Print TS & "=== Outlook Keyword Search started ==="

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
            Debug.Print TS & "Cancelled at PS script path prompt."
            MsgBox "Operation cancelled.", vbInformation
            Exit Sub
        End If

        psScriptPath = Trim(userInput)
        Debug.Print TS & "PS script path entered: " & psScriptPath

        If Dir(psScriptPath) <> "" Then
            Debug.Print TS & "PS script path validated OK."
            Exit Do
        Else
            Debug.Print TS & "PS script NOT FOUND: " & psScriptPath
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
            Debug.Print TS & "Cancelled at mode selection."
            MsgBox "Operation cancelled.", vbInformation
            Exit Sub
        End If

        modeChoice = UCase(Trim(userInput))
        Debug.Print TS & "Mode selected: " & modeChoice

        If modeChoice = "S" Or modeChoice = "B" Then
            Exit Do
        Else
            Debug.Print TS & "Invalid mode entry: " & modeChoice
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
                Debug.Print TS & "Cancelled at keyword prompt."
                MsgBox "Operation cancelled.", vbInformation
                Exit Sub
            End If

            keyword = Trim(userInput)

            If Len(keyword) > 0 Then
                Debug.Print TS & "Keyword entered: " & keyword
                Exit Do
            Else
                Debug.Print TS & "Blank keyword — prompting again."
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
                Debug.Print TS & "Cancelled at Excel file path prompt."
                MsgBox "Operation cancelled.", vbInformation
                Exit Sub
            End If

            filePath = Trim(userInput)
            Debug.Print TS & "Excel file path entered: " & filePath

            If Dir(filePath) <> "" Then
                Debug.Print TS & "Excel file validated OK."
                Exit Do
            Else
                Debug.Print TS & "Excel file NOT FOUND: " & filePath
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
                Debug.Print TS & "Cancelled at column prompt."
                MsgBox "Operation cancelled.", vbInformation
                Exit Sub
            End If

            colRef = UCase(Trim(userInput))
            Debug.Print TS & "Column entered: " & colRef

            If colRef Like "[A-Z]" Or colRef Like "[A-Z][A-Z]" Then
                Debug.Print TS & "Column validated OK: " & colRef
                Exit Do
            Else
                Debug.Print TS & "Invalid column: " & colRef
                MsgBox "Invalid column reference '" & colRef & "'. Enter a letter such as A or AB.", _
                       vbExclamation, "Invalid Column"
            End If
        Loop

        psCmd = "powershell.exe -ExecutionPolicy Bypass -File """ & psScriptPath & _
                """ -Mode B -FilePath """ & filePath & """ -Column """ & colRef & """"
    End If

    ' -------------------------------------------------------
    ' Step 4: Launch PS, capture PID, echo details to console.
    ' -------------------------------------------------------
    Debug.Print TS & "Launching PowerShell..."
    Debug.Print TS & "Command: " & psCmd

    pid = Shell(psCmd, vbHide)

    If pid = 0 Then
        Debug.Print TS & "ERROR: Shell() returned PID=0 — process did not start."
        MsgBox "ERROR: PowerShell process could not be started." & vbCrLf & vbCrLf & _
               "Command attempted:" & vbCrLf & psCmd, _
               vbCritical, "Launch Failed"
        Exit Sub
    End If

    Debug.Print TS & "PowerShell launched successfully."
    Debug.Print TS & "Process : powershell.exe"
    Debug.Print TS & "PID     : " & pid
    Debug.Print TS & "Mode    : " & modeChoice
    If modeChoice = "S" Then
        Debug.Print TS & "Keyword : " & keyword
    ElseIf modeChoice = "B" Then
        Debug.Print TS & "File    : " & filePath
        Debug.Print TS & "Column  : " & colRef
    End If
    Debug.Print TS & "Check Task Manager > Details tab > PID " & pid & " to confirm process is running."
    Debug.Print TS & "A Windows toast notification will appear when the job completes."

    MsgBox "Search started in background." & vbCrLf & vbCrLf & _
           "Process : powershell.exe" & vbCrLf & _
           "PID     : " & pid & vbCrLf & vbCrLf & _
           "You can verify this in Task Manager > Details tab > PID " & pid & "." & vbCrLf & _
           "Full launch details are in the VBA Immediate window (Ctrl+G)." & vbCrLf & vbCrLf & _
           "Outlook remains fully usable while the search runs." & vbCrLf & _
           "A Windows notification will appear when the job completes." & vbCrLf & _
           "For batch mode, results are written to your Excel file automatically.", _
           vbInformation, "Background Search Running"

    Debug.Print TS & "=== Launcher complete. Waiting for PS toast notification. ==="
End Sub
