Attribute VB_Name = "OutlookKeywordSearch_PS"
Option Explicit

' ============================================================
' Outlook Keyword Search — PowerShell-Assisted VBA Launcher
' Author  : Varun Biswas
' Repo    : varunbiswasgit/aiacode
' Fix     : Replaced Shell() with WScript.Shell.Run() to
'           correctly handle paths/arguments containing spaces.
'           Shell() silently returns PID=0 in that case with
'           no error shown. WScript.Shell.Run() is the correct
'           approach for spawning external processes from VBA.
'           Added Q() helper for clean argument quoting.
'           Added DEBUG_MODE — set True to keep PS window open
'           so any script errors are visible on screen.
' ============================================================

Private Const PS_SCRIPT_DEFAULT As String = "C:\Users\Varun\scripts\OutlookKeywordSearch_PS.ps1"

' Set True to show the PowerShell window (recommended for first run / troubleshooting).
' Set False for normal silent background operation once confirmed working.
Private Const DEBUG_MODE As Boolean = True

' Timestamp prefix for Immediate window lines (Ctrl+G in VBA editor)
Private Function TS() As String
    TS = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] "
End Function

' Wraps a value in double-quotes for safe PS argument passing
Private Function Q(s As String) As String
    Q = Chr(34) & s & Chr(34)
End Function

Public Sub RunKeywordSearch()
    Dim psScriptPath As String
    Dim modeChoice   As String
    Dim keyword      As String
    Dim filePath     As String
    Dim colRef       As String
    Dim psCmd        As String
    Dim userInput    As String
    Dim wsh          As Object
    Dim windowStyle  As Integer

    Debug.Print TS & "=== Outlook Keyword Search started ==="
    Debug.Print TS & "DEBUG_MODE = " & DEBUG_MODE

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
        Debug.Print TS & "PS script path: " & psScriptPath

        If Dir(psScriptPath) <> "" Then
            Debug.Print TS & "PS script path OK."
            Exit Do
        Else
            Debug.Print TS & "PS script NOT FOUND: " & psScriptPath
            MsgBox "PowerShell script not found:" & vbCrLf & psScriptPath & vbCrLf & vbCrLf & _
                   "Please check the path and try again.", _
                   vbExclamation, "Script Not Found"
        End If
    Loop

    ' -------------------------------------------------------
    ' Step 2: Mode — retry until S, B, or Cancelled.
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
        Debug.Print TS & "Mode: " & modeChoice

        If modeChoice = "S" Or modeChoice = "B" Then Exit Do
        MsgBox "Invalid choice. Please enter S or B.", vbExclamation, "Invalid Mode"
    Loop

    ' -------------------------------------------------------
    ' Step 3a: Single keyword
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
                Debug.Print TS & "Keyword: " & keyword
                Exit Do
            End If
            MsgBox "Keyword cannot be blank.", vbExclamation
        Loop

        ' Replace embedded double-quotes to keep PS argument clean
        keyword = Replace(keyword, Chr(34), "'")

        psCmd = "powershell.exe -ExecutionPolicy Bypass" & _
                " -File " & Q(psScriptPath) & _
                " -Mode S" & _
                " -Keyword " & Q(keyword)

    ' -------------------------------------------------------
    ' Step 3b: Batch mode
    ' -------------------------------------------------------
    ElseIf modeChoice = "B" Then

        Do
            userInput = InputBox( _
                "Enter full Excel file path (e.g. C:\Users\You\keywords.xlsx):", _
                "Batch Mode — File Path")

            If userInput = "" Then
                Debug.Print TS & "Cancelled at Excel path prompt."
                MsgBox "Operation cancelled.", vbInformation
                Exit Sub
            End If

            filePath = Trim(userInput)
            Debug.Print TS & "Excel file: " & filePath

            If Dir(filePath) <> "" Then
                Debug.Print TS & "Excel file OK."
                Exit Do
            End If
            MsgBox "File not found:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
                   "Please check the path and try again.", vbExclamation, "File Not Found"
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
            Debug.Print TS & "Column: " & colRef

            If colRef Like "[A-Z]" Or colRef Like "[A-Z][A-Z]" Then
                Debug.Print TS & "Column OK: " & colRef
                Exit Do
            End If
            MsgBox "Invalid column '" & colRef & "'. Enter a letter such as A or AB.", _
                   vbExclamation, "Invalid Column"
        Loop

        psCmd = "powershell.exe -ExecutionPolicy Bypass" & _
                " -File " & Q(psScriptPath) & _
                " -Mode B" & _
                " -FilePath " & Q(filePath) & _
                " -Column " & Q(colRef)
    End If

    ' -------------------------------------------------------
    ' Step 4: Launch via WScript.Shell.Run()
    '
    '   WHY NOT Shell():
    '     VBA Shell() silently fails (returns 0) when the command
    '     string contains spaces in paths, even if double-quoted.
    '     It also hides all PowerShell error output.
    '
    '   WHY WScript.Shell.Run():
    '     Correctly handles quoted arguments with spaces.
    '     windowStyle=1 keeps the PS window visible so errors
    '     are readable. windowStyle=0 hides it once confirmed OK.
    '     bWaitOnReturn=False fires and forgets (non-blocking).
    ' -------------------------------------------------------
    windowStyle = IIf(DEBUG_MODE, 1, 0)   ' 1=visible, 0=hidden

    Debug.Print TS & "Full command: " & psCmd
    Debug.Print TS & "Window style: " & IIf(DEBUG_MODE, "VISIBLE (DEBUG_MODE=True)", "Hidden")

    On Error GoTo LaunchError
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run psCmd, windowStyle, False
    Set wsh = Nothing
    On Error GoTo 0

    Debug.Print TS & "WScript.Shell.Run fired OK."
    Debug.Print TS & "Mode   : " & modeChoice
    If modeChoice = "S" Then
        Debug.Print TS & "Keyword: " & keyword
    Else
        Debug.Print TS & "File   : " & filePath
        Debug.Print TS & "Column : " & colRef
    End If
    Debug.Print TS & "Progress log: Documents\OutlookKeywordSearch.log"
    Debug.Print TS & "Set DEBUG_MODE = False to hide PS window once working."

    MsgBox "Search started." & vbCrLf & vbCrLf & _
           "Mode   : " & modeChoice & vbCrLf & _
           IIf(modeChoice = "S", "Keyword: " & keyword, "File   : " & filePath & vbCrLf & "Column : " & colRef) & vbCrLf & vbCrLf & _
           IIf(DEBUG_MODE, "PowerShell window is VISIBLE (DEBUG_MODE=True)." & vbCrLf & _
               "You can see any errors directly in that window." & vbCrLf & vbCrLf, "") & _
           "Progress is logged to:" & vbCrLf & _
           "  Documents\OutlookKeywordSearch.log" & vbCrLf & vbCrLf & _
           "A Windows notification will appear when complete.", _
           vbInformation, "Search Running"

    Debug.Print TS & "=== Launcher complete ==="
    Exit Sub

LaunchError:
    Debug.Print TS & "LAUNCH ERROR: " & Err.Number & " — " & Err.Description
    MsgBox "ERROR: Could not launch PowerShell." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Command attempted:" & vbCrLf & psCmd, _
           vbCritical, "Launch Failed"
End Sub
