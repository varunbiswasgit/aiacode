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
' ============================================================

' Default path shown in the InputBox as a starting suggestion.
' Change this to match your usual script location.
Private Const PS_SCRIPT_DEFAULT As String = "C:\Users\Varun\scripts\OutlookKeywordSearch_PS.ps1"

Public Sub RunKeywordSearch()
    Dim psScriptPath As String
    Dim modeChoice   As String
    Dim keyword      As String
    Dim filePath     As String
    Dim colRef       As String
    Dim psCmd        As String

    ' -------------------------------------------------------
    ' Step 1: Ask the user for the PowerShell script location.
    '         Pre-fill the InputBox with the default path so
    '         the user can simply press OK if it matches.
    ' -------------------------------------------------------
    psScriptPath = Trim(InputBox( _
        "Enter the full path to the PowerShell script:" & vbCrLf & _
        "(Edit or accept the default below)", _
        "PowerShell Script Path", _
        PS_SCRIPT_DEFAULT))

    If Len(psScriptPath) = 0 Then
        MsgBox "No script path entered. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    If Dir(psScriptPath) = "" Then
        MsgBox "PowerShell script not found:" & vbCrLf & psScriptPath & vbCrLf & vbCrLf & _
               "Please check the path and try again.", _
               vbCritical, "Script Not Found"
        Exit Sub
    End If

    ' -------------------------------------------------------
    ' Step 2: Choose search mode.
    ' -------------------------------------------------------
    modeChoice = Trim(InputBox( _
        "Choose mode:" & vbCrLf & _
        "S = Single keyword search" & vbCrLf & _
        "B = Batch mode from Excel", _
        "Outlook Keyword Search"))

    If Len(modeChoice) = 0 Then Exit Sub

    Select Case UCase(modeChoice)

        Case "S"
            keyword = Trim(InputBox("Enter keyword or phrase to search in email BODY:", "Single Keyword Search"))
            If Len(keyword) = 0 Then
                MsgBox "No keyword entered.", vbExclamation
                Exit Sub
            End If
            keyword = Replace(keyword, """", "'")
            psCmd = "powershell.exe -ExecutionPolicy Bypass -File """ & psScriptPath & _
                    """ -Mode S -Keyword """ & keyword & """"

        Case "B"
            filePath = Trim(InputBox("Enter full Excel file path (e.g. C:\Users\You\keywords.xlsx):", "Batch Mode — File Path"))
            If Len(filePath) = 0 Then
                MsgBox "No file path entered.", vbExclamation
                Exit Sub
            End If
            If Dir(filePath) = "" Then
                MsgBox "File not found: " & filePath, vbCritical
                Exit Sub
            End If
            colRef = Trim(InputBox("Enter column letter containing keywords (e.g. A):", "Batch Mode — Keyword Column"))
            If Len(colRef) = 0 Then
                MsgBox "No column entered.", vbExclamation
                Exit Sub
            End If
            psCmd = "powershell.exe -ExecutionPolicy Bypass -File """ & psScriptPath & _
                    """ -Mode B -FilePath """ & filePath & """ -Column """ & colRef & """"

        Case Else
            MsgBox "Invalid mode. Please enter S or B.", vbExclamation
            Exit Sub

    End Select

    Shell psCmd, vbHide

    MsgBox "Search started in background." & vbCrLf & vbCrLf & _
           "Outlook remains fully usable while the search runs." & vbCrLf & _
           "You will receive a Windows notification when the search completes." & vbCrLf & _
           "For batch mode, results are written to your Excel file automatically.", _
           vbInformation, "Background Search Running"
End Sub
