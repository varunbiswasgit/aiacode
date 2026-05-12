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
'           Update PS_SCRIPT_PATH below to match your path.
' ============================================================

Private Const PS_SCRIPT_PATH As String = "C:\Users\Varun\scripts\OutlookKeywordSearch_PS.ps1"

Public Sub RunKeywordSearch()
    Dim modeChoice As String
    Dim keyword    As String
    Dim filePath   As String
    Dim colRef     As String
    Dim psCmd      As String

    If Dir(PS_SCRIPT_PATH) = "" Then
        MsgBox "PowerShell script not found:" & vbCrLf & PS_SCRIPT_PATH & vbCrLf & vbCrLf & _
               "Update PS_SCRIPT_PATH in the VBA module and try again.", _
               vbCritical, "Script Not Found"
        Exit Sub
    End If

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
            psCmd = "powershell.exe -ExecutionPolicy Bypass -File """ & PS_SCRIPT_PATH & _
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
            psCmd = "powershell.exe -ExecutionPolicy Bypass -File """ & PS_SCRIPT_PATH & _
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
