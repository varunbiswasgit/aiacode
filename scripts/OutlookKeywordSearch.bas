Attribute VB_Name = "OutlookKeywordSearch"
Option Explicit

' ============================================================
' Outlook Keyword Search Macro — VBA LAUNCHER (v2)
' Author  : Varun Biswas
' Repo    : varunbiswasgit/aiacode
' Purpose : Thin launcher that collects user inputs and
'           delegates all search work to OutlookKeywordSearch.ps1
'           running as a separate PowerShell process.
'           Outlook UI remains fully responsive during search.
' Requires: OutlookKeywordSearch.ps1 in the same folder as this
'           script, or update PS_SCRIPT_PATH below.
' ============================================================

' --- CONFIGURE THIS PATH to match where you saved the PS script ---
Private Const PS_SCRIPT_PATH As String = "C:\Users\Varun\scripts\OutlookKeywordSearch.ps1"

' ------------------------------------------------------------
' ENTRY POINT
' ------------------------------------------------------------
Public Sub RunKeywordSearch()
    Dim modeChoice As String
    Dim keyword    As String
    Dim filePath   As String
    Dim colRef     As String
    Dim psCmd      As String

    ' Validate PS script exists
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
            ' Escape double quotes in keyword for PS
            keyword = Replace(keyword, """", "'")
            psCmd = "powershell.exe -ExecutionPolicy Bypass -File """ & PS_SCRIPT_PATH & _
                    """ -Mode S -Keyword """ & keyword & """"

        Case "B"
            filePath = Trim(InputBox("Enter full Excel file path (example: C:\Users\You\keywords.xlsx):", "Batch Mode - File Path"))
            If Len(filePath) = 0 Then
                MsgBox "No file path entered.", vbExclamation
                Exit Sub
            End If
            If Dir(filePath) = "" Then
                MsgBox "File not found: " & filePath, vbCritical
                Exit Sub
            End If
            colRef = Trim(InputBox("Enter column letter containing keywords (example: A):", "Batch Mode - Keyword Column"))
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

    ' Fire and forget — PS runs in background, Outlook stays responsive
    Shell psCmd, vbHide

    MsgBox "Search started in background." & vbCrLf & vbCrLf & _
           "Outlook remains fully usable while the search runs." & vbCrLf & _
           "You will receive a Windows notification when the search completes." & vbCrLf & _
           "For batch mode, results are written to your Excel file automatically.", _
           vbInformation, "Background Search Running"
End Sub
