Attribute VB_Name = "modLogging"
Option Explicit

'0 vblogAuto - iif(NT, Log to event log, log to path)
'1 vbLogOff - logging = false
'16/&H10 vbLogOverwrite - overwrite log file each program start
'32/&H20 vbLogThreadID - add threadID to log
'2 vbLogToFile - force file logging
'3 vbLogToNT - force event logging (If not NT, do nothing)

Private Const vbLogAuto = &H0&
Private Const vbLogOff = &H1&

Public Enum eLogEventTypes
    LogError = 1
    LogWarning = 2
    LogInformation = 4
End Enum

Private pbLogging As Boolean
Private sActLogFile As String

Public Property Get bLogging() As Boolean
bLogging = pbLogging
End Property
Public Property Let bLogging(bVal As Boolean)
pbLogging = bVal

If pbLogging Then
    App.StartLogging GetLogPath(), vbLogAuto
Else
    App.StartLogging GetLogPath(), vbLogOff
End If
End Property

Private Function GetLogPath() As String
GetLogPath = modSettings.GetUserSettingsPath() & "Log.log"
End Function

Public Sub LogEvent(sEvent As String, Optional vLogType As eLogEventTypes = LogInformation)

App.LogEvent sEvent, vLogType

End Sub

Public Sub addToActivityLog(sEvent As String)
Dim f As Integer

If modLoadProgram.bIsIDE Then Exit Sub

If modLoadProgram.frmMain_Loaded Then
    If frmMain.mnuOptionsMessagingLoggingActivity.Checked = False Then Exit Sub
    
    
    If LenB(sActLogFile) = 0 Then
        sActLogFile = frmMain.GetCurrentLogFolder() & frmMain.MakeTimeFile() & " - Activity Log.txt"
    End If
    
    f = FreeFile()
    
    On Error GoTo EH
    Open sActLogFile For Append As #f
        Print #f, Time$() & " - " & sEvent
    Close #f
    f = 0
End If

EH:
If f <> 0 Then
    On Error Resume Next
    Close #f
End If
End Sub
