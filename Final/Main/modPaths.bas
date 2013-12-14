Attribute VB_Name = "modPaths"
Option Explicit

Private pSavedFilesPath As String, plogPath As String

'####################################################################################
Public Property Get SavedFilesPath() As String

If Len(pSavedFilesPath) = 0 Then
    default_savedFilesPath
End If

SavedFilesPath = pSavedFilesPath

End Property
Public Property Let SavedFilesPath(ByVal S As String)
Dim b As Boolean

If LenB(S) Then
    If FileExists(S, vbDirectory) Then
        b = CanAccessPath(S)
    End If
End If

If b Then
    pSavedFilesPath = S
Else
    default_savedFilesPath
End If

End Property
Public Sub default_savedFilesPath()
pSavedFilesPath = modVars.Comm_Safe_Path
End Sub
'####################################################################################
Public Property Get logPath() As String

If Len(plogPath) = 0 Then
    default_logPath
End If

If FileExists(plogPath, vbDirectory) = False Then
    On Error Resume Next
    MkDir plogPath
End If

logPath = plogPath

End Property
Public Property Let logPath(ByVal S As String)
Dim b As Boolean

If LenB(S) Then
    If FileExists(S, vbDirectory) Then
        b = CanAccessPath(S)
    End If
End If

If b Then
    plogPath = S
Else
    default_logPath
End If

End Property
Public Sub default_logPath()
plogPath = modVars.Comm_Safe_Path

If Right$(plogPath, 1) <> "\" Then plogPath = plogPath & "\"

plogPath = plogPath & "Communicator Logs\"
End Sub
'####################################################################################

Private Function CanAccessPath(sPath As String) As Boolean
Dim sFile As String
Dim f As Integer

f = FreeFile()
sFile = sPath & IIf(Right$(sPath, 1) = "\", vbNullString, "\") & "test.tmp"

On Error GoTo EH
Open sFile For Output As #f
Close #f

On Error Resume Next
Kill sFile

CanAccessPath = True

Exit Function
EH:
CanAccessPath = False
End Function
