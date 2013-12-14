Attribute VB_Name = "modLogin"
Option Explicit

Public Const Detail_Sep = "@"
Private psNameUsed As String
Private pbUploaded_User As Boolean

Public Property Get bUploaded_User() As Boolean
bUploaded_User = pbUploaded_User
End Property
Public Property Get sNameUsed() As String
sNameUsed = psNameUsed
End Property

Public Property Get Users_Path() As String
Users_Path = FTP_Online_Users_Path()
End Property

Public Sub AddToFTPList(Optional bStealth As Boolean = True) 'allow user to choose login name?

Dim Txt As String, FName As String, sError As String, sTmpIP As String
Dim i As Integer
Const DateReplace = "-"

'################################################
i = modFTP.iCurrent_FTP_Details
modFTP.iCurrent_FTP_Details = 0 'byethost server

modFTP.FTP_StealthMode = bStealth
'################################################

psNameUsed = frmMain.LastName
FName = modFTP.FTP_Online_Users_Path() & psNameUsed & Dot & FileExt '& vbSpace & Date_to_Filename(DateReplace) & "." & modVars.FileExt

'Txt = "Name: " & LastName & vbNewLine & _
        "Time/Date: " & Now() & vbNewLine & _
        "Internal IP: " & modWinsock.LocalIP & vbNewLine & _
        "External IP: " & modWinsock.RemoteIP

sTmpIP = modWinsock.RemoteIP
If LenB(sTmpIP) = 0 Then
    sTmpIP = "N/A"
End If

Txt = psNameUsed & Detail_Sep & _
    Format$(Time$, "hh:mm am/pm") & Detail_Sep & _
    modVars.GetVersion() & Detail_Sep & _
    modVars.PC_Name & Detail_Sep & _
    modVars.User_Name & Detail_Sep & _
    modWinsock.LocalIP & Detail_Sep & _
    sTmpIP


If modFTP.PutFileStr(Txt, FName, frmMain.mnuFileExit, sError, False) Then
    AddConsoleText "Added to Login List"
    pbUploaded_User = True
Else
    AddConsoleText "Error Adding to Login List - " & sError
    pbUploaded_User = False
End If


'################################################
modFTP.iCurrent_FTP_Details = i
modFTP.FTP_StealthMode = False
'################################################
End Sub

Private Function Date_to_Filename(sReplaceChar As String) As String
Date_to_Filename = Replace$(Replace$(Now(), "/", sReplaceChar), ":", sReplaceChar)
End Function

Public Function RemoveFromFTPList(Optional ByVal sNameToRemove As String = vbNullString, _
    Optional ByVal bStealth As Boolean = True, Optional ByRef sRetError As String) As Boolean

Dim Txt As String, FName As String, sError As String
Dim i As Integer
Const DateReplace = "-"
Dim bIsRemovingMe As Boolean

If LenB(sNameToRemove) = 0 Then
    sNameToRemove = psNameUsed
    bIsRemovingMe = True
Else
    bIsRemovingMe = False
End If


If LenB(sNameToRemove) = 0 Then
    Exit Function
End If

'################################################
i = modFTP.iCurrent_FTP_Details
modFTP.iCurrent_FTP_Details = 0 'byethost server
modFTP.FTP_StealthMode = bStealth
'################################################


FName = modFTP.FTP_Online_Users_Path() & sNameToRemove

If LenB(modFTP.FTP_Details(modFTP.iCurrent_FTP_Details).FTP_File_Ext) Then
    FName = FName & Dot & FileExt
End If


If modFTP.DelFTPFile(FName, sError) Then
    sRetError = "Deleted " & sNameToRemove & " from Login List"
    AddConsoleText sRetError
    RemoveFromFTPList = True
Else
    sRetError = "Error Deleting from Login List - " & sError
    AddConsoleText sRetError
    RemoveFromFTPList = False
End If

If bIsRemovingMe Then
    pbUploaded_User = False
    psNameUsed = vbNullString
End If

'################################################
modFTP.iCurrent_FTP_Details = i
modFTP.FTP_StealthMode = False
'################################################
End Function
