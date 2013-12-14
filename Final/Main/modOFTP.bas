Attribute VB_Name = "modFTP"
Option Explicit

'for ftp
Private Const FTP_Host_Name As String = "ftp.byethost13.com"
Private Const FTP_User_Name As String = "b13_1256618"
Private Const FTP_Password As String = "communicator"
Private Const FTP_Location As String = "microbsoft.byethost13.com/htdocs"
Private Const FTP_Wait_Time As Integer = 2000

Private Const FTP_IPRemote_File As String = FTP_Location & "/IPs.txt"
Public FTP_IPLocal_File As String 'can't be longer than something - C:\CommunicatorIPs.txt
Public Const Communicator_Rar As String = "Communicator.rar"
Public RootDrive As String

Private Const FTP_UpdateTxt As String = "Version.txt"
Private Const Local_UpdateTxt As String = "Version.txt"

'Public Const UpdatePage As String = "http://microbsoft.byethost13.com/" ' - index.html
'                                     was  ftp_location & "/Index.html"

'Private m_GettingDir As Boolean

Public Sub DownloadFTPFile(ByVal LocalF As String, ByVal RemoteF As String)
Dim host_name As String

'DoEvents

' You must set the URL before the user name and
' password. Otherwise the control cannot verify
' the user name and password and you get the error:
'
'       Unable to connect to remote host

host_name = FTP_Host_Name

If LCase$(Left$(host_name, 6)) <> "ftp://" Then host_name = "ftp://" & host_name

With frmInet.Inet
    .URL = host_name
    
    .UserName = FTP_User_Name
    .Password = FTP_Password
    
    ' Do not include the host name here. That will make
    ' the control try to use its default user name and
    ' password and you'll get the error again.
    On Error GoTo EH
    .Execute , "Get " & _
        RemoteF & " " & LocalF
    
    '    m_GettingDir = True
    '    .Execute , "Dir"
End With

Exit Sub
EH:
AddText "Error - " & Err.Description, TxtError, True
End Sub

Public Sub UploadFTPFile(ByVal LocalF As String, ByVal RemoteF As String)
Dim host_name As String

'DoEvents

' You must set the URL before the user name and
' password. Otherwise the control cannot verify
' the user name and password and you get the error:
'
'       Unable to connect to remote host


host_name = FTP_Host_Name

If LCase$(Left$(host_name, 6)) <> "ftp://" Then host_name = "ftp://" & host_name

With frmInet.Inet
    .URL = host_name
    
    .UserName = FTP_User_Name
    .Password = FTP_Password
    
    ' Do not include the host name here. That will make
    ' the control try to use its default user name and
    ' password and you'll get the error again.
    On Error GoTo EH
    .Execute , "Put " & _
        LocalF & " " & RemoteF
    
    '    m_GettingDir = True
    '    .Execute , "Dir"
End With

Exit Sub
EH:
AddText "Error - " & Err.Description, TxtError, True
End Sub

Public Function DownloadIPs() As String
Call DownloadFTPFile(FTP_IPLocal_File, FTP_IPRemote_File)

Pause FTP_Wait_Time

Dim f As Integer
Dim Str As String

f = FreeFile()

On Error Resume Next
CarryOn:
On Error GoTo EH
Open FTP_IPLocal_File For Input As #f
    Str = Input(LOF(f), f)
Close #f
On Error GoTo 0

On Error Resume Next
Kill FTP_IPLocal_File
On Error GoTo 0

DownloadIPs = Str

Exit Function
EH:
If InStr(1, Err.Description, "permission", vbTextCompare) Then
    Err.Clear
    GoTo CarryOn
End If
End Function

Public Sub UploadIPs()
Call UploadFTPFile(FTP_IPLocal_File, FTP_IPRemote_File)
End Sub

Public Function GetVersion() As String
Call DownloadFTPFile( _
        modFTP.RootDrive & "\" & Local_UpdateTxt, _
        "/" & FTP_Location & "/" & FTP_UpdateTxt)


Pause FTP_Wait_Time 'allow the system time to do the file


Dim f As Integer
Dim Str As String

f = FreeFile()

On Error Resume Next
CarryOn:
On Error GoTo EH
Open (modFTP.RootDrive & "\" & Local_UpdateTxt) For Input As #f
    Str = Input(LOF(f), f)
Close #f
On Error GoTo 0

GetVersion = Str

On Error Resume Next
Kill (modFTP.RootDrive & "\" & Local_UpdateTxt)
On Error GoTo 0

Exit Function
EH:
If InStr(1, Err.Description, "permission", vbTextCompare) Then
    Err.Clear
    GoTo CarryOn
End If
End Function

Public Function DownloadLatest() As Boolean

DownloadLatest = False

Call DownloadFTPFile(modFTP.RootDrive & "\" & Communicator_Rar, _
                    "/" & FTP_Location & "/" & Communicator_Rar)

Pause FTP_Wait_Time

If Dir$(RootDrive & "\" & Communicator_Rar) <> vbNullString Then
    DownloadLatest = True
End If

End Function



'Private Sub inet_StateChanged(ByVal State As Integer)
'Select Case State
'    Case icError
'        addmessage "Error: " & _
'            "    " & Inet.ResponseCode & vbCrLf & _
'            "    " & Inet.ResponseInfo
'    Case icNone
'        addmessage "None"
'    Case icConnecting
'        addmessage "Connecting"
'    Case icConnected
'        addmessage "Connected"
'    Case icDisconnecting
'        addmessage "Disconnecting"
'    Case icDisconnected
'        addmessage "Disconnected"
'    Case icRequestSent
'        addmessage "Request Sent"
'    Case icRequesting
'        addmessage "Requesting"
'    Case icReceivingResponse
'        addmessage "Receiving Response"
'    Case icRequestSent
'        addmessage "Request Sent"
'    Case icResponseReceived
'        addmessage "Response Received"
'    Case icResolvingHost
'        addmessage "Resolving Host"
'    Case icHostResolved
'        addmessage "Host Resolved"
'
'    Case icResponseCompleted
'        addmessage Inet.ResponseInfo
'
'
'   Case Else
'        addmessage "State = " & Format$(State)
'End Select
'End Sub
'
'Private Sub addmessage(ByVal Str As String)
'Debug.Print Str
'End Sub

