VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents pSck As MSWinsockLib.Winsock
Attribute pSck.VB_VarHelpID = -1
Private bServer As Boolean
Private Const FileNameSep = "¤"
Private Const Default_FT_Port = 28801

Private pSaveDir As String 'has trailing \

'WINSOCK
'receive vars
Private bFileArriving As Boolean
Private sFileName As String
Private sArriving As String


'send vars
Private bSendingFile As Boolean
Private pbytesSent As Long, pbytesRemaining As Long
'END WINSOCK

Public Enum eTransferStatus
    tDisconnected = 0
    tConnected = 1
    
    tListening = 2
    tConnecting = 3
    
    tSending = 4
    tReceiving = 5
    tSent = 6
    tReceived = 7
End Enum

Public Enum eFTErrors
    FT_Err_Not_Connected = 1
    FT_Err_ListenError = 2
End Enum

Private sCurStatus As String
Private eCurStatus As eTransferStatus

Public Event Diconnected()
Public Event Connected(IP As String)
Public Event SendingFile(sFile As String)
Public Event ReceivingFile(sFileName As String)
Public Event SentFile(sFileName As String)
Public Event ReceivedFile(sFileName As String)
Public Event Error(Description As String)

Public Property Get iCurStatus() As eTransferStatus
iCurStatus = eCurStatus
End Property

Public Property Let SaveDir(sSaveDir As String)

pSaveDir = sSaveDir & IIf(Right$(sSaveDir, 1) = "\", vbNullString, "\")

End Property

Public Property Get SaveDir() As String
SaveDir = pSaveDir
End Property

Private Sub Class_Initialize()

End Sub

'###########################################################################
'winsock subs
'###########################################################################

Private Sub pSck_Close()
If pSck.State <> sckClosed Then pSck.Close

eCurStatus = "Disconnected"
eCurStatus = tDisconnected
RaiseEvent Diconnected
End Sub

Private Sub pSck_ConnectionRequest(ByVal requestID As Long)
' Check if the control's State is closed. If not,
' close the connection before accepting the new
' connection.
If pSck.State <> sckClosed Then pSck.Close

pSck.Accept requestID

eCurStatus = "Connected"
eCurStatus = tConnected
RaiseEvent Connected(pSck.RemoteHostIP)
End Sub

'receiving bit
Private Sub pSck_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Dim hFile As Integer, i As Integer

'pSck.GetData strData, vbString, bytesTotal
'strdata =

If Right$(strData, 7) = "FILEEND" Then
    
    bFileArriving = False
    sCurStatus = "Saving File to " & sFileName
    
    sArriving = sArriving & Left$(strData, Len(strData) - 7)
    
    hFile = FreeFile()
    Open (pSaveDir & sFileName) For Binary Access Write As #hFile
        Put #hFile, 1, sArriving
    Close #hFile
    
    sCurStatus = "Completed Transfer"
    eCurStatus = tReceived
    RaiseEvent ReceivedFile(pSaveDir & sFileName)
    
ElseIf Left$(strData, 4) = "FILE" Then
    
    bFileArriving = True
    sArriving = vbNullString
    
    sFileName = Mid$(strData, 5)
    i = InStr(1, sFileName, FileNameSep, vbTextCompare)
    
    If i Then 'if it's a slow connection, data could be appended, so separate it
        sFileName = Mid$(strData, 4, i - 4)
        sArriving = Mid$(strData, i)
    End If
    
    sCurStatus = "Receiving " & sFileName
    eCurStatus = tReceiving
    RaiseEvent ReceivingFile(sFileName)
    
ElseIf bFileArriving Then
    
    If Left$(strData, 1) = FileNameSep Then
        strData = Mid$(strData, 2)
    End If
    
    sCurStatus = "Receiving " & bytesTotal & " bytes for " & sFileName & " from " '& pSck.RemoteHostIP
    sArriving = sArriving & strData
    
End If
End Sub
'end receiving bit

'sending bit
Public Sub SendFile(sFilePath As String, sRemoteFileName As String)

On Error GoTo EH

Dim sSend As String, sBuf As String
Dim hFile As Integer
Dim lRead As Long, lLen As Long, lThisRead As Long, lLastRead As Long

If eCurStatus = tConnected Then
    
    hFile = FreeFile()
    
    ' Open file for binary access:
    Open sFilePath For Binary Access Read As #hFile
        lLen = LOF(hFile)
        
        'Loop through the file, loading it up in chunks of 64k:
        Do While lRead < lLen
            lThisRead = 65536
            If lThisRead + lRead > lLen Then
                lThisRead = lLen - lRead
            End If
            If Not lThisRead = lLastRead Then
                sBuf = Space$(lThisRead)
            End If
            
            Get #hFile, , sBuf
            lRead = lRead + lThisRead
            sSend = sSend & sBuf
        Loop
        
        'lTotal = lLen
    Close hFile
    
    hFile = 0
    
    bSendingFile = True
    
    sCurStatus = "Sending " & sFilePath
    eCurStatus = tSending
    RaiseEvent SendingFile(sFilePath)
    
    '// Send the file notification
    SendData "FILE" & sRemoteFileName
    DoEvents
    '// Send the file
    SendData FileNameSep & sSend
    DoEvents
    '// Finished
    SendData "FILEEND"
    bSendingFile = False
    
    sCurStatus = "Sent " & sFilePath
    eCurStatus = tSent
    RaiseEvent SentFile(sFilePath)
Else
    RaiseErr FT_Err_Not_Connected, "Can't Send File - Not Connected"
End If

Exit Sub
EH:

sCurStatus = "Send Error - " & Err.Description
eCurStatus = tConnected
RaiseEvent Error(sCurStatus)
End Sub

Private Sub pSck_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
pbytesSent = bytesSent
pbytesRemaining = bytesRemaining
End Sub

Public Sub Connect(IP As String, Optional ByVal Port As Integer = Default_FT_Port)

pSck_Close

On Error Resume Next
pSck.Connect IP, Port

sCurStatus = "Connecting"
eCurStatus = tConnecting

End Sub

Public Sub Listen(Optional ByVal Port As Integer = Default_FT_Port)

pSck.LocalPort = Port
pSck.RemotePort = 0

On Error Resume Next
pSck.Listen

If pSck.State <> sckListening Then
    RaiseErr FT_Err_ListenError, "Error Listening on Port: " & CStr(Port)
Else
    sCurStatus = "Listening"
    eCurStatus = tListening
End If

End Sub

Public Sub Disconnect()
pSck_Close
End Sub

Private Sub RaiseErr(ErrNo As eFTErrors, Desc As String)
Err.Raise CLng(ErrNo), "clsFileTransfer", Desc
End Sub
