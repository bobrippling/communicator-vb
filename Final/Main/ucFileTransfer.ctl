VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl ucFileTransfer 
   BackColor       =   &H000000FF&
   CanGetFocus     =   0   'False
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   615
   ScaleWidth      =   615
   Begin MSWinsockLib.Winsock pSck 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ucFileTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private bServer As Boolean
Private Const DataSep = "Ù"
Private Const CommandSep = "Ø"
Private Const FileNameSep = "¾"
'Private Const Default_FT_Port = 28801
'Private CurPort As Integer

Private pSaveDir As String 'has trailing \
Private pbCloseOnReceived As Boolean 'close connection when received a file
Private pbOverwriteFiles As Boolean

'WINSOCK
'receive vars
Private sArrived As String

'send vars
Private bSendingFile As Boolean, pbSentFile As Boolean
Private pbytesSent As Long, pbytesRemaining As Long, pTotalBytes As Long
Private pCurFileName As String
'END WINSOCK

Public Enum eTransferStatus
    tDisconnected = 0
    tConnected = 1
    
    tlistening = 2
    tConnecting = 3
    
    tSending = 4
    tReceiving = 5
    tSent = 6
    tReceived = 7
End Enum

Public Enum eFTErrors
    FT_Err_Not_Connected = 1
    FT_Err_ListenError = 2
    FT_Err_DirDoesntExist = 3
    FT_Err_CantConnect = 4
    FT_Err_Custom = 5
End Enum

Private sCurStatus As String
Private eCurStatus As eTransferStatus

Public Event Diconnected()
Public Event Connected(IP As String)
Public Event SendingFile(sFileName As String, ByVal BytesSent As Long, ByVal BytesRemaining As Long, ByVal lTotalBytes As Long)
Public Event ReceivingFile(sFileName As String, ByVal BytesReceived As Long, ByVal BytesRemaining As Long, ByVal lTotalBytes As Long)
Public Event SentFile(sFileName As String)
Public Event ReceivedFile(sFileName As String)
Public Event Error(Description As String, ErrNo As eFTErrors)
Public Event ConnectionRequest(IP As String, ByRef bAccept As Boolean)

Public Property Get iCurSockStatus() As StateConstants
iCurSockStatus = pSck.State
End Property

Public Property Get iCurStatus() As eTransferStatus
iCurStatus = eCurStatus
End Property

Public Property Get CurStatus() As String
CurStatus = sCurStatus
End Property

Public Property Let OverwriteFiles(nB As Boolean)
pbOverwriteFiles = nB
End Property

Public Property Let CloseOnReceived(nB As Boolean)
pbCloseOnReceived = nB
End Property

Public Property Let SaveDir(sSaveDir As String)

If LenB(Dir$(sSaveDir, vbDirectory)) Then
    pSaveDir = sSaveDir & IIf(Right$(sSaveDir, 1) = "\", vbNullString, "\")
Else
    RaiseErr FT_Err_DirDoesntExist, "Save Dir Doesn't Exist"
End If

End Property

Public Property Get SaveDir() As String
'always has a trailing \
SaveDir = pSaveDir
End Property

Private Function CanConnect() As Boolean
CanConnect = CBool(LenB(pSaveDir))
End Function

Private Sub pSck_Close()
If pSck.State <> sckClosed Then pSck.Close

sCurStatus = "Disconnected"
eCurStatus = tDisconnected
RaiseEvent Diconnected

ResetVars
End Sub

Private Sub ResetVars()
sArrived = vbNullString
pCurFileName = vbNullString
pbytesRemaining = 0
pbytesSent = 0
pTotalBytes = 0
End Sub

Private Sub pSck_Connect()
sCurStatus = "Connected"
eCurStatus = tConnected
RaiseEvent Connected(pSck.RemoteHostIP)
End Sub

Private Sub pSck_ConnectionRequest(ByVal requestID As Long)
'this event is only raised if pSck is listening - if connected, others can't connect

Dim bAccept As Boolean


If CanConnect() Then
    
    bAccept = True
    'default to accept
    
    RaiseEvent ConnectionRequest(pSck.RemoteHostIP, bAccept)
    
    If bAccept Then
        If pSck.State <> sckClosed Then pSck.Close
        
        pSck.accept requestID
        
        sCurStatus = "Connected"
        eCurStatus = tConnected
        RaiseEvent Connected(pSck.RemoteHostIP)
    End If
    
    'else, stay listening
    
Else
    pSck.Close
End If

End Sub

'receiving bit
Private Sub pSck_DataArrival(ByVal bytesTotal As Long)
Dim strData As String, sFileName As String, sNewFilename As String
Dim hFile As Integer, i As Integer, j As Integer
Dim lTotalBytes As Long
Const File_End_Signature As String = CommandSep & "FILEEND" & CommandSep, kDot = "."

On Error GoTo EH
pSck.GetData strData, vbString, bytesTotal

sArrived = sArrived & strData

i = InStr(1, sArrived, FileNameSep)
On Error Resume Next
lTotalBytes = Mid$(sArrived, i + 1, InStr(1, sArrived, DataSep) - i - 1)


RaiseEvent ReceivingFile(Left$(sArrived, InStr(1, sArrived, FileNameSep) - 1), _
    LenB(sArrived), _
    lTotalBytes - LenB(sArrived), _
    lTotalBytes)


If Right$(sArrived, 9) = File_End_Signature Then
    
    'sArrived looks like this:
    'Timmy.jpg#000@<Data>ØFILEENDØ
    ' ^Name     ^nBytes ^Data ^Marker
    
    '#=FileNameSep
    '@=DataSep
    
    On Error Resume Next
    
    'chop off the command
    sArrived = Left$(sArrived, Len(sArrived) - 9)
    
    'find the filename
    sFileName = Left$(sArrived, InStr(1, sArrived, FileNameSep) - 1)
    sArrived = Mid$(sArrived, InStr(1, sArrived, DataSep) + 1) 'i should still be instr(datasep)
    
    
    sNewFilename = sFileName
    If pbOverwriteFiles Then
        If LenB(Dir$(pSaveDir & sNewFilename)) > 0 Then
            On Error Resume Next
            Kill pSaveDir & sFileName
        End If
    Else
        i = 1
        Do While LenB(Dir$(pSaveDir & sNewFilename)) > 0
            j = InStr(1, sFileName, kDot)
            sNewFilename = Left$(sFileName, j - 1) & " (" & CStr(i) & ")." & Mid$(sFileName, j + 1)
            i = i + 1
        Loop
    End If
    
    
    If Dir$(pSaveDir) = vbNullString Then
        On Error Resume Next
        MkDir pSaveDir
    End If
    
    hFile = FreeFile()
    Open (pSaveDir & sNewFilename) For Binary Access Write As #hFile
        Put #hFile, 1, sArrived 'no need for ";", it's "put", not "print"
    Close #hFile
    
    ResetVars
    
    
    RaiseEvent ReceivedFile(pSaveDir & sNewFilename)
    
    
    If pbCloseOnReceived Then pSck_Close
    
    
    'sCurStatus = "Completed Transfer"
    'eCurStatus = tReceived
    'RaiseEvent ReceivedFile(pSaveDir & sFileName)
End If


Exit Sub
EH:
Disconnect
RaiseErr FT_Err_Custom, Err.Description
End Sub
'If Left$(sData, 6) = (CommandSep & "FILE" & CommandSep) Then
'
'    bFileArriving = True
'    sArriving = vbNullString
'
'    sFileName = Mid$(sData, 7)
'
'    'sCurStatus = "Receiving " & sFileName
'    'eCurStatus = tReceiving
'    'RaiseEvent ReceivingFile(sFileName)
'
'ElseIf Right$(sData, 9) = (CommandSep & "FILEEND" & CommandSep) Then
'
'    bFileArriving = False
'    'sCurStatus = "Saving File to " & sFileName
'
'    'sArriving = sArriving & Left$(sData, Len(sData) - 9) 'chop off our bit
'
'    hFile = FreeFile()
'    Open (pSaveDir & sFileName) For Binary Access Write As #hFile
'        Put #hFile, 1, sArriving
'    Close #hFile
'
'    sArriving = vbNullString
'    sFileName = vbNullString
'
'    'sCurStatus = "Completed Transfer"
'    'eCurStatus = tReceived
'    'RaiseEvent ReceivedFile(pSaveDir & sFileName)
'
'ElseIf bFileArriving Then
'
'    'sCurStatus = "Receiving " & bytesTotal & " bytes for " & sFileName & " from " & pSck.RemoteHostIP
'    sArriving = sArriving & sData
'
'End If

'End Sub
'end receiving bit

'sending bit
Public Function SendFile(sFilePath As String, sRemoteFileName As String) As Boolean

On Error GoTo EH

Dim sSend As String, sBuf As String
Dim hFile As Integer
Dim lRead As Long, lLen As Long, lThisRead As Long, lLastRead As Long

If LenB(Dir$(sFilePath)) Then
    If pSck.State = sckConnected Then
        
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
        Close #hFile
        hFile = 0
        
        pTotalBytes = lLen
        pCurFileName = sFilePath
        pbSentFile = False
        bSendingFile = True
        
        sCurStatus = "Sending " & sFilePath
        eCurStatus = tSending
        RaiseEvent SendingFile(sFilePath, 0, lLen, pTotalBytes)
        
        ''// Send the file notification
        'pSck.SendData (CommandSep & "FILE" & CommandSep) & sRemoteFileName & DataSep
        'DoEvents
        
        
        '                                   pTotalBytes * 2 for LenB(file)
        pSck.SendData sRemoteFileName & FileNameSep & CStr(pTotalBytes * 2) & DataSep & _
                      sSend & (CommandSep & "FILEEND" & CommandSep)
        
        Do
            DoEvents
        Loop While Not pbSentFile And Not modVars.Closing And pSck.State = sckConnected
        
        bSendingFile = False
        
        If pSck.State = sckConnected Then
            sCurStatus = "Sent " & sFilePath
            eCurStatus = tSent
            RaiseEvent SentFile(sFilePath)
            SendFile = True
        Else
            sCurStatus = "Error - Disconnected"
            eCurStatus = tDisconnected
            RaiseErr FT_Err_Custom, "Couldn't Send - Disconnected"
            SendFile = False
        End If
    Else
        RaiseErr FT_Err_Not_Connected, "Can't Send File - Not Connected"
        SendFile = False
    End If
Else
    'RaiseErr FT_Err_Custom, "File to Send Doesn't Exist"
    'On Error GoTo 0
    RaiseErr FT_Err_Custom, "File Not Found"
    SendFile = False
End If

Exit Function
EH:

sCurStatus = "Send Error - " & Err.Description
eCurStatus = tConnected
RaiseEvent Error(sCurStatus, FT_Err_Custom)
End Function

Private Sub pSck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
pSck_Close

RaiseErr FT_Err_Custom, "Error: " & Description

End Sub

Private Sub pSck_SendProgress(ByVal BytesSent As Long, ByVal BytesRemaining As Long)
pbytesSent = BytesSent
pbytesRemaining = BytesRemaining


RaiseEvent SendingFile(pCurFileName, BytesSent, BytesRemaining, pTotalBytes)

If BytesRemaining = 0 Then pbSentFile = True

End Sub

Public Sub Connect(IP As String, ByVal Port As Integer)

If CanConnect() Then
    pSck_Close
    
    pSck.LocalPort = 0
    
    On Error Resume Next
    pSck.Connect IP, Port
    
    sCurStatus = "Connecting"
    eCurStatus = tConnecting
Else
    RaiseErr FT_Err_CantConnect, "Can't Connect - Save Directory Not Set"
End If

End Sub

Public Function Listen(ByVal Port As Integer) As Boolean

Listen = False

If CanConnect() Then
    pSck.Close
    pSck.LocalPort = Port
    pSck.RemotePort = 0
    
    On Error Resume Next
    pSck.Listen
    
    If pSck.State <> sckListening Then
        RaiseErr FT_Err_ListenError, "Error Listening on Port: " & CStr(Port)
    Else
        sCurStatus = "Listening"
        eCurStatus = tlistening
        Listen = True
    End If
Else
    RaiseErr FT_Err_CantConnect, "Can't Connect - Save Directory Not Set"
End If

End Function

Public Sub Disconnect()
pSck_Close
End Sub

Private Sub RaiseErr(ErrNo As eFTErrors, Desc As String)
'Err.Raise CLng(ErrNo), "ucFileTransfer", Desc
RaiseEvent Error(Desc, ErrNo)
End Sub

Private Sub UserControl_Initialize()
pbCloseOnReceived = True
pbOverwriteFiles = False
End Sub

Private Sub UserControl_Resize()
UserControl.width = ScaleX(32, vbPixels, vbTwips)
UserControl.height = ScaleX(32, vbPixels, vbTwips)
End Sub

Private Sub UserControl_Terminate()
pSck.Close
End Sub
