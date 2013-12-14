VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ClientFrm 
   Caption         =   "Client"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheck 
      Left            =   6480
      Top             =   3840
   End
   Begin MSWinsockLib.Winsock SckCheck 
      Left            =   7440
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SckOnline 
      Left            =   5880
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton bntListen 
      Caption         =   "Listen"
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   240
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock SckListen 
      Left            =   2880
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton bntSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton bntExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton bntConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Tag             =   "Connect"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   6120
      Width           =   3735
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "123"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtLog 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
   End
   Begin MSWinsockLib.Winsock SockAr 
      Index           =   0
      Left            =   5400
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Winsock Example by VirusFree - http://www.phoenixbit.com"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   -240
      TabIndex        =   9
      Top             =   6600
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   "Remote Host Port :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Remote Host IP :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "ClientFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SocketCounter As Long
Dim Server As Boolean
Const pOnline As Integer = 124

Private Sub bntConnect_Click()
On Error GoTo t

'SckListen is the name of our Winsock ActiveX Control

SckListen.Close 'we close it in case it was trying to connect

'txtIP is the textbox holding the host IP
'txtIP can contain both hostnames ( like www.google.com ) or IPs ( like 127.0.0.1 )
SckListen.RemoteHost = txtIP    'set the remote host to the ip we wrote
                            'in the txtIP textbox

'txtPort is the textbox holding the Port number
SckListen.RemotePort = txtPort  'set the port we want to connect to
                            '( the server must be listening on this port too)
                            
                            
SckListen.Connect               'try to connect


Exit Sub
t:
MsgBox "Error : " & Err.Description, vbCritical
End Sub

Private Sub bntListen_Click()
Dim n As Integer
On Error Resume Next
'close and unload all previous sockets
For n = 1 To SocketCounter
    SockAr(n).Close
    Unload SockAr(n)
Next

On Error GoTo t

'SckListen(0) is the name of our Winsock ActiveX Control

SockAr(0).Close 'we close it in case it listening before


'txtPort is the textbox holding the Port number
SockAr(0).LocalPort = txtPort  'set the port we want to listen to
                              '( the client will connect on this port too)
                            
                            
SockAr(0).Listen                'Start Listening


txtLog.Text = "Listening on Port " & txtPort & vbCrLf

Server = True

Exit Sub
t:
MsgBox "Error : " & Err.Description, vbCritical
End Sub

Private Sub bntExit_Click()
Unload Me
End Sub

Private Sub bntSend_Click()
On Error GoTo t
'we want to send the contents of txtSend textbox


If Server Then
    
    Call DataArrival(txtSend.Text)
    
    
Else
    SckListen.SendData txtSend  'trasmits the string to host


    'we have send the data to the server by we
    'also need to add them to our Chat Buffer
    'so we can se what we wrote
    AddText "Client : " & txtSend

End If

'and then we clear the txtSend textbox so the
'user can write the next message
txtSend = ""

'error handling
'( for example , we will get an error if try to send
'  any data without being connected )
Exit Sub
t:
MsgBox "Error : " & Err.Description
SckListen_Close   'close the connection
End Sub

Private Sub Form_Load()
SckOnline.LocalPort = pOnline
End Sub

Private Sub ShowOnline(Optional ByVal Activate As Boolean = True)

If Activate Then
    On Error Resume Next
    SckOnline.Listen
    If SckOnline.State <> sckListening Then
        MsgBox "Error"
    End If
    
Else
    SckOnline.Close
    
End If

End Sub

Private Sub SckOnline_ConnectionRequest(ByVal requestID As Long)

SckOnline.Accept requestID

SckOnline.SendData SckListen.LocalHostName

SckOnline.Close

End Sub

Private Sub sockAr_Close(index As Integer)
'handles the closing of the connection

SockAr(index).Close  'close connection

Unload SockAr(index) 'unload control

AddText "Client" & index & " -> *** Disconnected"

End Sub

Private Sub SckListen_Close()
'handles the closing of the connection

SckListen.Close  'close connection

AddText "*** Disconnected"

End Sub

Private Sub SckListen_Connect()
'txtLog is the textbox used as our
'chat buffer.

'SckListen.RemoteHost returns the hostname( or ip ) of the host
'SckListen.RemoteHostIP returns the IP of the host

txtLog.Text = "Connected to " & SckListen.RemoteHostIP & vbCrLf

End Sub

Private Sub SckListen_DataArrival(ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

Dim dat As String     'where to put the data

SckListen.GetData dat, vbString   'writes the new data in our string dat ( string format )

'add the new message to our chat buffer
AddText dat

End Sub

Private Sub SckListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'this event is to handle any kind of errors
'happend while using winsock

'Number gives you the number code of that specific error
'Description gives you string with a simple explanation about the error

'append the error message in the chat buffer
AddText "*** Error : " & Description

'and now we need to close the connection
SckListen_Close

'you could also use SckListen.close function but I
'prefer to call it within the SckListen_Close functions that
'handles the connection closing in general

End Sub

Private Sub sockar_ConnectionRequest(index As Integer, ByVal requestID As Long)
'txtLog is the textbox used as our log.

'this event is triggered when a client try to connect on our host
'we must accept the request for the connection to be completed,
'but we will create a new control and assign it to that, so
'SckListen(0) will still be listening for connection but
'SckListen(SocketCounter) , our new sock , will handle the current
'request and the general connection with the client

'increase counter
SocketCounter = SocketCounter + 1

'this will create a new control with index equal to SocketCounter
Load SockAr(SocketCounter)

'with this we accept the connection and we are now connected to
'the client and we can start sending/receiving data
SockAr(SocketCounter).Accept requestID

'add to the log
txtLog.Text = "Client Connected. IP : " & SockAr(0).RemoteHostIP & " , Client Nick : Client" & SocketCounter & vbCrLf

'tell our client his assigned nickname
SockAr(SocketCounter).SendData "Your Nick is ""Client" & SocketCounter & """"

End Sub

Private Sub sockar_DataArrival(index As Integer, ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

Dim dat As String     'where to put the data

SockAr(index).GetData dat, vbString   'writes the new data in our string dat ( string format )

Call DataArrival(dat, index)

End Sub

Private Sub DataArrival(ByVal Data As String, Optional ByVal index As Integer = (-1))
Dim Str As String
Dim n As Integer

'add the new message to our chat buffer
If index = (-1) Then
    Str = "Server : " & Data
Else
    Str = "Client" & index & " : " & Data
End If

AddText Str

'now the client says something, wich arrived at the server...
'the server must now redistibute this message to all other connected
'clients...
On Error Resume Next    'Error Handler
For n = 1 To SocketCounter
    If n <> index Then   'we don't want to send the msg back to the sender :)
        If SockAr(n).State = sckConnected Then   'if socket is connected
            SockAr(n).SendData Str
            DoEvents
        End If
    End If
Next

End Sub

Private Sub sockar_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'this event is to handle any kind of errors
'happend while using winsock

'Number gives you the number code of that specific error
'Description gives you string with a simple explanation about the error

'append the error message in the chat buffer
AddText "*** Error ( Client" & index & ") : " & Description

'and now we need to close the connection
sockAr_Close index

'you could also use sockar(Index).close function but i
'prefer to call it within the sockar_Close functions that
'handles the connection closing in general


End Sub

Private Sub AddText(ByVal Text As String)

txtLog.SelStart = Len(txtLog.Text)
txtLog.SelText = Text & vbCrLf

End Sub

Private Sub tmrCheck_Timer()
SckCheck.RemotePort = pOnline

SckCheck.Connect

End Sub

Private Sub txtSend_Change()
If Len(txtSend.Text) = 0 Then
    bntSend.Enabled = False
    bntSend.Default = False
Else
    bntSend.Enabled = True
    bntSend.Default = True
End If
End Sub
