VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUDP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Scan"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7245
   Begin projMulti.ScrollListBox lstName 
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3413
   End
   Begin VB.CommandButton cmdInvite 
      Caption         =   "Invite"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame fraDev 
      Caption         =   "Dev Commands"
      Height          =   1335
      Left            =   1320
      TabIndex        =   9
      Top             =   3120
      Width           =   5055
      Begin VB.PictureBox picCmds 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   4695
         TabIndex        =   10
         Top             =   240
         Width           =   4695
         Begin VB.CommandButton cmdSendDisco 
            Caption         =   "Send Disconnect Command"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            TabIndex        =   14
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton cmdSendConn 
            Caption         =   "Send Connect Command"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            TabIndex        =   12
            Top             =   0
            Width           =   2295
         End
         Begin VB.CommandButton cmdForceF 
            Caption         =   "Force them to Listen/Host"
            Enabled         =   0   'False
            Height          =   375
            Left            =   0
            TabIndex        =   13
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton cmdForceListen 
            Caption         =   "Force Listen (If Idle)"
            Enabled         =   0   'False
            Height          =   375
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   2295
         End
      End
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6720
      Top             =   0
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect to Selected"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock SckUDP 
      Left            =   0
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan for Communicators"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin projMulti.ScrollListBox lstIP 
      Height          =   1935
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   1815
      _ExtentX        =   4048
      _ExtentY        =   3413
   End
   Begin projMulti.ScrollListBox lstStatus 
      Height          =   1935
      Left            =   3840
      TabIndex        =   6
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3413
   End
   Begin projMulti.ScrollListBox lstVersion 
      Height          =   1935
      Left            =   5760
      TabIndex        =   7
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   3413
   End
   Begin VB.Label lblTitles 
      Caption         =   $"frmUDP.frx":0000
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   7215
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status: "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   7215
   End
End
Attribute VB_Name = "frmUDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Listening As Boolean
Private AllDev As Boolean 'all devcmds enabled?

Private Const FindPort As Integer = 4296
Private Const ResponsePort As Integer = FindPort - 1

Private Const Hi As String = "MESSAGE|Hello?"
Private Const ForceListen As String = "MESSAGE|Listen!"
Private Const ForceListen2 As String = "MESSAGE|Listen!/f"
Private Const ForceConnect As String = "MESSAGE|Connect!" '& ip
Private Const ForceDisco As String = "MESSAGE|Disco!"
Private Const Invitation As String = "MESSAGE|Invite"
Public UDPInfo As String '= "MESSAGE|Info"

'Private Declare Function SendMessageByLong Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
    wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Private Const LB_SETHORIZONTALEXTENT = &H194

Private Sub cmdConnect_Click()
Dim Text As String

Text = lstIP.Text

If LenB(Text) Then
    Unload Me
    frmMain.Connect Text
End If

End Sub

Private Sub cmdForceListen_Click()

cmdForceListen.Enabled = False

On Error GoTo EH
SendToSingle lstIP.Text, ForceListen

SetStatus "Sent ForceListen Message"

Call UDPListen

Exit Sub
EH:
MsgBoxEx "Error: " & Err.Description, "Random Error! Oh no", _
    vbExclamation, "Error", , , frmMain.Icon
End Sub

Private Sub cmdInvite_Click()

cmdInvite.Enabled = False

On Error Resume Next
SendInvite lstIP.Text

SetStatus "Invite Sent"

End Sub

Public Sub cmdScan_Click()

cmdScan.Enabled = False
tmrTimeOut.Enabled = True
'cmdConnect.Enabled = False <--done below
Call ClearList
Call UDPBroadcast
If bDevMode Then AddConsoleText "UDP Scanned - State: " & GetState(SckUDP.State)

End Sub

Private Sub cmdForceF_Click()


cmdForceF.Enabled = False

On Error GoTo EH
SendToSingle lstIP.Text, ForceListen2

Call UDPListen

SetStatus "Sent ForceListen /f Message + Listening"



Exit Sub
EH:
MsgBoxEx "Error: " & Err.Description, "Random Error, Oh No!", vbExclamation _
    , "Error", , , frmMain.Icon
End Sub

Private Sub cmdSendConn_Click()

cmdSendConn.Enabled = False
cmdSendDisco.Enabled = False

On Error GoTo EH
SendToSingle lstIP.Text, ForceConnect & SckUDP.LocalIP

Call UDPListen

SetStatus "Sent Message + Listening"


Exit Sub
EH:
SetStatus "Error: " & Err.Description
End Sub
Private Sub cmdSendDisco_Click()

cmdSendConn.Enabled = False
cmdSendDisco.Enabled = False

On Error GoTo EH
SendToSingle lstIP.Text, ForceDisco

Call UDPListen

SetStatus "Sent Message + Listening"

Exit Sub
EH:
SetStatus "Error: " & Err.Description
End Sub

Private Sub Form_Load()
'If bDevMode Then addConsoleText "Beginning UDP Listening...", , True
UDPInfo = "MESSAGE|Info"

Call UDPListen

'AddConsoleText "UDP Listening Complete", , , True

Call FormLoad(Me, , False)

End Sub

Public Sub ShowForm()
Dim i As Integer

Me.Left = frmMain.Left + frmMain.width / 2 - Me.width / 2
Me.Top = frmMain.Top + frmMain.height / 2 - Me.height / 2

If bDevMode = False Then
    Me.height = 3630 '3210
    fraDev.Visible = False
Else
    Me.height = 4965
    fraDev.Visible = True
End If


cmdScan.Default = True

cmdConnect.Enabled = False
cmdForceListen.Enabled = False
cmdForceF.Enabled = False
cmdInvite.Enabled = False
cmdSendConn.Enabled = False
cmdSendDisco.Enabled = False

SetStatus vbNullString 'listening


If frmMain.mnuOptionsAdvDisplayVistaControls.Checked Then
    modDisplay.SetButtonIcon cmdConnect.hWnd, frmMain.GetCommandIconHandle()
    modDisplay.SetButtonIcon cmdScan.hWnd, frmMain.GetCommandIconHandle()
    modDisplay.SetButtonIcon cmdInvite.hWnd, frmMain.GetCommandIconHandle()
    
    cmdConnect.Caption = "Connect"
    cmdScan.Caption = "Network Scan"
    
End If


'modImplode.ImplodeForm Me.hWnd, True
modImplode.AnimateAWindow Me.hWnd, aRandom

Show vbModeless, frmMain


SetFocus2 cmdScan

End Sub

Public Sub UDPListen()
'Dim i As Integer

SetStatus vbNullString 'listening

Listening = True
With SckUDP
    .Close
    .Protocol = sckUDPProtocol
    .RemoteHost = "255.255.255.255"
    .LocalPort = FindPort
    .RemotePort = ResponsePort
    
    'start listening for UDP packets
    On Error Resume Next
    'AddConsoleText "UDP Binding Socket..."
    .bind FindPort
End With

'i = SckUDP.State

If bDevMode Then AddConsoleText "Network Broadcast Socket State: " & GetState(SckUDP.State) '& " (" & CStr(i) & ")"

End Sub

Public Sub SendToSingle(ByVal IP As String, ByVal Msg As String, _
    Optional ByVal Notify As Boolean = True)

If Notify Then SetStatus "Sending..."

Listening = False

With SckUDP
    .Close
    .Protocol = sckUDPProtocol
    .RemoteHost = IP
    .LocalPort = ResponsePort
    .RemotePort = FindPort
    On Error Resume Next
    .bind ResponsePort
    
    .SendData Msg
End With

If bDevMode Then AddConsoleText "Sent UDP Data: " & Msg & " to " & IP

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetInactive
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

AllDev = False

With tmrTimeOut
    If .Enabled Then
        tmrTimeOut_Timer
        .Enabled = False
    End If
End With

Call FormLoad(Me, True, Me.Visible)

Call ClearList

If modVars.Closing = False Then
    Cancel = True
    Me.Hide
End If

End Sub

Private Sub ListClick(ByVal i As Integer)
lstName.ListIndex = i
lstIP.ListIndex = i
lstStatus.ListIndex = i
lstVersion.ListIndex = i

cmdConnect.Enabled = (lstName.ListIndex <> -1)
cmdConnect.Default = cmdConnect.Enabled
cmdForceListen.Enabled = (bDevMode And CBool(LenB(lstIP.Text)))

cmdInvite.Enabled = cmdConnect.Enabled And (Status = Connected Or Status = eStatus.Listening)

If AllDev Then
    cmdForceF.Enabled = True
    'cmdSendCmd.Enabled = True
End If

End Sub

Private Sub lstComs_DblClick()
If cmdConnect.Enabled Then
    cmdConnect_Click
End If
End Sub

Private Sub lstIP_Click()
Call ListClick(lstIP.ListIndex)
End Sub

Private Sub lstIP_DblClick()
lstName_DblClick
End Sub

Private Sub lstIP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetInactive
End Sub

Private Sub lstName_Click()
Call ListClick(lstName.ListIndex)
End Sub

Private Sub lstName_DblClick()
Call cmdConnect_Click
End Sub

Private Sub lstName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetInactive
End Sub

Private Sub lstStatus_Click()
Call ListClick(lstStatus.ListIndex)
End Sub

Private Sub lstStatus_DblClick()
lstName_DblClick
End Sub

Private Sub lstStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstName_MouseMove Button, Shift, X, Y
End Sub

Private Sub lstVersion_Click()
Call ListClick(lstVersion.ListIndex)
End Sub

Private Sub lstVersion_DblClick()
lstName_DblClick
End Sub

Private Sub lstVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lstName_MouseMove Button, Shift, X, Y
End Sub

'################################################################################################
'################################################################################################
Private Sub SckUDP_DataArrival(ByVal bytesTotal As Long)
Dim Msg As String, IP As String, Name As String, Ver As String
Dim tStatus As eStatus
Dim bHost As Boolean
Dim i As Integer

On Error Resume Next
SckUDP.GetData Msg, vbString, bytesTotal
On Error GoTo 0

If bDevMode Then AddConsoleText "UDP Socket Data Arrived - " & Msg ', , True

'Msg = "REPLY|" & SckUDP.LocalIP & "@" & frmMain.LastName '- testing broadcast reply
'Msg = Hi '<-- testing listen/reply

If Listening Then
    If Msg = Hi Then
        'broadcast message V
        SckUDP.SendData "REPLY|" & SckUDP.LocalIP & _
                                "@" & frmMain.LastName & _
                                "#" & modVars.Status & _
                                IIf(modVars.Server, CStr(1), CStr(0)) & _
                                GetVersion()
        
        'AddConsoleText "UDP Socket - Replied" ', , , True
        Call UDPListen
        
    ElseIf Left$(Msg, Len(Invitation)) = Invitation Then
        
        frmMain.InviteReceived Mid$(Msg, Len(Invitation) + 1)
        
    ElseIf Left$(Msg, Len(UDPInfo)) = UDPInfo Then
        
        If Me.Visible Then
            SetStatus Mid$(Msg, Len(UDPInfo) + 1), True
        Else
            AddText Mid$(Msg, Len(UDPInfo) + 1), TxtReceived, True
        End If
        
    ElseIf Msg = ForceListen Then
        If modVars.Status = Idle Then
            frmMain.Listen False
        End If
    ElseIf Msg = ForceListen2 Then
        'If modVars.Status = Idle Then
        frmMain.Listen False
        'End If
    ElseIf Msg = ForceDisco Then
        frmMain.CleanUp True
        
    ElseIf Msg Like (ForceConnect & "*") Then
        IP = Mid$(Msg, Len(ForceConnect) + 1)
        
        On Error Resume Next
        frmMain.Connect IP
        
    End If
Else 'broadcasting/scanning
    If Msg Like "REPLY|*" Then 'Left$(Msg, 6) = "REPLY|" Then
        
        Err.Clear
        
        i = InStr(1, Msg, "|", vbTextCompare) + 1
        
        On Error GoTo EH
        IP = Mid$(Msg, i, InStr(1, Msg, "@", vbTextCompare) - i)
        
        i = InStr(1, Msg, "#", vbTextCompare)
        Name = Mid$(Msg, InStr(1, Msg, "@", vbTextCompare) + 1, _
            i - InStr(1, Msg, "@", vbTextCompare) - 1)
        
        
        tStatus = CInt(Mid$(Msg, i + 1, 1))
        bHost = CBool(Mid$(Msg, i + 2, 1))
        
        Ver = Mid$(Msg, i + 3)
        On Error GoTo 0
        
        If Err.Number = 0 Then
            AddToList Name, IP, GetStatus(tStatus, bHost), Ver
        End If '15 = LenB(xxx.xxx.xxx.xxx)
        
    End If
End If

EH:
End Sub
'################################################################################################
'################################################################################################
Private Sub UDPBroadcast()

SetStatus "Scanning..."

Listening = False

With SckUDP
    .Close
    .Protocol = sckUDPProtocol
    .RemoteHost = "255.255.255.255"
    .LocalPort = ResponsePort
    .RemotePort = FindPort
    On Error Resume Next
    .bind ResponsePort
    
    .SendData Hi
End With
End Sub

Private Sub tmrTimeOut_Timer()
cmdScan.Enabled = True
SetStatus "Scan Complete"
Call UDPListen
tmrTimeOut.Enabled = False
End Sub

Private Sub SetStatus(ByVal Text As String, Optional ByVal bBold As Boolean = False)

lblStatus.Font.Bold = bBold

'If bRed Then
    'lblStatus.ForeColor = vbRed
'Else
    'lblStatus.ForeColor = &HFF0000
'End If

If LenB(Text) = 0 Then
    Text = "Listening..."
End If

lblStatus.Caption = Text

lblStatus.Refresh
End Sub

Private Sub AddToList(ByVal N As String, ByVal IP As String, ByVal Status As String, ByVal Version As String)

lstName.AddItem IIf(LenB(N) > 0, N, "?")
lstIP.AddItem IIf(LenB(IP) > 0, IP, "?")
lstStatus.AddItem IIf(LenB(Status) > 0, Status, "?")
lstVersion.AddItem IIf(LenB(Version) > 0, Version, "?")

End Sub

Private Sub ClearList()

lstName.Clear
lstIP.Clear
lstStatus.Clear
lstVersion.Clear

cmdConnect.Enabled = False
cmdInvite.Enabled = False

End Sub

Private Sub SendInvite(ByVal SendToIp As String)

Dim HostIP As String


If Status = eStatus.Listening Then
    HostIP = frmMain.SckLC.LocalIP
    
ElseIf Status = Connected Then
    
    If modVars.Server Then
        HostIP = frmMain.SckLC.LocalIP
    Else
        HostIP = frmMain.SckLC.RemoteHostIP
    End If
    
End If

If Len(HostIP) Then
    SendToSingle SendToIp, Invitation & frmMain.LastName & "#" & HostIP & "@" & modWinsock.LocalIP
    
    UDPListen
End If

End Sub
