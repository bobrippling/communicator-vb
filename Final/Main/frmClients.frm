VERSION 5.00
Begin VB.Form frmClients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clients"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   13080
   Begin VB.Frame fraClients 
      Caption         =   "Clients"
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.PictureBox picDisconnect 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1680
         ScaleHeight     =   495
         ScaleWidth      =   4335
         TabIndex        =   8
         Top             =   3000
         Width           =   4335
         Begin VB.CommandButton cmdCloseSock 
            Caption         =   "Disconnect/Kick Client "
            Enabled         =   0   'False
            Height          =   375
            Left            =   960
            TabIndex        =   9
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.Timer tmrRefresh 
         Interval        =   3000
         Left            =   3720
         Top             =   120
      End
      Begin projMulti.ScrollListBox lstName 
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   4048
      End
      Begin projMulti.ScrollListBox lstIP 
         Height          =   2295
         Left            =   2280
         TabIndex        =   3
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   4048
      End
      Begin projMulti.ScrollListBox lstSocket 
         Height          =   2295
         Left            =   3960
         TabIndex        =   4
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   4048
      End
      Begin projMulti.ScrollListBox lstVersion 
         Height          =   2295
         Left            =   5160
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   4048
      End
      Begin projMulti.ScrollListBox lstPing 
         Height          =   2295
         Left            =   6720
         TabIndex        =   6
         Top             =   480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   4048
      End
      Begin VB.Label lblTitles 
         Caption         =   "Name                                    IP                         Socket                    Version               Ping"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "Status:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   7335
      End
   End
   Begin VB.Frame fraIPs 
      Caption         =   "IPs"
      ForeColor       =   &H00FF0000&
      Height          =   3615
      Left            =   8040
      TabIndex        =   10
      Top             =   120
      Width           =   4935
      Begin VB.PictureBox picIPs 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   240
         ScaleHeight     =   3135
         ScaleWidth      =   4575
         TabIndex        =   11
         Top             =   240
         Width           =   4575
         Begin VB.CheckBox chkAllowAll 
            Caption         =   "Allow All Other IPs"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1920
            Width           =   1815
         End
         Begin VB.CommandButton cmdAllow 
            Caption         =   "Allow IP"
            Enabled         =   0   'False
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtIP 
            Height          =   285
            Left            =   360
            TabIndex        =   13
            Top             =   270
            Width           =   1575
         End
         Begin VB.CommandButton cmdBlock 
            Caption         =   "Block IP"
            Enabled         =   0   'False
            Height          =   375
            Left            =   360
            TabIndex        =   15
            Top             =   1320
            Width           =   1335
         End
         Begin projMulti.ScrollListBox lstBlocked 
            Height          =   2175
            Left            =   2040
            TabIndex        =   16
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   3836
         End
         Begin VB.Label lblBlocked 
            Alignment       =   2  'Center
            Caption         =   "Blocked Info"
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   120
            TabIndex        =   18
            Top             =   2400
            Width           =   4455
         End
         Begin VB.Label lblIP 
            Caption         =   "IP:"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   270
            Width           =   255
         End
      End
   End
End
Attribute VB_Name = "frmClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SelectedClientSock As Integer

Private Sub chkAllowAll_Click()

modMessaging.bAllBlocked = Not CBool(chkAllowAll.Value)
If modMessaging.bAllBlocked Then
    cmdBlock.Enabled = False
    fraIPs.Caption = "Allowed IPs"
    lblBlocked.Caption = "All IPs are now blocked, those in the list are allowed to connect"
Else
    cmdAllow.Enabled = False
    fraIPs.Caption = "Blocked IPs"
    lblBlocked.Caption = "All IPs are now allowed, those in the list are blocked from connecting"
End If

End Sub

Private Sub txtIP_Change()

If modMessaging.bAllBlocked Then
    cmdAllow.Enabled = LenB(txtIP.Text)
Else
    cmdBlock.Enabled = LenB(txtIP.Text)
End If

End Sub

Private Sub lstBlocked_Click()

If modMessaging.bAllBlocked Then
    cmdBlock.Enabled = LenB(lstBlocked.Text)
Else
    cmdAllow.Enabled = LenB(lstBlocked.Text)
End If

End Sub

Private Sub UpdateBlocked()
Dim i As Integer, iCount As Integer

iCount = lstBlocked.ListCount

If iCount = 0 Then
    ReDim modMessaging.BlockedIPs(0)
Else
    ReDim modMessaging.BlockedIPs(0 To iCount - 1)
    
    For i = 0 To iCount - 1
        modMessaging.BlockedIPs(i) = lstBlocked.List(i)
    Next i
End If

End Sub

'############################

Private Sub cmdAllow_Click()
Dim i As Integer
Dim IP As String

cmdAllow.Enabled = False

IP = Trim$(txtIP.Text)

If Not modMessaging.bAllBlocked Then
    'all are allowed, remove from list
    
    lstBlocked.RemoveItem lstBlocked.ListIndex
    
ElseIf CBool(LenB(IP)) Then
    'all are blocked, allow this one
    
    For i = 0 To lstBlocked.ListCount
        If lstBlocked.List(i) = IP Then
            MsgBoxEx "IP is already in the list", "IPs in the list must be unique, otherwise you'd confuse everyone", vbExclamation, "Error", , , frmMain.Icon
            Exit Sub
        End If
    Next i
    
    lstBlocked.AddItem txtIP.Text
    txtIP.Text = vbNullString
End If


UpdateBlocked

End Sub

Private Sub cmdBlock_Click()
Dim i As Integer
Dim IP As String

cmdBlock.Enabled = False

IP = Trim$(txtIP.Text)

If modMessaging.bAllBlocked Then
    'all are blocked, remove from list
    
    lstBlocked.RemoveItem lstBlocked.ListIndex
    
ElseIf CBool(LenB(IP)) Then
    'all are allowed, block this one
    
    For i = 0 To lstBlocked.ListCount
        If lstBlocked.List(i) = IP Then
            MsgBoxEx "IP is already in the list", "IPs in the list must be unique, otherwise you'd confuse everyone", vbExclamation, "Error", , , frmMain.Icon
            Exit Sub
        End If
    Next i
    
    lstBlocked.AddItem txtIP.Text
    txtIP.Text = vbNullString
End If


UpdateBlocked

End Sub

Private Sub cmdCloseSock_Click()
Dim i As Integer

cmdCloseSock.Enabled = False

i = SelectedClientSock
SelectedClientSock = -1

If i <> -1 Then
    frmMain.Kick i, lstName.Text
    tmrRefresh_Timer
End If

End Sub

'Private Sub cmdPing_Click()
'Dim RTT As Long
'Dim i As Integer, j As Integer
'Dim sIP As String
'Dim Acquired As Boolean
'
'cmdPing.Enabled = False
'
'i = SelectedClientSock
'SelectedClientSock = -1
'
'For j = 0 To UBound(Clients)
'    If Clients(j).iSocket = i Then
'        sIP = Clients(j).sIP
'        Exit For
'    End If
'Next j
'
'If LenB(sIP) = 0 Then
'    sIP = lstIP.Text
'    Clients(j).sIP = sIP
'    Acquired = CBool(LenB(sIP))
'ElseIf sIP = "-" Then
'    Acquired = False
'ElseIf LenB(Clients(j).sIP) > 0 Then
'    Acquired = True
'End If
'
'
'If Acquired Then
'
'    SetStatus "Pinging " & sIP
'
'    RTT = modVars.SimplePing(sIP)
'
'    If RTT <> -1 Then
'        SetStatus "Ping Complete - " & IIf(RTT = 0, "< 1", CStr(RTT)) & "ms"
'
'        Clients(j).iPing = IIf(RTT = 0, 1, RTT)
'
'    Else
'        SetStatus "Error Pinging"
'    End If
'Else
'    SetStatus "IP not acquired"
'End If
'
'End Sub

Private Sub SetStatus(ByVal S As String)
lblStatus.Caption = "Status: " & S
lblStatus.Refresh
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim lst As ScrollListBox

SelectedClientSock = -1

SetStatus "Loaded Window"

If frmMain.mnuOptionsAdvDisplayVistaControls.Checked Then
    modDisplay.SetButtonIcon cmdCloseSock.hWnd, frmMain.GetCommandIconHandle()
End If

'########################################################

For i = LBound(modMessaging.BlockedIPs) To UBound(modMessaging.BlockedIPs)
    If LenB(modMessaging.BlockedIPs(i)) Then
        lstBlocked.AddItem modMessaging.BlockedIPs(i)
    End If
Next i

chkAllowAll.Value = Abs(Not modMessaging.bAllBlocked)
chkAllowAll_Click 'get stuff set up
'########################################################


'txtIP.Enabled = Server 'And (Status = Connected)
Call FormLoad(Me)
tmrRefresh_Timer

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub

Private Sub lstIP_Click()
If lstIP.ListIndex <> -1 Then
    If lstIP.Text <> "-" Then
        txtIP.Text = lstIP.Text
    Else
        txtIP.Text = vbNullString
    End If
    Call listClick(lstIP.ListIndex)
End If
End Sub

Private Sub lstName_Click()
Call listClick(lstName.ListIndex)
End Sub

Private Sub lstPing_Click()
Call listClick(lstPing.ListIndex)
End Sub

Private Sub lstSocket_Click()
Call listClick(lstSocket.ListIndex)
End Sub

Private Sub lstVersion_Click()
Call listClick(lstVersion.ListIndex)
End Sub

Private Sub tmrRefresh_Timer()
Dim i As Integer, iList As Integer

If Status <> Connected Then
    If UBound(Clients) > 0 Then
        ReDim Clients(0)
    End If
    lstName.Clear
    lstSocket.Clear
    lstIP.Clear
    lstVersion.Clear
    lstPing.Clear
    
Else
    iList = lstName.ListIndex
    
    lstName.Clear
    lstSocket.Clear
    lstIP.Clear
    lstVersion.Clear
    lstPing.Clear
    
    For i = 0 To UBound(Clients)
        
        If LenB(Clients(i).sName) > 0 Then
            lstName.AddItem Clients(i).sName
        Else
            lstName.AddItem "?"
        End If
        
        
        lstSocket.AddItem CStr(Clients(i).iSocket)
        
        If LenB(Clients(i).sVersion) Then
            lstVersion.AddItem Clients(i).sVersion
        Else
            lstVersion.AddItem "?"
        End If
        
        
        If Clients(i).iPing > 0 Then
            lstPing.AddItem CStr(Clients(i).iPing)
        Else
            lstPing.AddItem "?"
        End If
        
        
        'On Error Resume Next
        If Server Then
            If Clients(i).iSocket <> -1 Then
                On Error Resume Next
                Clients(i).sIP = frmMain.SockAr(Clients(i).iSocket).RemoteHostIP
                
                lstIP.AddItem Clients(i).sIP
            Else
                'it's us
                lstIP.AddItem frmMain.SckLC.LocalIP
            End If
        Else
            If LenB(Clients(i).sIP) Then
                lstIP.AddItem Clients(i).sIP
            Else
                If Clients(i).iSocket <> -1 Then
                    lstIP.AddItem "-"
                Else
                    lstIP.AddItem frmMain.SckLC.RemoteHostIP
                End If
            End If
        End If
        
    Next i
    
    On Error Resume Next
    lstName.ListIndex = iList
End If

End Sub

Private Sub listClick(ByVal i As Integer)
Static Doing As Boolean

Dim j As Integer
Dim sSock As String

If i <> -1 And Not Doing Then
    Doing = True
    On Error Resume Next
    lstName.ListIndex = i
    lstIP.ListIndex = i
    lstSocket.ListIndex = i
    lstVersion.ListIndex = i
    lstPing.ListIndex = i
    
    tmrRefresh.Enabled = False
    tmrRefresh.Interval = tmrRefresh.Interval
    tmrRefresh.Enabled = True 'reset it
    
'###################################
    
    j = lstSocket.ListIndex
    SelectedClientSock = -1
    
    If j <> -1 Then
        sSock = lstSocket.List(i)
        
        If LenB(sSock) > 0 Then
            j = CInt(sSock)
            SelectedClientSock = j
        End If
    End If
    
    cmdCloseSock.Enabled = ((lstSocket.ListIndex <> -1) And Server) And (lstSocket.Text <> "-1")
    
'###################################
    
    Doing = False
End If

End Sub
