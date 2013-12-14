VERSION 5.00
Begin VB.Form frmPortForwarding 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Port Forwarding"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10830
   Begin VB.CommandButton cmdVoice 
      Caption         =   "Forward Voice Rec. Port"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdDP 
      Caption         =   "Forward Display Picture Port"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdFileTransfer 
      Caption         =   "Forward File Transfer Port"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton cmdStick 
      Caption         =   "Forward Stick Server Port"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdSpace 
      Caption         =   "Forward Space Server Port"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.Frame fraManual 
      Caption         =   "Manual Forwarding"
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   10575
      Begin VB.PictureBox picForward 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   10335
         TabIndex        =   12
         Top             =   240
         Width           =   10335
         Begin VB.CommandButton cmdForward 
            Caption         =   "Forward Port"
            Enabled         =   0   'False
            Height          =   690
            Left            =   9360
            TabIndex        =   21
            Top             =   0
            Width           =   855
         End
         Begin VB.ComboBox cboProtocol 
            Height          =   315
            Left            =   960
            TabIndex        =   18
            Text            =   "cboProtocol"
            Top             =   360
            Width           =   4695
         End
         Begin VB.TextBox txtIP 
            Height          =   285
            Left            =   6720
            TabIndex        =   16
            Top             =   0
            Width           =   2415
         End
         Begin VB.TextBox txtDesc 
            Height          =   285
            Left            =   960
            TabIndex        =   14
            Top             =   0
            Width           =   4695
         End
         Begin VB.TextBox txtPort 
            Height          =   285
            Left            =   6720
            TabIndex        =   20
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Port:"
            Height          =   255
            Left            =   6120
            TabIndex        =   19
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Protocol:"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblDesc 
            Caption         =   "Description:"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblIP 
            Caption         =   "IP:"
            Height          =   255
            Left            =   6360
            TabIndex        =   15
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.CheckBox chkShowDisabled 
      Caption         =   "Show Disabled Ports"
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdForwardMain 
      Caption         =   "Forward Main/Communicator Port"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Selected Port"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   600
      Width           =   2175
   End
   Begin projMulti.ScrollListBox lstEnabled 
      Height          =   2775
      Left            =   0
      TabIndex        =   26
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   4895
   End
   Begin projMulti.ScrollListBox lstDescription 
      Height          =   2775
      Left            =   1080
      TabIndex        =   25
      Top             =   3360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4895
   End
   Begin projMulti.ScrollListBox lstExtPort 
      Height          =   2775
      Left            =   3720
      TabIndex        =   28
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   4895
   End
   Begin projMulti.ScrollListBox lstProtocol 
      Height          =   2775
      Left            =   6120
      TabIndex        =   30
      Top             =   3360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   4895
   End
   Begin projMulti.ScrollListBox lstIntPort 
      Height          =   2775
      Left            =   4920
      TabIndex        =   27
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   4895
   End
   Begin projMulti.ScrollListBox lstExtIP 
      Height          =   2775
      Left            =   6960
      TabIndex        =   29
      Top             =   3360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4895
   End
   Begin projMulti.ScrollListBox lstIntIP 
      Height          =   2775
      Left            =   8880
      TabIndex        =   31
      Top             =   3360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4895
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      Caption         =   "Your router's UPnP function must be enabled for this application to interface with it."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   8175
   End
   Begin VB.Label lblIPAddresses 
      Caption         =   "IP Address"
      Height          =   180
      Left            =   5880
      TabIndex        =   23
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblPorts 
      Caption         =   "Ports"
      Height          =   195
      Left            =   4440
      TabIndex        =   22
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   10695
   End
   Begin VB.Label lblTitles 
      Caption         =   "Enabled                    Description                       External    Internal  Protocol      External           Internal"
      Height          =   255
      Left            =   15
      TabIndex        =   24
      Top             =   3060
      Width           =   7455
   End
End
Attribute VB_Name = "frmPortForwarding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

'http://www.knoxscape.com/Upnp/NAT.htm

Private Enum eProtocols
    TCP = 0
    UDP = 1
End Enum


Private Sub EnableCmds(ByVal En As Boolean)

cmdForwardMain.Enabled = En
cmdStick.Enabled = En
cmdSpace.Enabled = En
cmdFileTransfer.Enabled = En
cmdDP.Enabled = En
cmdVoice.Enabled = En

cmdRemove.Enabled = En
cmdRefresh.Enabled = En
chkShowDisabled.Enabled = En
cmdForward.Enabled = En

End Sub

Private Sub SetStatus(ByVal T As String)
If InStr(1, T, vbLf, vbTextCompare) Then
    T = Replace$(T, vbLf, " - ", , , vbTextCompare)
End If

lblStatus.Caption = "Status: " & T
lblStatus.Refresh
End Sub

Private Sub cboProtocol_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub chkShowDisabled_Click()
cmdRefresh_Click
End Sub

Private Sub cmdForward_Click()

Dim Desc As String, IP As String
Dim tPort As Long, Protocol As eProtocols
Dim i As Integer
Dim bCan As Boolean

EnableCmds False

Desc = txtDesc.Text
IP = txtIP.Text
If LenB(txtPort.Text) > 0 Then
    On Error GoTo EH
    tPort = CLng(txtPort.Text)
End If

Protocol = IIf(cboProtocol.Text = "UDP", eProtocols.UDP, eProtocols.TCP)

If LenB(Desc) = 0 Or LenB(IP) = 0 Or tPort = 0 Then
    SetStatus "Error - Please enter all the fields"
Else
    
    bCan = True
    
    For i = 0 To lstExtPort.ListCount - 1
        If lstExtPort.List(i) = CStr(tPort) Then
            If lstProtocol.List(i) = GetProtocol(Protocol) Then
                SetStatus "Error - Port already forwarded"
                bCan = False
                Exit For
            End If
        End If
    Next i
    
    If bCan Then
        If ForwardPort(tPort, Desc, IP, Protocol) Then
            SetStatus "Port Forwarded"
            Call ListPorts
        'Else
            'error already shown
        End If
    End If
    
End If

EnableCmds True

Exit Sub
EH:
SetStatus "Error - Please make the port smaller"
EnableCmds True
End Sub

Private Sub cmdRefresh_Click()
EnableCmds False
Call ListPorts
EnableCmds True
End Sub

Private Sub cmdRemove_Click()
Dim LPort As Long
Dim Protocol As eProtocols

Dim ListI As Integer

EnableCmds False

ListI = lstIntPort.ListIndex

If ListI <> lstProtocol.ListIndex Then
    SetStatus "Error - Please reselect the port to remove"
ElseIf ListI = -1 Then
    SetStatus "Please select a port to remove"
Else
    Protocol = IIf(lstProtocol.List(lstProtocol.ListIndex) = "TCP", _
        eProtocols.TCP, eProtocols.UDP)
    
    
    LPort = CLng(lstIntPort.List(lstIntPort.ListIndex))
    
    
    RemovePort LPort, Protocol
    
End If

EnableCmds True

End Sub

'------------------------------------------------------------------------------------

Private Sub ForwardGamePort(ByVal PortToForward As Integer, sName As String)
Text_ForwardPort PortToForward, sName, UDP
End Sub

Private Sub Text_ForwardPort(ByVal PortToForward As Integer, sName As String, vProt As eProtocols)
txtDesc.Text = App.ProductName & " - " & sName
txtIP.Text = modWinsock.LocalIP
cboProtocol.ListIndex = CInt(vProt)
txtPort.Text = PortToForward
End Sub

Private Sub cmdForwardMain_Click()
Text_ForwardPort modPorts.MainPort, "Main", TCP
End Sub
Private Sub cmdSpace_Click()
ForwardGamePort modPorts.SpacePort, "Space Game"
End Sub
Private Sub cmdStick_Click()
ForwardGamePort modPorts.StickPort, "Stick Game"
End Sub
Private Sub cmdFileTransfer_Click()
Text_ForwardPort modPorts.FTPort, "File Transfer", TCP
End Sub
Private Sub cmdDP_Click()
Text_ForwardPort modPorts.DPPort, "Display Picture", TCP
End Sub
Private Sub cmdVoice_Click()
Text_ForwardPort modPorts.VoicePort, "Voice", TCP
End Sub

Private Sub Form_Load()
'Call ListPorts
Call FormLoad(Me)

With cboProtocol
    .AddItem "TCP"
    .AddItem "UDP"
    .ListIndex = 0
End With

txtIP.Text = modWinsock.LocalIP

SetStatus "List the ports before forwarding"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub

Private Function CreateObjs(ByRef oNat As Object, _
    ByRef oMappingPorts As Object) As Boolean


'AddConsoleText "Creating Network + Port Objects...", , , , True

On Error GoTo EH
Set oNat = CreateObject("HNetCfg.NATUPnP") 'New UPnPNAT

If Not (oNat Is Nothing) Then
    
    'AddConsoleText "Created Network Object"
    Set oMappingPorts = oNat.StaticPortMappingCollection
    
    
    'If Not (oMappingPorts Is Nothing) Then
        'AddConsoleText "Created Port Object"
        
        'AddConsoleText "Created Network + Port Objects"
        
        CreateObjs = True
    'Else
        'CreateObjs = False
        'AddConsoleText "Error Creating Mapping Ports Object"
    'End If
Else
    CreateObjs = False
    'AddConsoleText "Error Creating UPnP Router Object"
End If

Exit Function
EH:
CreateObjs = False
SetStatus "Error Creating Network Objects - " & Err.Description
End Function

Private Sub ListPorts()
Dim theNatter As Object 'NATUPNPLib.UPnPNAT
Dim MappingPorts As Object 'IStaticPortMappingCollection
Dim MappingPort As Object 'IStaticPortMapping
Dim ShowDisabled As Boolean, CanAdd As Boolean
Dim i As Integer, j As Integer

SetStatus "Listing Ports..."

ShowDisabled = CBool(chkShowDisabled.Value)

lstEnabled.Clear
lstDescription.Clear
lstExtPort.Clear
lstIntPort.Clear
lstProtocol.Clear
lstExtIP.Clear
lstIntIP.Clear

Me.Refresh

If CreateObjs(theNatter, MappingPorts) Then
    
    If (MappingPorts Is Nothing) = False Then
        For Each MappingPort In MappingPorts
        'For i = 1 To MappingPorts.Count
            
            'MsgBox (MappingPort.Description & vbspace & MappingPort.ExternalPort & vbspace & MappingPort.Protocol & vbspace & MappingPort.InternalPort)
            On Error Resume Next
            
            If Not (MappingPort Is Nothing) Then
                CanAdd = (MappingPort.Enabled Or ShowDisabled)
                
'                For j = 0 To lstExtPort.ListCount - 1
'                    If lstExtPort.List(j) = MappingPort.ExternalPort Then
'                        If lstProtocol.List(j) = MappingPort.Protocol Then
'                            CanAdd = False
'                            Exit For
'                        End If
'                    End If
'                Next j
                
                If CanAdd Then
                    lstEnabled.AddItem CStr(MappingPort.Enabled)
                    lstDescription.AddItem CStr(MappingPort.Description)
                    lstExtPort.AddItem CStr(MappingPort.ExternalPort)
                    lstIntPort.AddItem CStr(MappingPort.InternalPort)
                    lstProtocol.AddItem MappingPort.Protocol
                    lstExtIP.AddItem MappingPort.ExternalIPAddress
                    lstIntIP.AddItem MappingPort.InternalClient
                End If
            End If
            
        Next
        SetStatus "Listed Ports"
    Else
        SetStatus "No Ports Forwarded"
    End If
    
    cmdForward.Enabled = True
    cmdForwardMain.Enabled = True
    cmdStick.Enabled = True
    cmdSpace.Enabled = True
    
Else
    SetStatus "Error: " & IIf(LenB(Err.Description) > 0, Err.Description, "Unknown")
End If

Set MappingPort = Nothing
Set MappingPorts = Nothing
Set theNatter = Nothing


End Sub

Private Function ForwardPort(ByVal ExternalPort As Long, _
    ByVal Name As String, ByVal IP As String, ByVal Protocol As eProtocols, _
    Optional ByVal InternalPort As Long = -1) As Boolean

Dim theNatter As Object
Dim MappingPorts As Object

ForwardPort = True

If InternalPort = -1 Then InternalPort = ExternalPort

'This part creates the upnp object
On Error GoTo EH

If CreateObjs(theNatter, MappingPorts) Then
    
    If Not (MappingPorts Is Nothing) Then
        On Error Resume Next
        MappingPorts.Remove InternalPort, GetProtocol(Protocol)  'remove then add
        
        On Error GoTo NatError
        MappingPorts.Add InternalPort, GetProtocol(Protocol), ExternalPort, IP, True, Name
        'Internal/Private Port, Protocol, External/Inbound Port, Internal IP, Enabled, Name
    Else
        SetStatus "Error forwarding port - Couldn't obtain link to list of ports"
        ForwardPort = False
    End If
    
Else
    SetStatus "Error forwarding port - Couldn't access router via UPnP"
    ForwardPort = False
End If

Set MappingPorts = Nothing
Set theNatter = Nothing

Exit Function

EH:
ForwardPort = False
SetStatus "Error - " & Err.Description
Exit Function

NatError:
ForwardPort = False
MsgBoxEx "'HNetCfg.NATUPnP' Error" & vbNewLine & _
    Err.Description & vbNewLine & "(Port may be too large)", "Oh dear. Been trying to break Communicator on purpose, have we?", _
    vbExclamation, "uPnP Error", , , frmMain.Icon

End Function

Private Sub RemovePort(ByVal LPort As Long, ByVal Protocol As eProtocols)
Dim theNatter As Object
Dim MappingPorts As Object

If CreateObjs(theNatter, MappingPorts) Then
    
    Err.Clear
    On Error Resume Next
    MappingPorts.Remove LPort, GetProtocol(Protocol)
    '                     ^Private/Internal Port
    
    
    Call ListPorts
    If Err.Number Then
        SetStatus Err.Description
    Else
        SetStatus "Removed Port " & CStr(LPort) & " Protocol: " & GetProtocol(Protocol) & " (It may just be disabled)"
    End If
End If

Set MappingPorts = Nothing
Set theNatter = Nothing

End Sub

Private Function GetProtocol(ByVal Protocol As eProtocols) As String
GetProtocol = IIf(Protocol = eProtocols.TCP, "TCP", "UDP")
End Function

Private Sub listClick(ByVal i As Integer)
Static Doing As Boolean

If Not Doing Then
    Doing = True
    Err.Clear
    On Error Resume Next
    lstEnabled.ListIndex = i
    lstDescription.ListIndex = i
    lstExtPort.ListIndex = i
    lstIntPort.ListIndex = i
    lstProtocol.ListIndex = i
    lstExtIP.ListIndex = i
    lstIntIP.ListIndex = i
    
    cmdRemove.Enabled = (Err.Number = 0)
    
    Doing = False
End If

End Sub

Private Sub lstDescription_Click()
Call listClick(lstDescription.ListIndex)
End Sub

Private Sub lstEnabled_Click()
Call listClick(lstEnabled.ListIndex)
End Sub

Private Sub lstExtIP_Click()
Call listClick(lstExtIP.ListIndex)
End Sub

Private Sub lstExtPort_Click()
Call listClick(lstExtPort.ListIndex)
End Sub

Private Sub lstIntIP_Click()
Call listClick(lstIntIP.ListIndex)
End Sub

Private Sub lstIntPort_Click()
Call listClick(lstIntPort.ListIndex)
End Sub

Private Sub lstProtocol_Click()
Call listClick(lstProtocol.ListIndex)
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
End If
End Sub


'Private Function GetLocalIP() As String
'Dim IPAddrs() As String
'
'IPAddrs = GetIpAddrTable()
'
'GetLocalIP = IPAddrs(UBound(IPAddrs))
'
'End Function
'
''Get IPs
''I got this part off of allapi.net
''I have 2 ips, a network one and one for my NIC card along
''with my real ip, so this gets the ip on the network instead
''of the nic card (i use wireless)
'
'Public Function GetIpAddrTable() As String
'Dim Buf(0 To 511) As Byte
'Dim BufSize As Long, rc As Long
'Dim NrOfEntries As Integer, i As Integer, j As Integer
'
'BufSize = UBound(Buf) + 1
'
'rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
'
'If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & CStr(rc)
'
'NrOfEntries = Buf(1) * 256 + Buf(0)
'
'If NrOfEntries = 0 Then
'    GetIpAddrTable = Array()
'Else
'    ReDim IPAddrs(0 To NrOfEntries - 1) As String
'
'    For i = 0 To NrOfEntries - 1
'        s = vbNullString
'
'        For j = 0 To 3
'            s = s & IIf(j > 0, ".", vbNullString) & Buf(4 + i * 24 + j)
'        Next j
'
'        IPAddrs(i) = s
'
'    Next i
'    GetIpAddrTable = IPAddrs
'End If
'
'End Function
