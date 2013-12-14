VERSION 5.00
Begin VB.Form frmLobby 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lobby"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9645
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox cboStartWeapon 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CheckBox chkStickSound 
      Alignment       =   1  'Right Justify
      Caption         =   "Load Sound"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.ComboBox cboMap 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdNorm 
      Height          =   375
      Index           =   8
      Left            =   600
      TabIndex        =   36
      Top             =   3600
      Width           =   3735
   End
   Begin VB.CommandButton cmdNorm 
      Height          =   375
      Index           =   7
      Left            =   2520
      TabIndex        =   35
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdNorm 
      Height          =   375
      Index           =   6
      Left            =   600
      TabIndex        =   34
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdNorm 
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   32
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdNorm 
      Height          =   375
      Index           =   4
      Left            =   600
      TabIndex        =   31
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdNorm 
      Height          =   375
      Index           =   3
      Left            =   7320
      TabIndex        =   28
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdNorm 
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   27
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdNorm 
      Height          =   375
      Index           =   1
      Left            =   7320
      TabIndex        =   24
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdNorm 
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   23
      Top             =   3120
      Width           =   1815
   End
   Begin projMulti.XPButton cmdXP 
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   8
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Edit Space Game"
      CapAlign        =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXP 
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Host Space Game"
      CapAlign        =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin VB.Frame fraOptions 
      Caption         =   "IP Options"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4920
         ScaleHeight     =   615
         ScaleWidth      =   2415
         TabIndex        =   2
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton optnRemote 
            Caption         =   "Send Remote IP to others"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   2175
         End
         Begin VB.OptionButton optnLocal 
            Caption         =   "Send Local IP to others"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   7320
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmLobby.frx":0000
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
   End
   Begin projMulti.IPTextBox txtStickIP 
      Height          =   285
      Left            =   960
      TabIndex        =   30
      Top             =   3840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
   End
   Begin projMulti.IPTextBox txtIP 
      Height          =   285
      Left            =   5760
      TabIndex        =   20
      Top             =   2640
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
   End
   Begin VB.TextBox txtStickGame 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   500
      Left            =   3960
      Top             =   1440
   End
   Begin VB.TextBox txtGame 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2280
      Width           =   3375
   End
   Begin projMulti.ScrollListBox lstStick 
      Height          =   975
      Left            =   960
      TabIndex        =   33
      Top             =   4200
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2355
   End
   Begin projMulti.ScrollListBox lstGames 
      Height          =   2175
      Left            =   5760
      TabIndex        =   37
      Top             =   3000
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3836
   End
   Begin projMulti.XPButton cmdXP 
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Join Space Game"
      CapAlign        =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXP 
      Height          =   375
      Index           =   3
      Left            =   7080
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Remove Selected"
      CapAlign        =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXP 
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   10
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Host Stick Game"
      CapAlign        =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXP 
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   11
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Join Stick Game"
      CapAlign        =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXP 
      Height          =   375
      Index           =   6
      Left            =   480
      TabIndex        =   12
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "'Install' DirectX"
      CapAlign        =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXP 
      Height          =   375
      Index           =   7
      Left            =   2280
      TabIndex        =   13
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Remove Selected"
      CapAlign        =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXP 
      Height          =   375
      Index           =   8
      Left            =   480
      TabIndex        =   14
      Top             =   2160
      Width           =   3975
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Map Editor"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin VB.Label lblStartWeapon 
      Caption         =   "Starting Weapon:"
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   3050
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   4800
      X2              =   4800
      Y1              =   1200
      Y2              =   5160
   End
   Begin VB.Label lblStickGame 
      Caption         =   "Game:"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblStickIP 
      Caption         =   "IP:"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblIP 
      Caption         =   "IP:"
      Height          =   255
      Left            =   5160
      TabIndex        =   19
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblGame 
      Caption         =   "Game:"
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   2280
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   4800
      X2              =   4800
      Y1              =   5160
      Y2              =   1200
   End
End
Attribute VB_Name = "frmLobby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefaultStatusCap = "Loaded Window and Refreshed Game List", _
    DX7VB_Dll = "dx7vb.dll", Reg_Cmd = "regsvr32 dx7vb.dll"

Private Type ptDef_Build
    sName As String
    iWeapon1 As eWeaponTypes
    iWeapon2 As eWeaponTypes
    iPerk As eStickPerks
End Type

Private Builds() As ptDef_Build
Private Const weapon_Sep As String = "-----", ServerGameStr = "Server's Game (?)"


'modSpaceGame.ProcessLobbyCmd elobbycmds.Refresh & "Timm#1.2.3.4@Greg#5.6.7.8@Jim#9.109.11.12"

Private Function GetIPToDist() As String

If optnRemote.Value Then
    GetIPToDist = modWinsock.RemoteIP
Else
    GetIPToDist = modWinsock.LocalIP
End If

End Function

Private Sub cmdEdit_Click()
modSpaceGame.SpaceEditing = True

Me.Hide
Unload Me

On Error Resume Next
Load frmGame
End Sub

Private Sub cmdHost_Click()
modSpaceGame.SpaceEditing = False

Call FormLoad(Me, True)
Me.Hide

modSpaceGame.HostSpaceGame GetIPToDist()

Unload Me
'lSetStatus "Hosting Game..."

'Str = "#" & frmMain.SckLC.LocalIP & "," & "Space Combat"

'If Server Then
    'modMessaging.LobbyStr = modMessaging.LobbyStr & Str
'Else
    'SendData eCommands.LobbyCmd & eLobbyCmds.Add & Str
'End If
End Sub

Private Sub cmdJoin_Click()
Dim IP As String

EnableCmd False, 1
IP = Trim$(txtIP.Text)

Me.Hide
Unload Me

modSpaceGame.JoinSpaceGame IP

End Sub

Private Sub cmdRemoveSpace_Click()
Dim IP As String
Dim i As Integer

EnableCmd False, 3
IP = Trim$(txtIP.Text)
txtIP.Text = vbNullString

For i = 0 To UBound(modSpaceGame.CurrentGames)
    If modSpaceGame.CurrentGames(i).bStickGame = False Then
        If modSpaceGame.CurrentGames(i).IP = IP Then
            modSpaceGame.ProcessLobbyCmd eLobbyCmds.Remove & IP
            Exit For
        End If
    End If
Next i

RefreshList
End Sub

Private Sub cmdRemoveStick_Click()
Dim IP As String
Dim i As Integer

EnableCmd False, 7
IP = Trim$(txtStickIP.Text)
txtStickIP.Text = vbNullString

For i = 0 To UBound(modSpaceGame.CurrentGames)
    If modSpaceGame.CurrentGames(i).bStickGame Then
        If modSpaceGame.CurrentGames(i).IP = IP Then
            modSpaceGame.ProcessLobbyCmd eLobbyCmds.Remove & IP & "S"
            Exit For
        End If
    End If
Next i

RefreshList

End Sub

Private Sub cmdMapSelect_Click()
Dim sFile As String
Dim bError As Boolean

frmMain.CommonDPath sFile, bError, "Select a map to load", "Stick Maps (*." & Map_Ext & ")|*." & Map_Ext, _
    modStickGame.GetStickMapPath(), True

If Not bError Then
    cboMap.Text = sFile
    cboMap.Selstart = Len(cboMap.Text)
End If

End Sub

Private Sub cmdXP_Click(Index As Integer)
cmdNorm_Click Index
End Sub
Private Sub cmdNorm_Click(Index As Integer)

Select Case Index
    Case 0
        cmdHost_Click
    Case 1
        cmdJoin_Click
    Case 2
        cmdEdit_Click
    Case 3
        cmdRemoveSpace_Click
    Case 4
        cmdStickHost_Click
    Case 5
        cmdStickJoin_Click
    Case 6
        cmdDX_Click
    Case 7
        cmdRemoveStick_Click
    Case 8
        cmdStickEdit_Click
End Select

End Sub

Private Sub lblStickIP_DblClick()
If Not Server Then
    txtStickGame.Text = ServerGameStr
    txtStickIP.Text = frmMain.SckLC.RemoteHostIP
End If
End Sub

Private Sub lstGames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Name As String
Dim Txt As String
Dim i As Integer

Txt = lstGames.Text
txtGame.Text = Txt

On Error GoTo EH
Name = Left$(Txt, InStrRev(Txt, "'s Game", , vbTextCompare) - 1)

For i = 1 To UBound(modSpaceGame.CurrentGames)
    If modSpaceGame.CurrentGames(i).HostName = Name Then
        txtIP.Text = modSpaceGame.CurrentGames(i).IP
        Exit For
    End If
Next i

EnableCmd CBool(LenB(Txt) And Server), 3

EH:
End Sub

Private Sub optnRemote_Click()

If LenB(modWinsock.RemoteIP) = 0 Then
    lblStatus.Caption = "External IP not obtained. Right click the status bar to obtain it"
    
    lblStatus.height = 615
    lblStatus.Top = 240
    lblStatus.ForeColor = vbRed
    
    optnLocal.Value = True
Else
    SetlblStatusCap
End If

End Sub

Private Sub tmrRefresh_Timer()
RefreshList
End Sub

Private Sub txtIP_Change()

EnableCmd CBool(LenB(txtIP.Text)), 1
    
If modDisplay.CanShow_XPButtons() Then
    cmdXP(0).Default = Not cmdXP(1).Enabled
    cmdXP(1).Default = cmdXP(1).Enabled
Else
    cmdNorm(0).Default = Not cmdNorm(1).Enabled
    cmdNorm(1).Default = cmdNorm(1).Enabled
End If

End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
Dim CurTxt As String
Dim i As Integer, l As Integer
Dim j As Integer

CurTxt = txtIP.Text
l = Len(CurTxt)

If l Then
    txtGame.Text = vbNullString
    
    For i = 1 To UBound(modSpaceGame.CurrentGames)
        If modSpaceGame.CurrentGames(i).bStickGame = False Then
            If Left$(modSpaceGame.CurrentGames(i).IP, l) = CurTxt Then
                txtGame.Text = modSpaceGame.CurrentGames(i).HostName & "'s Game"
                
                For j = 0 To lstGames.ListCount - 1
                    If lstGames.List(j) = txtGame.Text Then
                        lstGames.ListIndex = j
                        Exit For
                    End If
                Next j
                
                Exit For
            End If
        End If
    Next i
End If

End Sub

Private Sub txtIP_LostFocus()
txtIP.Text = Trim$(txtIP.Text)
End Sub
Private Sub SetlblStatusCap()
lblStatus.ForeColor = vbBlue
lblStatus.height = 495
lblStatus.Top = 360
lblStatus.Caption = DefaultStatusCap
End Sub


Private Sub Form_Load()
Dim AppPath_DX7VB As String, sDir As String
Dim i As Integer

lblStatus.BorderStyle = 0
SetlblStatusCap

'########################################################################
AppPath_DX7VB = AppPath() & DX7VB_Dll

cmdXP(6).ToolTipText = "DirectX is required for the stick game sounds"
cmdNorm(6).ToolTipText = cmdXP(6).ToolTipText


If FileExists(DX7VB_Path) Or FileExists(AppPath_DX7VB) Then
    
    cmdXP(6).Caption = "[DirectX Present]"
    cmdNorm(6).Caption = cmdXP(6).Caption
    
    EnableCmd False, 6
End If
'########################################################################

sDir = Dir$(modStickGame.GetStickMapPath())


Do While LenB(sDir)
    If Right$(sDir, 3) = Map_Ext Then
        cboMap.AddItem sDir
    End If
    
    sDir = Dir$()
Loop

'############# ERROR HERE (cboMap.Text = ...) ####################

For i = 0 To cboMap.ListCount - 1
    If cboMap.List(i) = "Default." & Map_Ext Then
        cboMap.ListIndex = i
        Exit For
    End If
Next i

If i = cboMap.ListCount Then
    'Default.map not found in list
    cboMap.AddItem "Default." & Map_Ext
    cboMap.ListIndex = cboMap.ListCount - 1
End If

'cboMap.Text = "Default." & Map_Ext
'############# ERROR HERE ####################


'########################################################################

For i = 0 To eWeaponTypes.Chopper - 1
    If i <> Knife Then
        If i <> USP Then
            cboStartWeapon.AddItem GetWeaponName(CInt(i))
        End If
    End If
Next i
cboStartWeapon.Text = GetWeaponName(modStickGame.cl_StartWeapon1)

cboStartWeapon.AddItem weapon_Sep

InitBuilds
For i = 0 To UBound(Builds)
    cboStartWeapon.AddItem Builds(i).sName
Next i

'########################################################################

If modLoadProgram.bSafeMode Then
    chkStickSound.Enabled = False
    chkStickSound.Value = 0
    chkStickSound.Caption = "[Safe Mode]"
End If


SetTBBanners
SetCmdButtons


Call FormLoad(Me)
Me.Show vbModeless, frmMain


RefreshList


If lstGames.ListCount > 0 Then
    On Error Resume Next
    'txtIP.Text = lstGames.List(lstGames.ListCount - 1)
    lstGames.ListIndex = 0
    lstGames_MouseDown 0, 0, 0, 0
    'lstGames_Click
    
    
    On Error Resume Next
    If modDisplay.CanShow_XPButtons() Then
        SetFocus2 cmdXP(1)
    Else
        SetFocus2 cmdNorm(1)
    End If
    
ElseIf lstStick.ListCount > 0 Then
    
    On Error Resume Next
    lstStick.ListIndex = 0
    lstStick_MouseDown 0, 0, 0, 0
    
    
    If modDisplay.CanShow_XPButtons() Then
        SetFocus2 cmdXP(5)
    Else
        SetFocus2 cmdNorm(5)
    End If
    
End If

End Sub

Private Sub cboStartWeapon_Change()
If cboStartWeapon.Text = weapon_Sep Then
    cboStartWeapon.ListIndex = cboStartWeapon.ListIndex - 1
End If
End Sub
Private Sub cboStartWeapon_Click()
cboStartWeapon_Change
End Sub
Private Sub cboStartWeapon_Scroll()
cboStartWeapon_Change
End Sub
Private Sub cboStartWeapon_LostFocus()
cboStartWeapon_Change
End Sub

Private Sub SetCmdButtons()
Dim i As Integer

If modDisplay.CanShow_XPButtons() Then
    For i = 0 To cmdXP.UBound
        cmdXP(i).Visible = True
        cmdNorm(i).Visible = False
    Next i
Else
    For i = 0 To cmdXP.UBound
        
        cmdXP(i).Visible = False
        cmdNorm(i).Visible = True
        
        cmdNorm(i).Caption = cmdXP(i).Caption
        cmdNorm(i).Top = cmdXP(i).Top
        cmdNorm(i).Enabled = cmdXP(i).Enabled
        
    Next i
End If

End Sub

Private Sub EnableCmd(bEn As Boolean, iCmd As Integer)
Me.cmdXP(iCmd).Enabled = bEn
Me.cmdNorm(iCmd).Enabled = bEn
End Sub

Private Sub SetTBBanners()

modDisplay.SetTextBoxBanner txtGame.hWnd, "Space Game Name Here"
modDisplay.SetTextBoxBanner txtIP.hWnd, "Enter an IP, or select one"

modDisplay.SetTextBoxBanner txtStickGame.hWnd, "Stick Game Name Here"
modDisplay.SetTextBoxBanner txtStickIP.hWnd, "Enter an IP, or select one"

End Sub

Private Sub RefreshList()
Dim i As Integer
Dim Spacei As Integer, Sticki As Integer
'remember list selection

Spacei = lstGames.ListIndex
lstGames.Clear

Sticki = lstStick.ListIndex
lstStick.Clear

For i = 0 To UBound(modSpaceGame.CurrentGames)
    If LenB(modSpaceGame.CurrentGames(i).HostName) Then
        If modSpaceGame.CurrentGames(i).bStickGame Then
            lstStick.AddItem modSpaceGame.CurrentGames(i).HostName & "'s Game"
        Else
            lstGames.AddItem modSpaceGame.CurrentGames(i).HostName & "'s Game"
        End If
    End If
Next i

On Error Resume Next
lstGames.ListIndex = Spacei
lstStick.ListIndex = Sticki

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode <> vbFormCode Then
    Call FormLoad(Me, True)
End If

End Sub

'Private Sub cmdRefresh_Click()
'lSetStatus "Getting Game List..."
'
'If Server Then
'    Call ParseLobbyReply(modMessaging.LobbyStr)
'Else
'    SendData eCommands.LobbyCmd & eLobbyCmds.Request
'End If
'
'End Sub

'Private Sub lstGameType_Click()
'Call ListClick(lstGameType.ListIndex)
'End Sub
'
'Private Sub lstIP_Click()
'Call ListClick(lstIP.ListIndex)
'End Sub
'
'Private Sub ListClick(ByVal i As Integer)
'
'Static bDoing As Boolean
'
'On Error Resume Next
'
'If bDoing = False Then
'    bDoing = True
'    lstIP.ListIndex = i
'    lstGameType.ListIndex = i
'    cmdJoin.Enabled = (i <> -1)
'    bDoing = False
'End If
'
'End Sub

'####################################################################################

Private Sub cmdStickHost_Click()
Dim i As Integer

Call FormLoad(Me, True)
Me.Hide

modAudio.bDXSoundEnabled = CBool(chkStickSound.Value) And Not modLoadProgram.bSafeMode

Call SetStartStats


modStickGame.HostStickGame GetIPToDist(), modStickGame.GetStickMapPath() & cboMap.Text

Unload Me

'lSetStatus "Hosting Game..."

'Str = "#" & frmMain.SckLC.LocalIP & "," & "Space Combat"

'If Server Then
    'modMessaging.LobbyStr = modMessaging.LobbyStr & Str
'Else
    'SendData eCommands.LobbyCmd & eLobbyCmds.Add & Str
'End If
End Sub

Private Sub cmdStickJoin_Click()
Dim IP As String
Dim i As Integer

EnableCmd False, 5
IP = Trim$(txtStickIP.Text)

Call SetStartStats

Me.Hide
Unload Me

modStickGame.JoinStickGame IP

End Sub

Private Sub SetStartStats()
Dim sWeapon As String: sWeapon = cboStartWeapon.Text
Dim i As Integer

For i = 0 To eWeaponTypes.Chopper
    If GetWeaponName(CInt(i)) = sWeapon Then
        modStickGame.cl_StartWeapon1 = i
        modStickGame.cl_StartWeapon2 = USP
        modStickGame.cl_StartPerk = pNone
        Exit For
    End If
Next i

If i = eWeaponTypes.Chopper + 1 Then
    'no weapon chose, check builds
    For i = 0 To UBound(Builds)
        If Builds(i).sName = sWeapon Then
            modStickGame.cl_StartWeapon1 = Builds(i).iWeapon1
            modStickGame.cl_StartWeapon2 = Builds(i).iWeapon2
            modStickGame.cl_StartPerk = Builds(i).iPerk
            Exit For
        End If
    Next i
    
    If i = UBound(Builds) + 1 Then
        modStickGame.cl_StartWeapon1 = AK
        modStickGame.cl_StartWeapon2 = USP
        modStickGame.cl_StartPerk = pNone
    End If
End If


End Sub
Private Sub InitBuilds()

ReDim Builds(0 To 4)

With Builds(0)
    .sName = "Terrorist"
    .iWeapon1 = AK
    .iWeapon2 = RPG
    .iPerk = pMartyrdom
End With
With Builds(1)
    .sName = "SAS CQB"
    .iWeapon1 = MP5
    .iWeapon2 = W1200
    .iPerk = pStoppingPower
End With
With Builds(2)
    .sName = "Sniper"
    .iWeapon1 = AWM
    .iWeapon2 = MP5
    .iPerk = pSniper
End With
With Builds(3)
    .sName = "Marine"
    .iWeapon1 = XM8
    .iWeapon2 = USP
    .iPerk = pSteadyAim
End With
With Builds(4)
    .sName = "Heavy Gunner"
    .iWeapon1 = M249
    .iWeapon2 = DEagle
    .iPerk = pJuggernaut
End With
End Sub

Private Sub cmdStickEdit_Click()
Dim sMap As String

EnableCmd False, 8

Me.Hide
sMap = cboMap.Text
Unload Me
modStickGame.EditStickGame modStickGame.GetStickMapPath() & sMap

End Sub

Private Sub txtStickIP_Change()

txtStickIP.Text = Trim$(txtStickIP.Text)

EnableCmd CBool(LenB(txtStickIP.Text)), 5
    
If modDisplay.CanShow_XPButtons() Then
    cmdXP(4).Default = Not cmdXP(5).Enabled
    cmdXP(5).Default = cmdXP(5).Enabled
Else
    cmdNorm(4).Default = Not cmdNorm(5).Enabled
    cmdNorm(5).Default = cmdNorm(5).Enabled
End If

End Sub

Private Sub txtStickIP_LostFocus()
txtStickIP.Text = Trim$(txtStickIP.Text)
End Sub

Private Sub lstStick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Name As String
Dim Txt As String
Dim i As Integer

Txt = lstStick.Text
txtStickGame.Text = Txt

On Error GoTo EH
Name = Left$(Txt, InStrRev(Txt, "'s Game", , vbTextCompare) - 1)

For i = 1 To UBound(modSpaceGame.CurrentGames)
    If modSpaceGame.CurrentGames(i).bStickGame Then
        If modSpaceGame.CurrentGames(i).HostName = Name Then
            txtStickIP.Text = modSpaceGame.CurrentGames(i).IP
            Exit For
        End If
    End If
Next i

EnableCmd CBool(LenB(Txt) And Server), 7
EH:
End Sub

'##########################################################################################

Private Sub cmdDX_Click()
Dim S As String
Dim Ans As VbMsgBoxResult


EnableCmd False, 6


S = Extract_DX7VB_Dll()

If LenB(S) = 0 Then
    MsgBoxEx "DirectX DLL has been 'installed'" & vbNewLine & _
             "Communicator may need restarting before sounds will work" & vbNewLine & _
             "The DLL may also need registering - '" & Reg_Cmd & "'", _
             "The library file required for Communicator to play sounds has been installed", _
             vbInformation, "DirectX Sound", , , , , Me.hWnd
    
Else
    Ans = MsgBoxEx("Communicator encountered an error installing the DirectX library" & vbNewLine & _
             "Contact " & App.CompanyName & vbNewLine & "(Error: " & S & ")" & vbNewLine & _
             "Extract to Communicator's Folder?", _
             "The library file required for Communicator to play sounds has not been installed", _
             vbQuestion Or vbYesNo, "DirectX Sound", , , , , Me.hWnd)
    
    If Ans = vbYes Then
        S = AppPath() & "dx7vb.dll"
        If LenB(Extract_DX7VB_Dll(S)) = 0 Then
            OpenFolder vbNormalFocus, , S
        End If
    End If
    
    
    
End If

End Sub

Private Property Get DX7VB_Path() As String
Dim WinPath As String

WinPath = Environ$("windir")
If Right$(WinPath, 1) <> "\" Then
    WinPath = WinPath & "\"
End If
DX7VB_Path = WinPath & "system32\" & DX7VB_Dll

End Property

Private Function Extract_DX7VB_Dll(Optional ByVal Path As String) As String
Dim WinPath As String
Dim f As Integer

If LenB(Path) Then
    WinPath = Path
Else
    WinPath = DX7VB_Path()
End If

If FileExists(WinPath) = False Then
    f = FreeFile()
    
    On Error GoTo EH
    Open WinPath For Output As #f
        Print #f, StrConv(LoadResData(101, "CUSTOM"), vbUnicode);
    Close #f
    
    Extract_DX7VB_Dll = IIf(FileExists(WinPath), vbNullString, "Communicator couldn't extract the file")
    
Else
    Extract_DX7VB_Dll = vbNullString
End If


Exit Function
EH:
Close #f
Extract_DX7VB_Dll = Err.Description
End Function
