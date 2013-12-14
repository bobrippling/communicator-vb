VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGameOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Options"
   ClientHeight    =   8085
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClient 
      Caption         =   "Graphics + Client Settings"
      Height          =   375
      Left            =   240
      TabIndex        =   40
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Frame fraServer 
      Caption         =   "Server Settings"
      ForeColor       =   &H000000FF&
      Height          =   4815
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   4695
      Begin VB.PictureBox picServer 
         Height          =   4455
         Left            =   120
         ScaleHeight     =   4395
         ScaleWidth      =   4395
         TabIndex        =   17
         Top             =   240
         Width           =   4455
         Begin VB.TextBox txtScore 
            Height          =   285
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   33
            Top             =   3000
            Width           =   495
         End
         Begin VB.OptionButton optnGameType 
            Caption         =   "Elimination"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   3840
            Width           =   1575
         End
         Begin VB.CommandButton cmdCTFTime 
            Caption         =   "Set Flag Capture Time"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2400
            TabIndex        =   37
            Top             =   3600
            Width           =   1935
         End
         Begin VB.TextBox txtCTFTime 
            Height          =   285
            Left            =   1800
            TabIndex        =   36
            Top             =   3600
            Width           =   495
         End
         Begin VB.OptionButton optnGameType 
            Caption         =   "Capture the Flag"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   3600
            Width           =   1575
         End
         Begin VB.OptionButton optnGameType 
            Caption         =   "Deathmatch"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   3360
            Width           =   1215
         End
         Begin VB.CheckBox chkBulletWalls 
            Caption         =   "Bullets Bounce Off The Edge"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   2640
            Width           =   4095
         End
         Begin VB.CheckBox chkClipMissiles 
            Caption         =   "Missiles Can Be Shot Down"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   2400
            Width           =   4095
         End
         Begin VB.CheckBox chkBulletShipVectorAdd 
            Caption         =   "Bullets Push Ships"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   4095
         End
         Begin VB.CheckBox chkBulletCollisions 
            Caption         =   "Bullets Collide With Each Other"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1920
            Width           =   4095
         End
         Begin VB.CommandButton cmdAutoSpeed 
            Caption         =   "V. Fast"
            Enabled         =   0   'False
            Height          =   375
            Index           =   4
            Left            =   3480
            TabIndex        =   25
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmdAutoSpeed 
            Caption         =   "Fast"
            Enabled         =   0   'False
            Height          =   375
            Index           =   3
            Left            =   2640
            TabIndex        =   24
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmdAutoSpeed 
            Caption         =   "Normal"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   23
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmdAutoSpeed 
            Caption         =   "Slow"
            Enabled         =   0   'False
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   22
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmdAutoSpeed 
            Caption         =   "V. Slow"
            Enabled         =   0   'False
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmdSetSpeed 
            Caption         =   "Set"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3840
            TabIndex        =   20
            Top             =   360
            Width           =   495
         End
         Begin MSComctlLib.Slider sldrSpeed 
            Height          =   255
            Left            =   120
            TabIndex        =   19
            ToolTipText     =   "Speed of the game - 0.1 to 2"
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   3
            Min             =   1
            Max             =   20
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider sldrBulletDamage 
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1560
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   3
            Min             =   10
            Max             =   50
            SelStart        =   10
            TickFrequency   =   5
            Value           =   10
         End
         Begin VB.Label lblMaxScore 
            Caption         =   "Score to win a round:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label lblBulletDamage 
            Caption         =   "Bullet Damage - WWW"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lblSpeedInfo 
            Alignment       =   2  'Center
            Caption         =   "Only the game host can adjust these settings"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   4200
            Width           =   4095
         End
         Begin VB.Label lblSpeed 
            Caption         =   "Speed - WW"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdBot 
      Caption         =   "Bot Controls"
      Height          =   375
      Left            =   3600
      TabIndex        =   42
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClients 
      Caption         =   "Player List"
      Height          =   375
      Left            =   2400
      TabIndex        =   41
      Top             =   7560
      Width           =   975
   End
   Begin VB.Frame fraShips 
      Caption         =   "Ship Type"
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.PictureBox picShipTypes 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1695
         ScaleWidth      =   1695
         TabIndex        =   1
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Star Destroyer"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   8
            Top             =   1440
            Width           =   1695
         End
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Infiltrator"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   7
            Top             =   1200
            Width           =   1695
         End
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Wraith"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   6
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Raptor"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Behemoth"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Hornet"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   4
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Mothership"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   5
            Top             =   720
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraTeam 
      Caption         =   "Team"
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   2640
      TabIndex        =   9
      Top             =   120
      Width           =   2055
      Begin VB.PictureBox picTeam 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   1815
         TabIndex        =   10
         Top             =   240
         Width           =   1815
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Spectator"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   14
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Blue"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   13
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Red"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   12
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Neutral"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   4695
   End
End
Attribute VB_Name = "frmGameOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private DblTeam As Boolean, DblShip As Boolean, DblFrm As Boolean
Private dStartX As Single, dStartY As Single


'Private Const DevHeight As Integer = 2850
'Private Const NormHeight As Integer = 1740

'Private Sub chkBlack_Click()
'modSpaceGame.cg_BlackBG = CBool(chkBlack.Value)
'
'If modSpaceGame.cg_BlackBG Then
'    frmGame.BackColor = vbBlack
'    chkStars.Enabled = True
'    'chkStars.Value = 1
'Else
'    frmGame.BackColor = &H8000000F
'    chkStars.Value = 0
'    chkStars.Enabled = False
'End If
'
'chkStars_Click
'
'End Sub

Private Sub chkBulletCollisions_Click()
modSpaceGame.sv_BulletsCollide = CBool(Me.chkBulletCollisions.Value)
frmGame.SendServerVarsUpdate True
End Sub

Private Sub chkBulletShipVectorAdd_Click()
modSpaceGame.sv_AddBulletVectorToShip = CBool(Me.chkBulletShipVectorAdd.Value)
frmGame.SendServerVarsUpdate True
End Sub

Private Sub chkBulletWalls_Click()
modSpaceGame.sv_BulletWallBounce = CBool(chkBulletWalls.Value)
End Sub

Private Sub chkClipMissiles_Click()
modSpaceGame.sv_ClipMissiles = CBool(chkClipMissiles.Value)
End Sub

'Private Sub chkSound_Click()
'modSpaceGame.Sound = CBool(chkSound.Value)
'End Sub

Private Sub cmdAutoSpeed_Click(Index As Integer)

With sldrSpeed
    Select Case Index
        Case 0
            .Value = 1
        Case 1
            .Value = 5
        Case 2
            .Value = 10
        Case 3
            .Value = 15
        Case 4
            .Value = 20
    End Select
End With

cmdSetSpeed.Default = True

SetFocus2 cmdSetSpeed

End Sub

Private Sub cmdBot_Click()
Unload frmBot
Load frmBot
frmBot.Show vbModeless, frmGame
End Sub

Private Sub cmdClient_Click()
Unload frmGameSettings
Load frmGameSettings
frmGameSettings.Show vbModeless, frmGame
End Sub

Private Sub cmdClients_Click()
Unload frmGameClients
Load frmGameClients
frmGameClients.Show vbModeless, frmGame
End Sub

Private Sub cmdCTFTime_Click()

Dim iTime As Integer
Dim Txt As String

cmdCTFTime.Enabled = False

Txt = txtCTFTime.Text

If LenB(Txt) Then
    On Error GoTo EH
    iTime = val(Txt)
    
    If iTime >= 5 And iTime <= 40 Then
        modSpaceGame.sv_CTFTime = iTime
        frmGame.SendServerVarsUpdate True '.sendCTFTime True
        SetStatus "Time Set"
    Else
        SetStatus "Time must be more than 5 and less than 40"
    End If
End If

Exit Sub
EH:
SetStatus "Error: " & Err.Description
End Sub

Private Sub SetStatus(Txt As String)

lblStatus.Caption = "Status: " & Txt
lblStatus.Refresh

End Sub

Private Sub cmdSetSpeed_Click()
cmdSetSpeed.Enabled = False

modSpaceGame.sv_GameSpeed = sldrSpeed.Value / 10

frmGame.SendGameSpeed True

End Sub

Private Sub Form_Load()

Dim i As Integer

TurnOffToolTip sldrSpeed.hWnd
TurnOffToolTip sldrBulletDamage.hWnd

'DblShip = False
'DblTeam = False
'DblFrm = False
dStartX = -1
dStartY = -1

'Me.Height = IIf(bDevMode, DevHeight, NormHeight)
Me.picServer.BorderStyle = 0
cmdCTFTime.Visible = False
txtCTFTime.Visible = False
'lblStatus.Caption = "St
SetStatus "Loaded Window"

Call Space_FormLoad(Me)

'pos
Me.Top = frmGame.Top + frmGame.height / 2 - Me.height / 2

Me.Left = frmGame.Left + frmGame.width '/ 2 - Me.Width / 2
If (Me.Left + Me.width) > Screen.width Then
    Me.Left = frmGame.Left - Me.width
    'Me.Left = Screen.width - Me.width
End If
If Me.Left < 0 Then
    Me.Left = Screen.width - Me.width 'frmGame.Left + frmGame.width - Me.width
End If

If Me.Top < 0 Then Me.Top = 0
'end pos

modSpaceGame.GameOptionFormLoaded = True

sldrBulletDamage.Value = sv_Bullet_Damage
sldrBulletDamage.Enabled = modSpaceGame.SpaceServer
lblBulletDamage.Enabled = modSpaceGame.SpaceServer
sldrBulletDamage_Click


'chkBlack.Value = IIf(modSpaceGame.cg_BlackBG, 1, 0)
'chkStars.Enabled = modSpaceGame.cg_BlackBG
'chkSound.Value = IIf(modSpaceGame.Sound, 1, 0)

optnShipType(Player(0).ShipType).Value = True
Me.optnTeam(Player(0).Team).Value = True
'chkExplode.Value = IIf(modSpaceGame.BulletsExplode, 1, 0)
Me.chkBulletCollisions.Value = IIf(modSpaceGame.sv_BulletsCollide, 1, 0)
Me.chkBulletShipVectorAdd.Value = IIf(modSpaceGame.sv_AddBulletVectorToShip, 1, 0)
Me.chkClipMissiles.Value = IIf(modSpaceGame.sv_ClipMissiles, 1, 0)
Me.chkBulletWalls.Value = IIf(modSpaceGame.sv_BulletWallBounce, 1, 0)

optnShipType(3).Enabled = frmGame.MotherShipAvail
optnShipType(4).Enabled = frmGame.WraithAvail
optnShipType(5).Enabled = frmGame.InfilAvail
optnShipType(6).Enabled = frmGame.SDAvail

cmdBot.Enabled = modSpaceGame.SpaceServer
fraServer.Enabled = modSpaceGame.SpaceServer
lblSpeedInfo.Visible = Not modSpaceGame.SpaceServer
sldrSpeed.Enabled = modSpaceGame.SpaceServer
lblSpeed.Enabled = modSpaceGame.SpaceServer
Me.chkBulletCollisions.Enabled = modSpaceGame.SpaceServer
Me.chkBulletShipVectorAdd.Enabled = modSpaceGame.SpaceServer
Me.chkClipMissiles.Enabled = modSpaceGame.SpaceServer
Me.chkBulletWalls.Enabled = modSpaceGame.SpaceServer
optnGameType(0).Enabled = modSpaceGame.SpaceServer
optnGameType(1).Enabled = modSpaceGame.SpaceServer
optnGameType(2).Enabled = modSpaceGame.SpaceServer
txtScore.Enabled = modSpaceGame.SpaceServer
lblMaxScore.Enabled = modSpaceGame.SpaceServer

optnGameType(modSpaceGame.sv_GameType).Value = True

sldrSpeed.Value = modSpaceGame.sv_GameSpeed * 10
sldrSpeed_Click

txtScore.Text = CStr(modSpaceGame.sv_ScoreReq)

If modSpaceGame.sv_GameType = CTF Then
    txtCTFTime.Text = CStr(modSpaceGame.sv_CTFTime)
    txtCTFTime.Visible = True
    cmdCTFTime.Visible = True
    cmdCTFTime.Enabled = False
End If

If modSpaceGame.SpaceServer Then
    For i = 0 To 4
        cmdAutoSpeed(i).Enabled = True
    Next i
End If


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

dStartX = X
dStartY = Y

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If dStartX <> -1 Then
    If dStartY <> -1 Then
        If GetDist(dStartX, dStartY, X, Y) > Me.height Then
            optnShipType(eShipTypes.MotherShip).Enabled = True
            optnShipType(eShipTypes.Wraith).Enabled = True
            optnShipType(eShipTypes.Infiltrator).Enabled = True
            optnShipType(eShipTypes.SD).Enabled = True
        End If
    End If
End If

dStartX = -1
dStartY = -1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
modSpaceGame.GameOptionFormLoaded = False
Call Space_FormLoad(Me, True)
End Sub

'-------------------
'Private Sub Form_DblClick()
'DblFrm = True
'Call CheckDev
'End Sub
'
'Private Sub fraShips_DblClick()
'DblShip = True
'Call CheckDev
'End Sub
'
'Private Sub fraTeam_DblClick()
'DblTeam = True
'Call CheckDev
'End Sub

'Private Sub CheckDev()
''If bDevMode And DblTeam And DblShip And DblFrm Then
'    optnShipType(eShipTypes.MotherShip).Enabled = True
'    optnShipType(eShipTypes.Wraith).Enabled = True
'    optnShipType(eShipTypes.Infiltrator).Enabled = True
'    optnShipType(eShipTypes.SD).Enabled = True
'End If
'End Sub
''-------------------

Private Sub optnGameType_Click(Index As Integer)
modSpaceGame.sv_GameType = Index

frmGame.SendGameType True

If Me.Visible Then 'only on actual _Click
    
    txtCTFTime.Visible = False
    cmdCTFTime.Visible = False
    
    Select Case modSpaceGame.sv_GameType
        Case eGameTypes.DM
            frmGame.FlagOwnerID = -1
            'frmGame.AddMainMessage "Game Type - Deathmatch"
            
        Case eGameTypes.CTF
            frmGame.FlagOwnerID = -1
            'frmGame.AddMainMessage "Game Type - CTF"
            
            txtCTFTime.Text = CStr(modSpaceGame.sv_CTFTime)
            txtCTFTime.Visible = True
            cmdCTFTime.Visible = True
            
    End Select
    
End If



End Sub

Private Sub optnShipType_Click(Index As Integer)

frmGame.ActivateShipType Index

If Index = eShipTypes.MotherShip Then
    modSpaceGame.UseAI = False
    If modSpaceGame.GameClientSettingsFormLoaded Then
        frmGameSettings.chkAI.Value = 0
        frmGameSettings.chkAI.Enabled = False
    End If
Else
    If modSpaceGame.GameClientSettingsFormLoaded Then
        frmGameSettings.chkAI.Enabled = True '(frmGame.BotID = 0)
    End If
End If

If (Player(0).State And Player_Secondary) = Player_Secondary Then
    frmGame.SubPlayerState frmGame.MyID, Player_Secondary
End If

End Sub

Private Sub optnTeam_Click(Index As Integer)

'If modSpaceGame.sv_GameType <> Elimination Or frmGame.Playing = False Then
    
    If Index = Player(0).Team Then
        Exit Sub
    End If
    
    If Index = eTeams.Spec Then
        frmGame.SetPlayerState frmGame.MyID, Player_None
        
        Player(0).Speed = 0
        Player(0).X = -frmGame.width - 500
        Player(0).Y = -frmGame.height - 500
        frmGame.AddMainMessage "Use W, A, S and D to spectate"
        
        frmGame.SetCursor False
        
    ElseIf Player(0).Team = Spec Then 'changing FROM spec to something else...
        Player(0).X = Me.ScaleWidth * Rnd()
        Player(0).Y = (Me.ScaleHeight - 500) * Rnd()
        Player(0).Facing = Pi2 * Rnd()
        frmGame.SetPlayerState frmGame.MyID, Player_None
        
        frmGame.SetCursor True
    End If
    
    frmGame.ActivateTeam Index
    
    
    'frmGame.MyTeam = Index
    'modSpaceGame.Player(frmGame.FindPlayer(frmGame.MyID)).Team = Index
'Else
    'optnTeam(Player(0).Team).Value = True
    'SetStatus "You can only pick your team during a round"
'End If
    
    
End Sub

Private Sub sldrBulletDamage_Click()
Const lblCap As String = "Bullet Damage - "
sv_Bullet_Damage = sldrBulletDamage.Value
lblBulletDamage.Caption = lblCap & CStr(sldrBulletDamage.Value)
End Sub

Private Sub sldrBulletDamage_Change()
sldrBulletDamage_Click
End Sub

Private Sub sldrBulletDamage_Scroll()
sldrBulletDamage_Click
End Sub

Private Sub sldrSpeed_Change()
sldrSpeed_Click
cmdSetSpeed.Enabled = True

cmdSetSpeed.Default = True

SetFocus2 cmdSetSpeed

End Sub

Private Sub sldrSpeed_Click()
Const lblCap As String = "Speed - "
lblSpeed.Caption = lblCap & CStr(sldrSpeed.Value / 10)
End Sub

Private Sub sldrSpeed_Scroll()
sldrSpeed_Click
End Sub

Private Sub txtCTFTime_Change()

If Me.Visible Then
    If txtCTFTime.Visible Then
        cmdCTFTime.Enabled = LenB(txtCTFTime.Text)
        cmdCTFTime.Default = cmdCTFTime.Enabled
    End If
End If

End Sub

Private Sub txtCTFTime_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr$(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtScore_Change() 'maxlen = 2

If LenB(txtScore.Text) Then
    If val(txtScore.Text) > 0 Then
        modSpaceGame.sv_ScoreReq = CInt(txtScore.Text)
        
        modDisplay.ShowBalloonTip txtScore, "Score Set", _
            "Score to reach has been set as " & CStr(modSpaceGame.sv_ScoreReq)
    Else
        txtScore.Text = CStr(modSpaceGame.sv_ScoreReq)
    End If
End If

End Sub

Private Sub txtScore_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
End If
End Sub
