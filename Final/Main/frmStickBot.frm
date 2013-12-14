VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStickBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bot Management"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCoOp 
      Caption         =   "Co-Op Settings"
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   4440
      Width           =   4575
      Begin VB.PictureBox picCoOp 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   4335
         TabIndex        =   29
         Top             =   240
         Width           =   4335
         Begin VB.CommandButton cmdCoOp 
            Caption         =   "Set up Co-Op Bots"
            Height          =   375
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Width           =   1935
         End
         Begin VB.CheckBox chkChopper 
            Alignment       =   1  'Right Justify
            Caption         =   "Last Bot is a Helicopter"
            Height          =   255
            Left            =   2280
            TabIndex        =   33
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtNBots 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3840
            MaxLength       =   2
            TabIndex        =   31
            Text            =   "10"
            Top             =   0
            Width           =   495
         End
         Begin VB.Label lblNBots 
            Caption         =   "Bots to Add:"
            Height          =   255
            Left            =   2280
            TabIndex        =   30
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton cmdChallenge 
      Caption         =   "Bot Challenge Set Up"
      Height          =   375
      Left            =   2520
      TabIndex        =   37
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "Remove All Bots"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   39
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Frame fraOther 
      Caption         =   "Weapon + AI Settings"
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   4575
      Begin VB.PictureBox picOther 
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1875
         ScaleWidth      =   4335
         TabIndex        =   13
         Top             =   240
         Width           =   4400
         Begin VB.CheckBox chkZombie 
            Alignment       =   1  'Right Justify
            Caption         =   "Zombie"
            Height          =   255
            Left            =   2280
            TabIndex        =   27
            Top             =   1680
            Width           =   2000
         End
         Begin VB.CheckBox chkStickChat 
            Alignment       =   1  'Right Justify
            Caption         =   "Sticks can Chat"
            Height          =   255
            Left            =   2280
            TabIndex        =   26
            Top             =   1440
            Width           =   2000
         End
         Begin VB.CheckBox chkMines 
            Alignment       =   1  'Right Justify
            Caption         =   "AI use mines"
            Height          =   255
            Left            =   2280
            TabIndex        =   20
            Top             =   720
            Width           =   1995
         End
         Begin VB.ComboBox cboFireMode 
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtRotationRate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   25
            Text            =   "9"
            Top             =   1320
            Width           =   450
         End
         Begin VB.CheckBox chkSnipers 
            Alignment       =   1  'Right Justify
            Caption         =   "Don't Add Snipers"
            Height          =   255
            Left            =   2280
            TabIndex        =   23
            Top             =   1200
            Width           =   2000
         End
         Begin VB.CheckBox chkChopperRocket 
            Alignment       =   1  'Right Justify
            Caption         =   "Chopper can rocket"
            Height          =   255
            Left            =   2280
            TabIndex        =   22
            Top             =   960
            Width           =   2000
         End
         Begin VB.CheckBox chkShh 
            Alignment       =   1  'Right Justify
            Caption         =   "Silence <WEAPON>"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   720
            Width           =   2000
         End
         Begin VB.CheckBox chkMartyrdom 
            Alignment       =   1  'Right Justify
            Caption         =   "Bot has Martyrdom Perk"
            Height          =   255
            Left            =   0
            TabIndex        =   21
            Top             =   960
            Width           =   2000
         End
         Begin VB.CheckBox chkFlashBang 
            Alignment       =   1  'Right Justify
            Caption         =   "AI use flashbangs"
            Height          =   255
            Left            =   2280
            TabIndex        =   18
            Top             =   480
            Width           =   1995
         End
         Begin VB.CheckBox chkAIShoot 
            Alignment       =   1  'Right Justify
            Caption         =   "AI can shoot"
            Height          =   255
            Left            =   2280
            TabIndex        =   16
            Top             =   240
            Width           =   2000
         End
         Begin VB.CheckBox chkAIMove 
            Alignment       =   1  'Right Justify
            Caption         =   "AI can move"
            Height          =   255
            Left            =   2280
            TabIndex        =   15
            Top             =   0
            Width           =   2000
         End
         Begin VB.ComboBox cboWeapon 
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label lblRotation 
            Caption         =   "Rotation Rate:"
            Height          =   255
            Left            =   15
            TabIndex        =   24
            Top             =   1320
            Width           =   1335
         End
      End
   End
   Begin projMulti.ScrollListBox lstBots 
      Height          =   1215
      Left            =   120
      TabIndex        =   35
      Top             =   5880
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2143
   End
   Begin VB.CommandButton cmdAddBot 
      Caption         =   "Add Bot"
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   7200
      Width           =   2175
   End
   Begin VB.CommandButton cmdRemoveBot 
      Caption         =   "Remove Bot"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Frame fraTeam 
      Caption         =   "Team + Colour"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.PictureBox picCol 
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   4155
         TabIndex        =   6
         Top             =   840
         Width           =   4215
         Begin VB.PictureBox picBotColour 
            Height          =   855
            Left            =   3840
            ScaleHeight     =   795
            ScaleWidth      =   195
            TabIndex        =   11
            Top             =   0
            Width           =   255
         End
         Begin MSComctlLib.Slider sldrCol 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   8
            Top             =   0
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   450
            _Version        =   393216
            SmallChange     =   5
            Max             =   255
            SelStart        =   100
            TickFrequency   =   10
            Value           =   100
         End
         Begin MSComctlLib.Slider sldrCol 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   450
            _Version        =   393216
            SmallChange     =   5
            Max             =   255
            SelStart        =   100
            TickFrequency   =   10
            Value           =   100
         End
         Begin MSComctlLib.Slider sldrCol 
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   480
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   450
            _Version        =   393216
            SmallChange     =   5
            Max             =   255
            SelStart        =   100
            TickFrequency   =   10
            Value           =   100
         End
         Begin VB.Label lblColourInfo 
            AutoSize        =   -1  'True
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   7
            Top             =   0
            Width           =   150
         End
      End
      Begin VB.PictureBox picTeam 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   4275
         TabIndex        =   1
         Top             =   240
         Width           =   4335
         Begin VB.CheckBox chkRandomCol 
            Alignment       =   1  'Right Justify
            Caption         =   "Random Colour"
            Height          =   255
            Left            =   2520
            TabIndex        =   5
            Top             =   240
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Neutral"
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   2
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Red"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Blue"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   4
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   5520
      Width           =   4575
   End
End
Attribute VB_Name = "frmStickBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const random_Weapon = "Random", rifleMan_Weapon = "Rifleman", _
                    weapon_Sep = "-----"

Private Const MaxBots = 21, _
    FireMode_Default = "Default", _
    FireMode_Random = random_Weapon


Private Function GetFireMode() As eFireModes
GetFireMode = frmStickGame.FireModeNameToInt(cboFireMode.Text)
End Function

Private Sub cboWeapon_Change()
Dim sWeapon As String
'Dim vWeapon As eWeaponTypes

sWeapon = cboWeapon.Text

If sWeapon = weapon_Sep Then
    cboWeapon.ListIndex = cboWeapon.ListIndex - 1
Else
    'vWeapon = WeaponNameToInt(sWeapon)
    
    With chkShh
        If sWeapon = random_Weapon Or sWeapon = rifleMan_Weapon Then
            .Caption = "Silenced"
            .Enabled = True
        Else
            .Caption = "Silenced " & sWeapon
            .Enabled = frmStickGame.WeaponSilencable(WeaponNameToInt(sWeapon))
            If .Enabled = False Then .Value = 0
        End If
    End With
End If

RefreshFireModeList

End Sub
Private Sub cboWeapon_Click()
cboWeapon_Change
End Sub
Private Sub cboWeapon_LostFocus()
cboWeapon_Change
End Sub
Private Sub cboWeapon_Scroll()
cboWeapon_Change
End Sub

Private Sub cboWeapon_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub chkAIMove_Click()
modStickGame.sv_AIMove = CBool(chkAIMove.Value)
End Sub

Private Sub chkChopperRocket_Click()
modStickGame.sv_AIHeliRocket = CBool(chkChopperRocket.Value)
End Sub

Private Sub chkAIShoot_Click()
modStickGame.sv_AIShoot = CBool(chkAIShoot.Value)
End Sub

Private Sub chkFlashBang_Click()
Dim i As Integer

modStickGame.sv_AIUseFlashBangs = CBool(chkFlashBang.Value)

If modStickGame.sv_AIUseFlashBangs = False Then
    For i = 0 To NumSticks - 1
        If Stick(i).IsBot Then
            Stick(i).iNadeType = nFrag
        End If
    Next i
End If

End Sub

Private Sub chkMines_Click()
modStickGame.sv_AIMine = CBool(chkMines.Value)
End Sub

Private Sub chkStickChat_Click()
modStickGame.cl_StickBotChat = CBool(chkStickChat.Value)
End Sub

Private Sub chkZombie_Click()
If chkZombie.Value = 1 Then
    optnTeam(1).Value = True
End If
End Sub

'Private Sub cboWeapon_KeyDown(KeyCode As Integer, Shift As Integer)
'cboWeapon_Change
'End Sub
'Private Sub cboWeapon_Change()
'chkShh_Click
'End Sub
'Private Sub cboWeapon_Scroll()
'cboWeapon_Change
'End Sub
'Private Sub chkShh_Click()
'Dim sWeapon As String
'Dim i As Integer
'
'If chkShh.Value = 1 Then
'    sWeapon = cboWeapon.Text
'    If sWeapon <> random_Weapon Then
'        For i = 0 To eWeaponTypes.Chopper
'            If GetWeaponName(CInt(i)) = sWeapon Then
'                chkShh.Value = Abs(frmStickGame.WeaponSilencable(CInt(i)))
'                Exit Sub 'SUB
'            End If
'        Next i
'
'        SetStatus "Error - Weapon not silencable", True
'
'    End If
'End If
'
'End Sub

Private Function GetRandomRifle() As eWeaponTypes

'weapons: AK = 0, XM8 = 1, SAR = 2, G3  = 3

GetRandomRifle = IntRand(0, 3)

End Function

Private Sub cmdAddBot_Click()
Dim i As Integer
Dim vWeapon As eWeaponTypes, sWeapon As String
Dim vTeam As eTeams
Dim nBots As Integer
Dim vFireMode As eFireModes

For i = 0 To NumSticks - 1
    If Stick(i).IsBot Then
        nBots = nBots + 1
    End If
Next i


If nBots > MaxBots Then
    SetStatus "Can't have more than " & CStr(MaxBots) & " bots", True
Else
    vTeam = -1
    vWeapon = -1
    
    For i = 0 To optnTeam.UBound
        If optnTeam(i).Value Then
            vTeam = i
            Exit For
        End If
    Next i
    
    
    
    sWeapon = cboWeapon.Text
    If sWeapon = random_Weapon Then
        vWeapon = frmStickGame.GetRandomStaticWeapon()
        
        If chkSnipers.Value = 1 Then
            Do While frmStickGame.WeaponIsSniper(vWeapon)
                vWeapon = frmStickGame.GetRandomStaticWeapon()
            Loop
        End If
        
    ElseIf sWeapon = rifleMan_Weapon Then
        vWeapon = GetRandomRifle()
        
    Else
        vWeapon = WeaponNameToInt(sWeapon)
    End If
    
    
    If vWeapon = -1 Then
        SetStatus "Select a Weapon", True
    Else
        
        
        i = frmStickGame.AddBot(vWeapon, vTeam, GetColourToUse())
        If chkMartyrdom.Value Then
            Stick(i).Perk = pMartyrdom
        ElseIf Stick(i).WeaponType = RPG Then
            Stick(i).Perk = pSleightOfHand
        End If
        SetSilenced i
        
        
        If cboFireMode.Text = FireMode_Default Then
            frmStickGame.Make_Weapon_Default_FireMode i
            
        ElseIf cboFireMode.Text = FireMode_Random Then
            
            Do
                vFireMode = CInt(Rnd() * eFireModes.Single_Shot)
                'prevent single shot [unless selected on purpose]
            Loop While vFireMode = Single_Shot
            
            If frmStickGame.WeaponSupportsFireMode(vWeapon, vFireMode) Then
                frmStickGame.SetFireMode i, vFireMode
            Else
                frmStickGame.Make_Weapon_Default_FireMode i
            End If
            
        Else
            frmStickGame.SetFireMode i, GetFireMode()
        End If
        
        
        If chkZombie.Value = 1 Then
            frmStickGame.MakeZombie i
        End If
        
        
        frmStickGame.SendChatPacketBroadcast "Bot Added: " & Trim$(Stick(i).Name), Stick(i).Colour
        
        SetStatus "Bot Added"
        
        RefreshBotList
        lstBots.ListIndex = 0 'top
    End If
    
    cmdRemoveAll.Enabled = True
End If

End Sub

Private Sub SetSilenced(i As Integer)
If chkShh.Value = 1 Or Rnd() > 0.7 Then
    Stick(i).bSilenced = frmStickGame.WeaponSilencable(Stick(i).WeaponType)
End If
End Sub

Private Sub cmdChallenge_Click()
Dim i As Integer
Dim vTeam As eTeams

'remove any bots
For i = 0 To NumSticks - 1
    If Stick(i).IsBot Then
        cmdRemoveAll_Click
        Exit For
    End If
Next i

For i = optnTeam.LBound To optnTeam.UBound
    If optnTeam(i).Value Then
        vTeam = i
        Exit For
    End If
Next i

If chkSnipers.Value = 1 Then
    frmStickGame.AddBot W1200, vTeam, GetColourToUse()
Else
    frmStickGame.AddBot AWM, vTeam, GetColourToUse()
End If
frmStickGame.AddBot AK, vTeam, GetColourToUse() 'different colour for each
frmStickGame.AddBot AUG, vTeam, GetColourToUse()
frmStickGame.AddBot Mac10, vTeam, GetColourToUse()
frmStickGame.AddBot G3, vTeam, GetColourToUse()


If chkMartyrdom.Value = 1 Then
    For i = 1 To NumSticks - 1
        If Stick(i).IsBot Then
            Stick(i).Perk = pMartyrdom
        End If
    Next i
End If
If chkShh.Value = 1 Then
    For i = 1 To NumSticks - 1
        If Stick(i).IsBot Then
            SetSilenced i
        End If
    Next i
End If
If chkZombie.Value = 1 Then
    For i = 1 To NumSticks - 1
        If Stick(i).IsBot Then
            frmStickGame.MakeZombie i
        End If
    Next i
End If


RefreshBotList

End Sub

Private Function GetColourToUse() As Long
GetColourToUse = IIf(CBool(chkRandomCol.Value), modSpaceGame.RandomRGBColour(), picBotColour.BackColor)
End Function

Private Sub cmdCoOp_Click()
Dim i As Integer, iMax As Integer, j As Integer

'remove any bots
For i = 0 To NumSticks - 1
    If Stick(i).IsBot Then
        cmdRemoveAll_Click
        Exit For
    End If
Next i

If LenB(txtNBots.Text) Then
    iMax = val(txtNBots.Text)
Else
    iMax = 10
End If
If iMax > MaxBots Then
    iMax = MaxBots
    txtNBots.Text = CStr(MaxBots)
End If


For i = 1 To iMax
    j = frmStickGame.AddBot(frmStickGame.GetRandomStaticWeapon(), Red, GetColourToUse())
    
    If chkMartyrdom.Value Then
        Stick(j).Perk = pMartyrdom
    ElseIf Stick(j).WeaponType = RPG Then
        Stick(j).Perk = pSleightOfHand
    End If
    
    If chkSnipers.Value = 1 Then
        Do While frmStickGame.WeaponIsSniper(Stick(j).WeaponType)
            'Stick(j).WeaponType = frmStickGame.GetRandomStaticWeapon()
            frmStickGame.SetSticksWeapon j, frmStickGame.GetRandomStaticWeapon()
        Loop
    End If
    
    SetSilenced j
    
    frmStickGame.RandomizeCoOpBot j
Next i

'bot 4 - shotty
For i = NumSticks - 1 To 0 Step -1
    If Trim$(Stick(i).Name) = "Bot 4" Then
        If Stick(i).WeaponType <> W1200 Then
            Stick(i).CurrentWeapons(1) = Stick(i).WeaponType 'holster current
            'Stick(i).WeaponType = W1200
            frmStickGame.SetSticksWeapon i, W1200
            Stick(i).CurrentWeapons(2) = W1200
            Exit For
        End If
    End If
Next i

If chkChopper.Value = 1 Then
    With Stick(j)
        '.WeaponType = Chopper
        frmStickGame.SetSticksWeapon j, Chopper
        .X = StickGameWidth - Rnd() * 1000
        .Y = 1000
        .bSilenced = False
    End With
End If

If chkZombie.Value = 1 Then
    For i = 1 To NumSticks - 1
        If Stick(i).IsBot Then
            frmStickGame.MakeZombie i
        End If
    Next i
End If

'move myself
frmStickGame.MoveStickToCoOpStart 0

RefreshBotList

modStickGame.sv_GameType = gCoOp
frmStickGame.SendServerVarPacket True
Call frmStickGame.GameTypeChanged
cmdRemoveAll.Enabled = True

End Sub

Private Sub cmdRemoveAll_Click()
Dim i As Integer

cmdRemoveAll.Enabled = False

Do While i < modStickGame.NumSticks
    If Stick(i).IsBot Then
        frmStickGame.SendBroadcast sExits & CStr(Stick(i).ID)
        
        frmStickGame.RemoveStick i
        i = i - 1
    End If
    i = i + 1
Loop


modStickGame.sv_GameType = gDeathMatch
For i = 0 To NumSticks - 1
    Stick(i).bAlive = True
Next i
frmStickGame.SendServerVarPacket True
frmStickGame.GameTypeChanged

RefreshBotList

frmStickGame.SendChatPacketBroadcast "All Bots: Removed", Stick(0).Colour

End Sub

Private Sub cmdRemoveBot_Click()
Dim i As Integer, BotID As Integer, BotI As Integer
Dim Txt As String

Txt = Trim$(lstBots.Text)

BotID = -1
For i = 0 To modStickGame.NumSticks - 1
    If Trim$(Stick(i).Name) = Txt Then
        BotID = Stick(i).ID
        BotI = i
        Exit For
    End If
Next i

cmdRemoveBot.Enabled = False

If BotID <> -1 Then
    frmStickGame.RemoveBot BotI
    
    SetStatus "Removed Bot"
    
    
    lstBots.ListIndex = IIf(lstBots.ListCount > 0, 0, -1)
    lstBots_Click
Else
    SetStatus "Bot Not Found", True
End If

RefreshBotList

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()

Dim RGBCol As ptRGB
Dim i As eWeaponTypes
Dim j As Integer

picTeam.BorderStyle = 0
picBotColour.BorderStyle = 0
picCol.BorderStyle = 0
picOther.BorderStyle = 0
lblColourInfo.Caption = "R" & vbNewLine & "G" & vbNewLine & "B"
chkChopper.Value = 1

RGBCol = RGBDecode(vbYellow)

sldrCol(0).Value = RGBCol.Red
sldrCol(1).Value = RGBCol.Green
sldrCol(2).Value = RGBCol.Blue

chkAIMove.Value = Abs(modStickGame.sv_AIMove)
chkAIShoot.Value = Abs(modStickGame.sv_AIShoot)
chkChopperRocket.Value = Abs(modStickGame.sv_AIHeliRocket)
chkFlashBang.Value = Abs(modStickGame.sv_AIUseFlashBangs)
chkMines.Value = Abs(modStickGame.sv_AIMine)
chkStickChat.Value = Abs(modStickGame.cl_StickBotChat)

txtRotationRate.Text = modStickGame.sv_AI_Rotation_Rate * 180 / Pi

cboWeapon.AddItem random_Weapon
cboWeapon.AddItem rifleMan_Weapon
cboWeapon.AddItem weapon_Sep
For i = 0 To eWeaponTypes.Chopper
    If i <> Knife Then
        cboWeapon.AddItem GetWeaponName(i)
    End If
Next i
cboWeapon.ListIndex = 0

RefreshFireModeList

For j = 0 To NumSticks - 1
    If Stick(j).IsBot Then
        cmdRemoveAll.Enabled = True
        Exit For
    End If
Next j


'pos
Me.Top = frmStickOptions.Top + frmStickOptions.height / 2 - Me.height / 2
Me.Left = frmStickOptions.Left + frmStickOptions.width / 2
If (Me.Left + Me.width) > Screen.width - 10 Then
    Me.Left = Screen.width - Me.width - 10
End If

RefreshBotList

Call Stick_FormLoad(Me)
'end pos

SetStatus "Loaded Window"

End Sub
Private Sub RefreshFireModeList()
Dim i As Integer
Dim vWeapon As eWeaponTypes

Dim sCurrent As String

sCurrent = cboFireMode.Text

cboFireMode.Clear
cboFireMode.AddItem FireMode_Random
cboFireMode.AddItem FireMode_Default

vWeapon = WeaponNameToInt(cboWeapon.Text)

If vWeapon > -1 Then
    For i = 0 To eFireModes.Single_Shot
        If frmStickGame.WeaponSupportsFireMode(vWeapon, CInt(i)) Then
            cboFireMode.AddItem frmStickGame.GetFireModeName(CInt(i))
        End If
    Next i
End If


For i = 0 To cboFireMode.ListCount - 1
    If cboFireMode.List(i) = sCurrent Then
        cboFireMode.ListIndex = i
        Exit For
    End If
Next i

If i = cboFireMode.ListCount Then cboFireMode.ListIndex = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Stick_FormLoad(Me, True)
If modStickGame.StickOptionFormLoaded Then
    SetFocus2 frmStickOptions
End If
End Sub

Private Sub lstBots_Click()
cmdRemoveBot.Enabled = (Len(lstBots.Text) > 0)
End Sub

Private Sub optnTeam_Click(Index As Integer)

Select Case Index
    Case eTeams.Blue
        sldrCol(0).Value = 0
        sldrCol(1).Value = 0
        sldrCol(2).Value = 255
    Case eTeams.Red
        sldrCol(0).Value = 255
        sldrCol(1).Value = 0
        sldrCol(2).Value = 0
    Case eTeams.Neutral
        sldrCol(0).Value = 255
        sldrCol(1).Value = 255
        sldrCol(2).Value = 0
End Select

End Sub

Private Sub sldrCol_Change(Index As Integer)
sldrCol_Click Index
End Sub

Private Sub sldrCol_Click(Index As Integer)
picBotColour.BackColor = RGB(sldrCol(0).Value, sldrCol(1).Value, sldrCol(2).Value)
End Sub

Private Sub sldrCol_Scroll(Index As Integer)
sldrCol_Click Index
End Sub

Private Sub SetStatus(ByVal T As String, Optional ByVal Red As Boolean = False)
Const K As String = "Status: "
If Red Then
    lblStatus.Caption = T
Else
    lblStatus.Caption = K & T
End If

If Red Then
    lblStatus.ForeColor = vbRed
Else
    lblStatus.ForeColor = &HFF0000
End If
lblStatus.Refresh
End Sub

Private Sub RefreshBotList()
Dim i As Integer
Dim ListText As String

ListText = lstBots.Text
lstBots.Clear

For i = 0 To modStickGame.NumSticks - 1
    If Stick(i).IsBot Then
        lstBots.AddItem Trim$(Stick(i).Name)
    End If
Next i

For i = 0 To lstBots.ListCount - 1
    If lstBots.List(i) = ListText Then
        On Error Resume Next
        lstBots.ListIndex = i
        Exit For
    End If
Next i

If i = lstBots.ListCount Then
    If i > 0 Then
        'none selected, select top
        lstBots.ListIndex = 0
    End If
End If

If lstBots.ListCount > 0 Then
    cmdRemoveAll.Enabled = True
Else
    cmdRemoveAll.Enabled = False
End If

End Sub

Private Sub txtNBots_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Not IsNumeric(Chr$(KeyAscii)) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtRotationRate_Change()
Dim ang As Single
Dim Txt As String

Txt = txtRotationRate.Text

If LenB(Txt) Then
    ang = val(Txt)
    
    If ang > 0 Then
        If ang < 100 Then
            SetStatus "Rotation Rate Set as " & CStr(ang)
            modStickGame.UpdateBotRotationRate ang * Pi / 180
        Else
            SetStatus "Rotation Rate must be less than 100", True
        End If
    Else
        SetStatus "Rotation Rate must be greater than 0", True
    End If
    
End If

End Sub

Private Sub txtRotationRate_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Not IsNumeric(Chr$(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub
