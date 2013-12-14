VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStickPerk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Team and Stick Settings"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPerks 
      Caption         =   "Perks"
      ForeColor       =   &H00FF0000&
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   3375
      Begin VB.PictureBox picPerks 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         ScaleHeight     =   3975
         ScaleWidth      =   3135
         TabIndex        =   11
         Top             =   240
         Width           =   3135
         Begin VB.OptionButton optnPerk 
            Caption         =   "Spy - Take another's name and colour"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   25
            Top             =   2280
            Width           =   3135
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   375
            Left            =   600
            TabIndex        =   23
            Top             =   3600
            Width           =   1815
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Focus - Zoom for all weapons"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   24
            Top             =   1800
            Width           =   3135
         End
         Begin VB.PictureBox picSpy 
            Height          =   735
            Left            =   0
            ScaleHeight     =   675
            ScaleWidth      =   3075
            TabIndex        =   20
            Top             =   2760
            Visible         =   0   'False
            Width           =   3135
            Begin VB.ComboBox cboSpyStick 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   360
               Width           =   2775
            End
            Begin VB.Label lblStick 
               Alignment       =   2  'Center
               Caption         =   "Stick to masquerade as..."
               Height          =   255
               Left            =   240
               TabIndex        =   21
               Top             =   0
               Width           =   2415
            End
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Martyrdom - Drop a grenade on death"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   19
            Top             =   2040
            Width           =   3135
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Sniper/Stealth - Name is hidden in prone + sniper rifle training"
            Height          =   375
            Index           =   6
            Left            =   0
            TabIndex        =   18
            Top             =   1440
            Width           =   3135
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Conditioning - Run more quickly"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   17
            Top             =   1200
            Width           =   2655
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Radar Jammer"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   16
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Stopping Power - Higer bullet damage"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   15
            Top             =   720
            Width           =   3015
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Sleight of Hand - Reload more quickly"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   14
            Top             =   480
            Width           =   3015
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Juggernaut - Take less damage"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   13
            Top             =   240
            Width           =   3015
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "No Perk"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   1815
         End
      End
   End
   Begin VB.Frame fraTeam 
      Caption         =   "Team Settings"
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.PictureBox picTeam 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   3135
         TabIndex        =   1
         Top             =   240
         Width           =   3135
         Begin VB.OptionButton optnTeam 
            Caption         =   "Spectator"
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   5
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton optnTeam 
            Caption         =   "Blue"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   4
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optnTeam 
            Caption         =   "Red"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   3
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optnTeam 
            Caption         =   "Neutral"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   1695
         End
         Begin MSComctlLib.Slider sldrSpecSpeed 
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   1080
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   3
            Min             =   5
            Max             =   30
            SelStart        =   10
            TickFrequency   =   5
            Value           =   10
         End
         Begin VB.Label lblSpecSpeed 
            Caption         =   "Spectator Speed - WW"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   720
            Width           =   2175
         End
      End
   End
   Begin VB.CheckBox chkShh 
      Caption         =   "Supressor/Silencer"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CheckBox chkLaserSight 
      Caption         =   "Laser Sight"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "frmStickPerk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NormHeight = 6465
Private Const PerkNormHeight = 3495, picPerkNormHeight = 3135
Private Const ApplyNormTop = 2760

Private Const SpyHeight = 7335
Private Const PerkSpyHeight = 4335, picPerkSpyHeight = 3975
Private Const ApplySpyTop = 3600

Private Sub ShowSpy(Optional bShow As Boolean = True)
Dim i As Integer

If bShow Then
    Me.height = SpyHeight
    fraPerks.height = PerkSpyHeight
    picPerks.height = picPerkSpyHeight
    cmdApply.Top = ApplySpyTop
    
    cboSpyStick.Clear
    For i = 1 To NumSticks - 1
        cboSpyStick.AddItem Trim$(Stick(i).Name)
    Next i
    'If NumSticks >= 2 Then
        'cboSpyStick.Text = Trim$(Stick(1).Name)
        'cmdApply.Enabled = True
    'End If
    
Else
    Me.height = NormHeight
    fraPerks.height = PerkNormHeight
    picPerks.height = picPerkNormHeight
    cmdApply.Top = ApplyNormTop
End If

picSpy.Visible = bShow

End Sub

Private Sub cboSpyStick_Change()
If LenB(cboSpyStick.Text) Then
    cmdApply.Enabled = True
Else
    cmdApply.Enabled = False
End If
End Sub

Private Sub cboSpyStick_Click()
cboSpyStick_Change
End Sub

Private Sub cboSpyStick_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdApply_Click()
Dim i As Integer

cmdApply.Enabled = False

If Stick(0).Perk = pSleightOfHand Then
    For i = 0 To eWeaponTypes.Knife - 1
        modDXSound.SetRelativeFrequency CInt(i + eWeaponTypes.Chopper + 1), Stick(0).sgTimeZone
    Next i
End If

For i = optnPerk.LBound To optnPerk.UBound
    If optnPerk(i).Value Then
        Stick(0).Perk = i
        Exit For
    End If
Next i

If Stick(0).Perk = pSpy Then
    For i = 1 To NumSticks - 1
        If Trim$(Stick(i).Name) = cboSpyStick.Text Then
            Stick(0).MaskID = Stick(i).ID
            Exit For
        End If
    Next i
ElseIf Stick(0).Perk = pSleightOfHand Then
    For i = 0 To eWeaponTypes.Knife - 1
        modDXSound.SetRelativeFrequency CInt(i + eWeaponTypes.Chopper + 1), Stick(0).sgTimeZone * modStickGame.SleightOfHandReloadDecrease
    Next i
End If

End Sub

Private Sub chkLaserSight_Click()

If Stick(0).WeaponType <> Chopper Then
    modStickGame.cg_LaserSight = CBool(chkLaserSight.Value)
Else
    modStickGame.cg_LaserSight = False
    chkLaserSight.Value = 0
    'lblStatus.Caption = "Can't have laser sight on chopper"
End If

End Sub

Private Sub chkShh_Click()

If frmStickGame.WeaponSilencable(Stick(0).WeaponType) Then
    Stick(0).bSilenced = CBool(chkShh.Value)
Else
    'lblStatus.Caption = "Error - Weapon Not Silencable"
    chkShh.Value = 0
End If

End Sub

Private Sub optnPerk_Click(Index As Integer)
If Index = eStickPerks.pSpy Then
    ShowSpy
    cmdApply.Enabled = False
Else
    ShowSpy False
    cmdApply.Enabled = (Index <> Stick(0).Perk)
End If

cmdApply.Default = cmdApply.Enabled

End Sub

Private Sub optnTeam_Click(Index As Integer)
'Dim bMoveCamera As Boolean

If Stick(0).Team = Spec And (Index <> Spec) Then
    frmStickGame.RandomizeMyStickPos
    
ElseIf Me.Visible Then
    If Index = Spec Then
        frmStickGame.AddMainMessage "Use W, A, S and D to spectate"
        modStickGame.cg_sZoom = 1
        'modStickGame.cg_sCamera.X = 0
        'modStickGame.cg_sCamera.Y = 0
        
        'bMoveCamera = True
        'can't do it here - not a spectator yet
        
        Stick(0).X = -10: Stick(0).Y = -10
    End If
End If

Stick(0).Team = Index
If Index = eTeams.Spec Or frmStickGame.StickInGame(0) = False Then
    frmStickGame.SetCursor False
    
    sldrSpecSpeed.Enabled = True
    lblSpecSpeed.Enabled = True
Else
    frmStickGame.SetCursor True
    
    sldrSpecSpeed.Enabled = False
    lblSpecSpeed.Enabled = False
End If

'If bMoveCamera Then
'    'frmStickGame.MoveCameraX modStickGame.cg_sCamera.X
'    'frmStickGame.MoveCameraY modStickGame.cg_sCamera.Y
'    frmStickGame.CentreCameraOnPoint CSng(modStickGame.cg_sCamera.X), CSng(modStickGame.cg_sCamera.Y)
'End If

End Sub

Private Sub sldrSpecSpeed_Change()
sldrSpecSpeed_Click
End Sub

Private Sub sldrSpecSpeed_Click()
modStickGame.cl_SpecSpeed = sldrSpecSpeed.Value / 10
lblSpecSpeed.Caption = "Spectator Speed - " & CStr(modStickGame.cl_SpecSpeed)
End Sub

Private Sub sldrSpecSpeed_Scroll()
sldrSpecSpeed_Click
End Sub

Private Sub Form_Load()

StickTeamFormLoaded = True
picSpy.BorderStyle = 0

LoadStats

'pos
Me.Top = frmStickGame.Top + frmStickGame.height / 2 - Me.height / 2
If Me.Top < 0 Then Me.Top = 0

Me.Left = frmStickGame.Left + frmStickGame.width '/ 2 - Me.Width / 2
If (Me.Left + Me.width) > Screen.width Then
    Me.Left = Screen.width - Me.width - 10
End If
If Me.Left < 0 Then
    Me.Left = Screen.width - Me.width
End If
'end pos

Call Stick_FormLoad(Me)
End Sub

Private Sub LoadStats()
Dim i As Integer

optnPerk(Stick(0).Perk).Value = True

TurnOffToolTip sldrSpecSpeed.hWnd

chkLaserSight.Value = IIf(modStickGame.cg_LaserSight, 1, 0)
chkShh.Value = IIf(Stick(0).bSilenced, 1, 0)
sldrSpecSpeed.Value = modStickGame.cl_SpecSpeed * 10
lblSpecSpeed.Caption = "Spectator Speed - " & CStr(modStickGame.cl_SpecSpeed)
optnTeam(Stick(0).Team).Value = True



'If Stick(0).Team <> Spec Then
'    If frmStickGame.StickInGame(0) = False Then
'        For i = 0 To 3
'            optnTeam(i).Enabled = False
'        Next i
'    End If
'End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
StickTeamFormLoaded = False
Call Stick_FormLoad(Me, True)
End Sub
