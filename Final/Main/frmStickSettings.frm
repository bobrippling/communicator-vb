VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStickGraphics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Settings"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSound 
      Caption         =   "Sound"
      Height          =   1335
      Left            =   120
      TabIndex        =   32
      Top             =   5760
      Width           =   4455
      Begin VB.PictureBox picSound 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   4215
         TabIndex        =   33
         Top             =   240
         Width           =   4215
         Begin MSComctlLib.Slider sldrVol 
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   100
            SmallChange     =   10
            Min             =   -3000
            Max             =   0
            TickFrequency   =   300
         End
         Begin VB.CheckBox chkEnableSound 
            Caption         =   "Enable Sound"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   3015
         End
         Begin VB.Label lblVol 
            Caption         =   "Volume:"
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   360
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraGraphics 
      Caption         =   "Graphics"
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   4455
      Begin VB.PictureBox picGraphics 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   4215
         TabIndex        =   11
         Top             =   360
         Width           =   4215
         Begin VB.CheckBox chkFPS 
            Caption         =   "Show FPS"
            Height          =   255
            Left            =   2400
            TabIndex        =   37
            Top             =   1680
            Width           =   1455
         End
         Begin VB.CheckBox chkSniperScope 
            Caption         =   "Use Sniper Scope"
            Height          =   255
            Left            =   0
            TabIndex        =   31
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox chkWallMarks 
            Caption         =   "Wall Marks"
            Height          =   255
            Left            =   2400
            TabIndex        =   30
            Top             =   720
            Width           =   1335
         End
         Begin VB.CheckBox chkSimple 
            Caption         =   "Draw simple weapons (Two Weapon Mode)"
            Height          =   375
            Left            =   0
            TabIndex        =   21
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdClients 
            Caption         =   "Player List"
            Height          =   375
            Left            =   2640
            TabIndex        =   29
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CheckBox chkInvert 
            Caption         =   "Invert Colours"
            Height          =   255
            Left            =   2400
            TabIndex        =   20
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chHolstered 
            Caption         =   "Holstered Weapon"
            Height          =   255
            Left            =   2400
            TabIndex        =   22
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chkRPGFlame 
            Caption         =   "RPG Flame"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   1455
         End
         Begin VB.PictureBox picBG 
            BackColor       =   &H00808080&
            Height          =   375
            Index           =   1
            Left            =   960
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   26
            Top             =   2160
            Width           =   375
         End
         Begin VB.PictureBox picBG 
            BackColor       =   &H00C0FFC0&
            Height          =   375
            Index           =   2
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   27
            Top             =   2160
            Width           =   375
         End
         Begin VB.PictureBox picBG 
            BackColor       =   &H00FF8080&
            Height          =   375
            Index           =   3
            Left            =   480
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   28
            Top             =   2160
            Width           =   375
         End
         Begin VB.CheckBox chkSmoke 
            Caption         =   "Bullet Smoke"
            Height          =   255
            Left            =   2400
            TabIndex        =   13
            Top             =   0
            Width           =   1575
         End
         Begin VB.CheckBox chkExplosions 
            Caption         =   "Explosions"
            Height          =   255
            Left            =   2040
            TabIndex        =   24
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CheckBox chkBlood 
            Caption         =   "Blood"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkCasing 
            Caption         =   "Bullet Casings"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   1695
         End
         Begin VB.PictureBox picBG 
            BackColor       =   &H00404040&
            Height          =   375
            Index           =   0
            Left            =   1440
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   25
            Top             =   2160
            Width           =   375
         End
         Begin VB.CheckBox chkDead 
            Caption         =   "Dead Sticks"
            Height          =   255
            Left            =   2400
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkMagazines 
            Caption         =   "Magazines"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox chkSparks 
            Caption         =   "Sparks"
            Height          =   255
            Left            =   2400
            TabIndex        =   17
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chkBlackChopper 
            Caption         =   "Black Choppers"
            Height          =   255
            Left            =   2400
            TabIndex        =   19
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label lblColour 
            Caption         =   "Background Colour"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   1800
            Width           =   1455
         End
      End
   End
   Begin VB.Frame fraCol 
      Caption         =   "Stick Settings"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox picCol 
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1875
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         Begin VB.PictureBox picColour 
            Height          =   855
            Left            =   3840
            ScaleHeight     =   795
            ScaleWidth      =   195
            TabIndex        =   6
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   240
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "<Name>"
            Top             =   960
            Width           =   3735
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   9
            Top             =   1440
            Width           =   1695
         End
         Begin MSComctlLib.Slider sldrCol 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   3
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
            TabIndex        =   4
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
            TabIndex        =   5
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
            TabIndex        =   2
            Top             =   0
            Width           =   150
         End
      End
   End
End
Attribute VB_Name = "frmStickGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chHolstered_Click()
modStickGame.cg_HolsteredWeap = CBool(chHolstered.Value)
End Sub

Private Sub chkEnableSound_Click()
modAudio.bDXSoundEnabled = CBool(chkEnableSound.Value)

lblVol.Enabled = modAudio.bDXSoundEnabled
sldrVol.Enabled = modAudio.bDXSoundEnabled
End Sub

Private Sub sldrVol_Change()

On Error Resume Next
frmStickGame.SetDXSoundVol sldrVol.Value

End Sub

Private Sub sldrVol_Click()
sldrVol_Change
End Sub

Private Sub sldrVol_Scroll()
sldrVol_Change
End Sub

'###############################################################

Private Sub chkFPS_Click()
modStickGame.cg_DrawFPS = CBool(chkFPS.Value)
End Sub

Private Sub chkBlackChopper_Click()
modStickGame.cg_ChopperCol = IIf(chkBlackChopper.Value = 1, vbBlack, MSilver)
End Sub

Private Sub chkBlood_Click()
modStickGame.cg_Blood = CBool(chkBlood.Value)
End Sub

Private Sub chkCasing_Click()
modStickGame.cg_Casing = CBool(chkCasing.Value)
End Sub

Private Sub chkDead_Click()
modStickGame.cg_DeadSticks = CBool(chkDead.Value)
End Sub

Private Sub chkExplosions_Click()
modStickGame.cg_Explosions = CBool(chkExplosions.Value)
End Sub

Private Sub chkInvert_Click()

If chkInvert.Value Then
    modStickGame.cg_DisplayMode = modStickGame.cg_DisplayMode_Invert
Else
    modStickGame.cg_DisplayMode = modStickGame.cg_DisplayMode_Normal
End If

End Sub

Private Sub chkMagazines_Click()
modStickGame.cg_Magazines = CBool(chkMagazines.Value)
End Sub

Private Sub chkRPGFlame_Click()
modStickGame.cg_RPGFlame = CBool(chkRPGFlame.Value)
End Sub

Private Sub chkSimple_Click()
modStickGame.cg_SimpleStaticWeapons = CBool(chkSimple.Value)
End Sub

Private Sub chkSmoke_Click()
modStickGame.cg_BulletSmoke = CBool(chkSmoke.Value)
End Sub

Private Sub chkSniperScope_Click()
modStickGame.cl_SniperScope = CBool(chkSniperScope.Value)
End Sub

Private Sub chkSparks_Click()
modStickGame.cg_Sparks = CBool(chkSparks.Value)
End Sub

Private Sub chkWallMarks_Click()
modStickGame.cg_WallMarks = CBool(chkWallMarks.Value)
End Sub

'##############################################################################

Private Sub cmdApply_Click()
Dim NameOp As String, LName As String
Dim PrevName As String

Stick(0).Colour = picColour.BackColor

LName = txtName.Text
NameOp = Trim$(Replace$(LName, "@", vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, "#", vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, modMessaging.MsgEncryptionFlag, vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, ":", vbNullString, , , vbTextCompare)) 'for chat in game
NameOp = Trim$(Replace$(NameOp, modSpaceGame.mPacketSep, vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, modSpaceGame.UpdatePacketSep, vbNullString, , , vbTextCompare))

PrevName = Trim$(Stick(0).Name)
Stick(0).Name = Trim$(NameOp)
txtName.Text = Trim$(Stick(0).Name)

cmdCancel_Click

If Trim$(Stick(0).Name) <> LName Then
    modDisplay.ShowBalloonTip txtName.hWnd, "Tsk tsk", "Certain characters aren't allowed and have been removed"
Else
    modDisplay.ShowBalloonTip txtName.hWnd, "Success", "Name and other stuff have been set"
    
    frmStickGame.SendChatPacket PrevName & " renamed to " & Trim$(Stick(0).Name), Stick(0).Colour
End If

End Sub

Private Sub cmdCancel_Click()
Dim RGBCol As ptRGB

RGBCol = RGBDecode(Stick(0).Colour)

sldrCol(0).Value = RGBCol.Red
sldrCol(1).Value = RGBCol.Green
sldrCol(2).Value = RGBCol.Blue

picColour.BackColor = Stick(0).Colour

txtName.Text = Trim$(Stick(0).Name)

cmdApply.Enabled = False
cmdCancel.Enabled = False

modDisplay.ShowBalloonTip txtName.hWnd, "Reset", "Name and other stuff have been reset"

End Sub

Private Sub cmdClients_Click()
Unload frmStickClients
Load frmStickClients
frmStickClients.Show vbModeless, frmStickGame
End Sub

Private Sub Form_Load()
Dim i As Integer

chkCasing.Value = Abs(modStickGame.cg_Casing)
chkExplosions.Value = Abs(modStickGame.cg_Explosions)
chkBlood.Value = Abs(modStickGame.cg_Blood)
chkRPGFlame.Value = Abs(modStickGame.cg_RPGFlame)
chkFPS.Value = Abs(modStickGame.cg_DrawFPS)
chkDead.Value = Abs(modStickGame.cg_DeadSticks)
chkMagazines.Value = Abs(modStickGame.cg_Magazines)
chkSparks.Value = Abs(modStickGame.cg_Sparks)
chkBlackChopper.Value = Abs(modStickGame.cg_ChopperCol = vbBlack)
chkSmoke.Value = Abs(modStickGame.cg_BulletSmoke)
chkSimple.Value = Abs(modStickGame.cg_SimpleStaticWeapons)
chkWallMarks.Value = Abs(modStickGame.cg_WallMarks)
chkSniperScope.Value = Abs(modStickGame.cl_SniperScope)
chHolstered.Value = Abs(modStickGame.cg_HolsteredWeap)

chkEnableSound.Enabled = modAudio.bDXSoundInited
chkEnableSound.Value = Abs(modAudio.bDXSoundEnabled)
TurnOffToolTip sldrVol.hWnd
If modAudio.bDXSoundInited Then
    On Error Resume Next
    sldrVol.Value = modDXSound.GetVolume(0)
End If

If modStickGame.cg_DisplayMode = modStickGame.cg_DisplayMode_Invert Then
    chkInvert.Value = 1
End If

cmdCancel_Click
lblColourInfo.Caption = "R" & vbNewLine & "G" & vbNewLine & "B"
'For i = sldrCol.LBound To sldrCol.UBound
'    TurnOffToolTip sldrCol(i).hWnd
'Next i
picCol.BorderStyle = 0
picColour.BorderStyle = 0

'pos
Me.Top = frmStickGame.Top + frmStickGame.height / 2 - Me.height / 2

Me.Left = frmStickGame.Left + frmStickGame.width '/ 2 - Me.Width / 2
If (Me.Left + Me.width) > Screen.width Then
    Me.Left = frmStickGame.Left - Me.width
    'Me.Left = Screen.width - Me.width
End If
If Me.Left < 0 Then
    Me.Left = Screen.width - Me.width 'frmGame.Left + frmGame.width - Me.width
End If

If Me.Top < 0 Then Me.Top = 0
'end pos

Call Stick_FormLoad(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Stick_FormLoad(Me, True)
End Sub

Private Sub picBG_Click(Index As Integer)
frmStickGame.picMain.BackColor = picBG(Index).BackColor
frmStickGame.BackColor = picBG(Index).BackColor
modStickGame.cg_BGColour = picBG(Index).BackColor
End Sub

'name + colour stuff
Private Sub sldrCol_Change(Index As Integer)

picColour.BackColor = RGB(sldrCol(0).Value, sldrCol(1).Value, sldrCol(2).Value)

cmdCancel.Enabled = True
cmdApply.Enabled = CBool(LenB(txtName.Text))
End Sub

Private Sub sldrCol_Scroll(Index As Integer)
sldrCol_Change Index
End Sub

Private Sub txtName_Change()
cmdApply.Enabled = CBool(LenB(txtName.Text))
cmdCancel.Enabled = True
End Sub

