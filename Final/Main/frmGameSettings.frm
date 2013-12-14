VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGameSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game/Graphics Settings"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCol 
      Caption         =   "Player Settings"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   120
      TabIndex        =   24
      Top             =   5040
      Width           =   4455
      Begin VB.PictureBox picCol 
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1875
         ScaleWidth      =   4155
         TabIndex        =   25
         Top             =   240
         Width           =   4215
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   33
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   32
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   240
            MaxLength       =   20
            TabIndex        =   31
            Text            =   "<Name>"
            Top             =   960
            Width           =   3735
         End
         Begin VB.PictureBox picColour 
            Height          =   855
            Left            =   3840
            ScaleHeight     =   795
            ScaleWidth      =   195
            TabIndex        =   30
            Top             =   0
            Width           =   255
         End
         Begin MSComctlLib.Slider sldrCol 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   27
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
            TabIndex        =   28
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
            TabIndex        =   29
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
            TabIndex        =   26
            Top             =   0
            Width           =   150
         End
      End
   End
   Begin VB.CommandButton cmdCrosshair 
      Caption         =   "Crosshair Options"
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   7560
      Width           =   4455
   End
   Begin VB.Frame fraGraphics 
      Caption         =   "Graphics"
      ForeColor       =   &H00FF0000&
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4455
      Begin VB.PictureBox picGraphics 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   4095
         TabIndex        =   7
         Top             =   360
         Width           =   4095
         Begin VB.CheckBox chkInvert 
            Alignment       =   1  'Right Justify
            Caption         =   "Invert Colours"
            Height          =   255
            Left            =   2280
            TabIndex        =   17
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chk3DStars 
            Alignment       =   1  'Right Justify
            Caption         =   "3D Stars"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton cmdLow 
            Caption         =   "Low Graphics"
            Height          =   375
            Left            =   0
            TabIndex        =   22
            Top             =   2520
            Width           =   1935
         End
         Begin VB.CommandButton cmdHigh 
            Caption         =   "High Graphics"
            Height          =   375
            Left            =   2160
            TabIndex        =   23
            Top             =   2520
            Width           =   1935
         End
         Begin VB.CheckBox chkMissileLock 
            Alignment       =   1  'Right Justify
            Caption         =   "Show Missile Lock"
            Height          =   255
            Left            =   2280
            TabIndex        =   15
            Top             =   720
            Width           =   1815
         End
         Begin VB.CheckBox chkMap 
            Alignment       =   1  'Right Justify
            Caption         =   "Draw Map"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chkBulletSmoke 
            Alignment       =   1  'Right Justify
            Caption         =   "Draw Bullet Smoke"
            Height          =   255
            Left            =   2280
            TabIndex        =   13
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox chkDrawThick 
            Alignment       =   1  'Right Justify
            Caption         =   "Draw Thick"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1815
         End
         Begin VB.CheckBox chkSmoke 
            Alignment       =   1  'Right Justify
            Caption         =   "Draw Smoke"
            Height          =   255
            Left            =   2280
            TabIndex        =   9
            Top             =   0
            Width           =   1815
         End
         Begin VB.CheckBox chkNoCls 
            Alignment       =   1  'Right Justify
            Caption         =   "Crazy Mode"
            Height          =   255
            Left            =   2280
            TabIndex        =   19
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox chkExplosions 
            Alignment       =   1  'Right Justify
            Caption         =   "Draw Explosions"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   720
            Width           =   1815
         End
         Begin VB.CheckBox chkFPS 
            Alignment       =   1  'Right Justify
            Caption         =   "Show FPS"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox chkStars 
            Alignment       =   1  'Right Justify
            Caption         =   "Star Background"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox chkGunSmoke 
            Alignment       =   1  'Right Justify
            Caption         =   "Draw Gun Smoke"
            Height          =   255
            Left            =   2280
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
         Begin MSComctlLib.Slider sldrMapLen 
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   1800
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   450
            _Version        =   393216
            Min             =   6
            Max             =   30
            SelStart        =   6
            TickFrequency   =   4
            Value           =   6
         End
         Begin VB.Label lblMapLen 
            Caption         =   "Map Size - WWW"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   1800
            Width           =   1575
         End
      End
   End
   Begin VB.Frame fraClient 
      Caption         =   "Other"
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox picOther 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   4215
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         Begin VB.CheckBox chkMouse 
            Alignment       =   1  'Right Justify
            Caption         =   "Use Mouse Aiming"
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   1815
         End
         Begin VB.CheckBox chkAI 
            Alignment       =   1  'Right Justify
            Caption         =   "AI"
            Height          =   255
            Left            =   2280
            TabIndex        =   3
            Top             =   0
            Width           =   1815
         End
         Begin MSComctlLib.Slider sldrSens 
            Height          =   255
            Left            =   1680
            TabIndex        =   5
            Top             =   600
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   3
            Min             =   5
            Max             =   20
            SelStart        =   5
            TickFrequency   =   3
            Value           =   5
         End
         Begin VB.Label lblSens 
            Caption         =   "Turning Speed - WW"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   600
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "frmGameSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chk3DStars_Click()
modSpaceGame.cg_Stars3D = CBool(chk3DStars.Value)
End Sub

Private Sub chkBulletSmoke_Click()
modSpaceGame.cg_BulletSmoke = CBool(chkBulletSmoke.Value)
End Sub

Private Sub chkGunSmoke_Click()
modSpaceGame.cg_GunSmoke = CBool(chkGunSmoke.Value)
End Sub

Private Sub chkInvert_Click()
modSpaceGame.cg_SpaceDisplayMode = IIf(chkInvert.Value, _
    modStickGame.cg_DisplayMode_Invert, modStickGame.cg_DisplayMode_Normal)
End Sub

Private Sub chkMap_Click()
modSpaceGame.cg_ShowMap = CBool(chkMap.Value)

sldrMapLen.Enabled = modSpaceGame.cg_ShowMap
lblMapLen.Enabled = modSpaceGame.cg_ShowMap
End Sub

Private Sub chkMissileLock_Click()
modSpaceGame.cg_ShowMissileLock = CBool(chkMissileLock.Value)
End Sub

Private Sub cmdApply_Click()
Dim NameOp As String
Dim LName As String

Player(0).Colour = picColour.BackColor

LName = txtName.Text
NameOp = Trim$(Replace$(LName, "@", vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, "#", vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, modMessaging.MsgEncryptionFlag, vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, ":", vbNullString, , , vbTextCompare)) 'for chat in game
NameOp = Trim$(Replace$(NameOp, modSpaceGame.mPacketSep, vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, modSpaceGame.UpdatePacketSep, vbNullString, , , vbTextCompare))
Player(0).Name = Trim$(NameOp)

txtName.Text = Trim$(Player(0).Name)

If Trim$(Player(0).Name) <> LName Then
    modDisplay.ShowBalloonTip txtName, "Tsk tsk", "Certain characters aren't allowed and have been removed"
Else
    modDisplay.ShowBalloonTip txtName, "Success", "Name and other stuff have been set"
End If

cmdCancel_Click

End Sub

Private Sub cmdCancel_Click()
Dim RGBCol As ptRGB

RGBCol = RGBDecode(Player(0).Colour)

sldrCol(0).Value = RGBCol.Red
sldrCol(1).Value = RGBCol.Green
sldrCol(2).Value = RGBCol.Blue

picColour.BackColor = Player(0).Colour

txtName.Text = Trim$(Player(0).Name)

cmdApply.Enabled = False
cmdCancel.Enabled = False

modDisplay.ShowBalloonTip txtName, "Reset", "Name and other stuff have been reset"

End Sub

Private Sub cmdHigh_Click()
Call GraphicChecks(1)
End Sub

Private Sub cmdLow_Click()
Call GraphicChecks(0)
End Sub

Private Sub GraphicChecks(iVal As Integer)

Me.chkExplosions.Value = iVal
Me.chkStars.Value = iVal
Me.chkSmoke.Value = iVal
Me.chkGunSmoke.Value = iVal
Me.chkBulletSmoke.Value = iVal
Me.chkMap.Value = iVal

End Sub

Private Sub Form_Load()
Dim St As eShipTypes

modSpaceGame.GameClientSettingsFormLoaded = True


TurnOffToolTip sldrSens.hWnd
TurnOffToolTip sldrMapLen.hWnd

'##########
'settings

sldrSens.Value = frmGame.ROTATION_RATE
sldrSens.Enabled = Not modSpaceGame.cl_UseMouse
lblSens.Enabled = sldrSens.Enabled
sldrSens_Click

chkMouse.Value = IIf(modSpaceGame.cl_UseMouse, 1, 0)
chkMouse_Click

chkNoCls.Value = IIf(modSpaceGame.cg_Cls, 0, 1)
Me.chkExplosions.Value = IIf(modSpaceGame.cg_DrawExplosions, 1, 0)
chkDrawThick.Value = IIf(modSpaceGame.cg_DrawThick, 1, 0)
chkAI.Value = IIf(modSpaceGame.UseAI, 1, 0)
chkStars.Value = IIf(modSpaceGame.cg_StarBG, 1, 0)
chkFPS.Value = IIf(modSpaceGame.cg_ShowFPS, 1, 0)
Me.chkSmoke.Value = IIf(modSpaceGame.cg_Smoke, 1, 0)
Me.chkGunSmoke.Value = IIf(modSpaceGame.cg_GunSmoke, 1, 0)
Me.chkBulletSmoke.Value = IIf(modSpaceGame.cg_BulletSmoke, 1, 0)
Me.chk3DStars.Value = IIf(modSpaceGame.cg_Stars3D, 1, 0)

Me.chkMap.Value = IIf(modSpaceGame.cg_ShowMap, 1, 0)
sldrMapLen.Value = modSpaceGame.cg_MapLen / 100
sldrMapLen_Click
sldrMapLen.Enabled = modSpaceGame.cg_ShowMap
lblMapLen.Enabled = modSpaceGame.cg_ShowMap

chkInvert.Value = IIf(modSpaceGame.cg_SpaceDisplayMode = modStickGame.cg_DisplayMode_Normal, 0, 1)
Me.chkMissileLock.Value = IIf(modSpaceGame.cg_ShowMissileLock, 1, 0)

St = Player(frmGame.FindPlayer(frmGame.MyID)).ShipType

If St = MotherShip Or St = SD Then
    modSpaceGame.UseAI = False
    chkAI.Value = 0
    chkAI.Enabled = False
End If

cmdCancel_Click
picColour.BorderStyle = 0
picCol.BorderStyle = 0
lblColourInfo.Caption = "R" & vbNewLine & "G" & vbNewLine & "B"
cmdApply.Enabled = False
cmdCancel.Enabled = False

Call Space_FormLoad(Me)

Me.Top = frmGame.Top + frmGame.height / 2 - Me.height / 2

'pos
Me.Left = frmGame.Left + frmGame.width '/ 2 - Me.Width / 2
If (Me.Left + Me.width) > Screen.width Then
    Me.Left = frmGame.Left - Me.width
End If
If Me.Left < 0 Then
    Me.Left = Screen.width - Me.width - 10 'frmGame.Left + frmGame.width - Me.width
End If
'end pos

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Space_FormLoad(Me, True)
modSpaceGame.GameClientSettingsFormLoaded = False
End Sub

'checkboxes
Private Sub chkDrawThick_Click()
modSpaceGame.cg_DrawThick = CBool(chkDrawThick.Value)
End Sub

Private Sub chkExplosions_Click()
modSpaceGame.cg_DrawExplosions = CBool(chkExplosions.Value)
End Sub

'Private Sub chkExplode_Click()
'modSpaceGame.BulletsExplode = CBool(chkExplode.Value)
'End Sub

Private Sub chkFPS_Click()
modSpaceGame.cg_ShowFPS = CBool(chkFPS.Value)
End Sub

Private Sub chkMouse_Click()
modSpaceGame.cl_UseMouse = CBool(chkMouse.Value)
sldrSens.Enabled = Not modSpaceGame.cl_UseMouse
lblSens.Enabled = sldrSens.Enabled
cmdCrosshair.Enabled = Not sldrSens.Enabled
End Sub

Private Sub chkAI_Click()
modSpaceGame.UseAI = CBool(chkAI.Value)

If modSpaceGame.UseAI = False Then
    frmGame.SetPlayerState frmGame.MyID, Player_None
ElseIf Player(frmGame.FindPlayer(frmGame.MyID)).ShipType = MotherShip Then
    modSpaceGame.UseAI = False
End If
End Sub

Private Sub chkNoCls_Click()
modSpaceGame.cg_Cls = Not CBool(chkNoCls.Value)
End Sub

Private Sub chkSmoke_Click()
modSpaceGame.cg_Smoke = CBool(chkSmoke.Value)

'If modSpaceGame.cg_Smoke = False Then
    'chkGunSmoke.Value = 0
'End If

'chkGunSmoke.Enabled = modSpaceGame.cg_Smoke

End Sub

Private Sub chkStars_Click()
modSpaceGame.cg_StarBG = CBool(chkStars.Value)
End Sub

Private Sub sldrMapLen_Change()
sldrMapLen_Click
End Sub

Private Sub sldrMapLen_Click()
Const Cap As String = "Map Size - "

modSpaceGame.cg_MapLen = sldrMapLen.Value * 100
lblMapLen.Caption = Cap & CStr(modSpaceGame.cg_MapLen / 100)

End Sub

Private Sub sldrMapLen_Scroll()
sldrMapLen_Click
End Sub

Private Sub sldrSens_Change()
sldrSens_Click
End Sub

Private Sub sldrSens_Click()
Const lblCap As String = "Sensetivity - "
frmGame.ROTATION_RATE = sldrSens.Value
lblSens.Caption = lblCap & CStr(sldrSens.Value)
End Sub

Private Sub cmdCrosshair_Click()
Unload frmGameCrossHair
Load frmGameCrossHair
frmGameCrossHair.Show vbModeless, frmGame
End Sub

'####################################################
'colour + name
'####################################################

Private Sub sldrCol_Change(Index As Integer)
sldrCol_Click Index
End Sub

Private Sub sldrCol_Click(Index As Integer)
Dim RGBCol As ptRGB
Dim lCol As Long

lCol = RGB(sldrCol(0).Value, sldrCol(1).Value, sldrCol(2).Value)
RGBCol = modSpaceGame.RGBDecode(lCol)

If RGBCol.Red < 150 And RGBCol.Blue < 150 And RGBCol.Green < 150 Then
    fraCol.Caption = "Player Settings - Colour is too dark"
Else
    fraCol.Caption = "Player Settings"
    picColour.BackColor = lCol
End If

cmdCancel.Enabled = True
cmdApply.Enabled = CBool(LenB(txtName.Text))

End Sub

Private Sub sldrCol_Scroll(Index As Integer)
sldrCol_Click Index
End Sub

Private Sub txtName_Change()
cmdApply.Enabled = CBool(LenB(txtName.Text))
cmdCancel.Enabled = True
End Sub
