VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGameCrossHair 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CrossHair Options"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkLead 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Lead CrossHair"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CheckBox chkPredCrossHair 
      Alignment       =   1  'Right Justify
      Caption         =   "Predator CrossHair"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Frame fraLead 
      Caption         =   "Lead CrossHair Colour"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4575
      Begin VB.PictureBox picCol2 
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   4155
         TabIndex        =   8
         Top             =   240
         Width           =   4215
         Begin VB.PictureBox picLead 
            Height          =   855
            Left            =   3840
            ScaleHeight     =   795
            ScaleWidth      =   195
            TabIndex        =   13
            Top             =   0
            Width           =   255
         End
         Begin MSComctlLib.Slider sldrCol2 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   10
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
         Begin MSComctlLib.Slider sldrCol2 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   11
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
         Begin MSComctlLib.Slider sldrCol2 
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   12
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
         Begin VB.Label lblCol2 
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
            TabIndex        =   9
            Top             =   0
            Width           =   150
         End
      End
   End
   Begin VB.Frame fraMain 
      Caption         =   "Main CrossHair Colour"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.PictureBox picCol 
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         Begin VB.PictureBox picMain 
            Height          =   855
            Left            =   3840
            ScaleHeight     =   795
            ScaleWidth      =   195
            TabIndex        =   6
            Top             =   0
            Width           =   255
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
   Begin MSComctlLib.Slider sldrThickness 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label lblWidth 
      Caption         =   "Crosshair Width - WW"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "frmGameCrossHair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ELead()
Dim i As Integer

fraLead.Enabled = modSpaceGame.cg_DrawLeadCrossHair

For i = 0 To 2
    sldrCol2(i).Enabled = modSpaceGame.cg_DrawLeadCrossHair
Next i

lblCol2.Enabled = modSpaceGame.cg_DrawLeadCrossHair

End Sub

Private Sub chkLead_Click()
modSpaceGame.cg_DrawLeadCrossHair = CBool(chkLead.Value)
Call ELead
End Sub

Private Sub chkPredCrossHair_Click()
modSpaceGame.cg_PredatorCrossHair = CBool(chkPredCrossHair.Value)
End Sub

Private Sub Form_Load()
Dim cRGB As ptRGB

TurnOffToolTip sldrThickness.hWnd

modSpaceGame.GameCrosshairFormLoaded = True

picMain.BorderStyle = 0
picLead.BorderStyle = 0

lblColourInfo.Caption = "R" & vbNewLine & "G" & vbNewLine & "B"
lblCol2.Caption = "R" & vbNewLine & "G" & vbNewLine & "B"

chkPredCrossHair.Value = IIf(modSpaceGame.cg_PredatorCrossHair, 1, 0)
chkLead.Value = IIf(modSpaceGame.cg_DrawLeadCrossHair, 1, 0)
sldrThickness.Value = modSpaceGame.cg_CrossHairWidth
sldrThickness_Click
Call ELead

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

cRGB = modSpaceGame.RGBDecode(modSpaceGame.cg_SpaceMainCrosshair)
sldrCol(0).Value = cRGB.Red
sldrCol(1).Value = cRGB.Green
sldrCol(2).Value = cRGB.Blue

cRGB = modSpaceGame.RGBDecode(modSpaceGame.cg_SpaceLeadCrosshair)
sldrCol2(0).Value = cRGB.Red
sldrCol2(1).Value = cRGB.Green
sldrCol2(2).Value = cRGB.Blue

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
modSpaceGame.GameCrosshairFormLoaded = False
Call Space_FormLoad(Me, True)
End Sub

'---------------------------------------------------
Private Sub sldrCol_Change(Index As Integer)
sldrCol_Click Index
End Sub

Private Sub sldrCol_Scroll(Index As Integer)
sldrCol_Click Index
End Sub

Private Sub sldrCol_Click(Index As Integer)

modSpaceGame.cg_SpaceMainCrosshair = RGB(sldrCol(0).Value, sldrCol(1).Value, sldrCol(2).Value)
picMain.BackColor = modSpaceGame.cg_SpaceMainCrosshair

End Sub

'---------------------------------------------------
Private Sub sldrCol2_Change(Index As Integer)
sldrCol2_Click Index
End Sub

Private Sub sldrCol2_Scroll(Index As Integer)
sldrCol2_Click Index
End Sub

Private Sub sldrCol2_Click(Index As Integer)

modSpaceGame.cg_SpaceLeadCrosshair = RGB(sldrCol2(0).Value, sldrCol2(1).Value, sldrCol2(2).Value)
picLead.BackColor = modSpaceGame.cg_SpaceLeadCrosshair

End Sub

Private Sub sldrThickness_Click()
modSpaceGame.cg_CrossHairWidth = sldrThickness.Value
lblWidth.Caption = "Crosshair Width - " & CStr(modSpaceGame.cg_CrossHairWidth)
End Sub

Private Sub sldrThickness_Scroll()
sldrThickness_Click
End Sub

Private Sub sldrThickness_Change()
sldrThickness_Click
End Sub
