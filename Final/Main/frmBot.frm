VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bot Management"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "Remove All Bots"
      Height          =   375
      Left            =   2400
      TabIndex        =   28
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   500
      Left            =   4200
      Top             =   4200
   End
   Begin projMulti.ScrollListBox lstBots 
      Height          =   1215
      Left            =   240
      TabIndex        =   26
      Top             =   5280
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2143
   End
   Begin VB.Frame fraCol 
      Caption         =   "Bot Colour"
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
         Begin VB.PictureBox picBotColour 
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
   Begin VB.Frame fraType 
      Caption         =   "Ship Type"
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
      Begin VB.PictureBox picShipType 
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   1755
         TabIndex        =   8
         Top             =   240
         Width           =   1815
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Star Destroyer"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   12
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Hornet"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   11
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Behemoth"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optnShipType 
            Alignment       =   1  'Right Justify
            Caption         =   "Raptor"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   1695
         End
      End
   End
   Begin VB.CommandButton cmdAddBot 
      Caption         =   "Add Bot"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   4320
      Width           =   4335
   End
   Begin VB.CommandButton cmdRemoveBot 
      Caption         =   "Remove Bot"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   6600
      Width           =   1935
   End
   Begin VB.OptionButton optnBot 
      Caption         =   "Easy Bot (Low Hull)"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   3540
      Width           =   1215
   End
   Begin VB.OptionButton optnBot 
      Caption         =   "Normal Bot"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   21
      Top             =   3600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CheckBox chkBotAI 
      Caption         =   "Bot AI"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   3600
      Width           =   855
   End
   Begin VB.Frame fraTeam 
      Caption         =   "Team"
      Height          =   1095
      Left            =   2640
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
      Begin VB.PictureBox picTeam 
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   1755
         TabIndex        =   14
         Top             =   240
         Width           =   1815
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Neutral"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   15
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
            TabIndex        =   16
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Blue"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   17
            Top             =   480
            Width           =   1695
         End
      End
   End
   Begin MSComctlLib.Slider sldrReaction 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   25
      SmallChange     =   10
      Min             =   100
      Max             =   300
      SelStart        =   100
      TickFrequency   =   25
      Value           =   100
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Bot Name"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label lblAIReaction 
      AutoSize        =   -1  'True
      Caption         =   "Bot Reaction Time (ms) - WWW"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   2880
      Width           =   2280
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub ControlsE(ByVal Enable As Boolean)
'Dim Ctrl As Control
'
'For Each Ctrl In Controls
'
'    If Ctrl.Name <> cmdRemoveBot.Name Then
'        If Ctrl.Name <> chkBotAI.Name Then
'            If Ctrl.Name <> sldrReaction.Name Then
'                If Ctrl.Name <> lblAIReaction.Name Then
'                    If Ctrl.Name <> lstBots.Name Then
'                        If Ctrl.Name <> lblTitle.Name Then
'                            Ctrl.Enabled = Enable
'                        End If
'                    End If
'                End If
'            End If
'        End If
'    End If
'
'Next Ctrl
'
'End Sub

Private Sub chkBotAI_Click()
modSpaceGame.sv_BotAI = CBool(chkBotAI.Value)
End Sub

Private Sub cmdAddBot_Click()
Dim i As Integer
Dim ShipT As eShipTypes
Dim vTeam As eTeams

'cmdAddBot.Enabled = False

'For i = 1 To frmGame.NumPlayers - 1
'    If Player(i).IsBot Then
'        SetStatus "A Bot is already in the game"
'        Exit Sub
'    End If
'Next i

For i = 0 To optnShipType.UBound
    If optnShipType(i).Value Then
        If i = 3 Then
            ShipT = SD
        Else
            ShipT = i
        End If
        Exit For
    End If
Next i

For i = eTeams.Neutral To eTeams.Blue
    If optnTeam(i).Value Then
        vTeam = i
        Exit For
    End If
Next i

frmGame.AddBot optnBot(0).Value, ShipT, picBotColour.BackColor, vTeam

SetStatus "Bot Added"

cmdRemoveAll.Enabled = True

'Call ControlsE(False)

'cmdRemoveBot.Enabled = True
'chkBotAI.Enabled = True

'If modSpaceGame.GameOptionFormLoaded Then
'    frmGameOptions.chkAI.Value = 0
'    frmGameOptions.chkAI.Enabled = False
'End If

End Sub

Private Sub cmdRemoveAll_Click()
Dim i As Integer

cmdRemoveAll.Enabled = False

Do While i < frmGame.NumPlayers
    If Player(i).IsBot Then
        frmGame.RemovePlayer i
        i = i - 1
    End If
    i = i + 1
Loop

frmGame.SendChatPacketBroadcast "All Bots Removed", Player(0).Colour

End Sub

Private Sub cmdRemoveBot_Click()
Dim i As Integer, BotID As Integer, BotI As Integer
Dim Txt As String
'Dim Ch As String * 1

cmdRemoveBot.Enabled = False

'For i = 1 To frmGame.NumPlayers - 1
    'If Player(i).IsBot Then
        'frmGame.RemovePlayer i
        'SetStatus "Removed Bot"
        'Exit Sub
    'End If
'Next i

Txt = Trim$(lstBots.Text)

'For i = Len(Txt) To 0 Step -1
'    Ch = Mid$(Txt, i, 1)
'    If Not IsNumeric(Ch) Then
'
'        BotID = Mid$(Txt, i + 1)
'
'        Exit For
'    End If
'Next i
BotID = -1

For i = 0 To frmGame.NumPlayers - 1
    If Trim$(Player(i).Name) = Txt Then
        BotID = Player(i).ID
        BotI = i
        Exit For
    End If
Next i


If BotID <> -1 Then
    
    'BotI = frmGame.FindPlayer(BotID)
    
    frmGame.SendChatPacketBroadcast "Bot Removed: " & Txt, _
        Player(BotI).Colour 'picBotColour.BackColor 'vbYellow
    
    frmGame.RemovePlayer BotI
    
    '#########################
    'If there's only one left, just erase the array
'    If NumBotIDs = 1 Then
'        Erase modSpaceGame.BotIDs
'    Else
'        'Remove the bullet
'        For i = BotID - 1 To NumBotIDs - 2
'            BotIDs(i) = BotIDs(i + 1)
'        Next i
'        'Resize the array
'        ReDim Preserve BotIDs(NumBotIDs - 1)
'    End If
'    NumBotIDs = NumBotIDs - 1
    '#########################
    
    SetStatus "Removed Bot"
    
'    If modSpaceGame.GameOptionFormLoaded Then
'        frmGameOptions.chkAI.Enabled = True
'    End If
    
    lstBots.ListIndex = IIf(lstBots.ListCount > 0, 0, -1)
    lstBots_Click
    
Else
    SetStatus "Bot Not Found"
    lstBots.RemoveItem lstBots.ListIndex
End If

End Sub

Private Sub Form_Load()

Dim RGBCol As ptRGB
Dim i As Integer

picShipType.BorderStyle = 0
picTeam.BorderStyle = 0
picBotColour.BorderStyle = 0
picCol.BorderStyle = 0
lblColourInfo.Caption = "R" & vbNewLine & "G" & vbNewLine & "B"

modSpaceGame.GameBotFormLoaded = True

'If frmGame.BotID = -1 Then
RGBCol = RGBDecode(vbYellow)
'Else
    'RGBCol = RGBDecode(Player(frmGame.FindPlayer(frmGame.BotID)).Colour)
'End If

sldrCol(0).Value = RGBCol.Red
sldrCol(1).Value = RGBCol.Green
sldrCol(2).Value = RGBCol.Blue

'cmdAddBot.Enabled = modSpaceGame.SpaceServer And Not modSpaceGame.UseAI
'cmdRemoveBot.Enabled = cmdAddBot.Enabled
'chksv_BotAI.Enabled = (frmGame.BotID <> 0)
chkBotAI.Value = IIf(modSpaceGame.sv_BotAI, 1, 0)

sldrReaction.Value = frmGame.AI_Sample_Rate
'sldrReaction_Change

'If frmGame.BotID <> -1 Then
    'Call ControlsE(False)
'Else
    'cmdRemoveBot.Enabled = False
'End If

For i = 0 To frmGame.NumPlayers - 1
    If Player(i).IsBot Then
        cmdRemoveAll.Enabled = True
        Exit For
    End If
Next i


Call FormLoad(Me, False, False)

'pos
'Me.Top = frmGame.Top + frmGame.height / 2 - Me.height / 2
'Me.Left = frmGame.Left + frmGame.width '/ 2 - Me.Width / 2
'If (Me.Left + Me.width) > Screen.width Then
'    Me.Left = frmGame.Left - Me.width
'End If
'If Me.Left < 0 Then
'    Me.Left = Screen.width - Me.width - 10 'frmGame.Left + frmGame.width - Me.width
'End If

Me.Top = frmGameOptions.Top + frmGameOptions.height / 2 - Me.height / 2
Me.Left = frmGameOptions.Left + 2 * frmGameOptions.width / 3
If (Me.Left + Me.width) > Screen.width Then
    Me.Left = Screen.width - Me.width - 10 'frmGame.Left + frmGame.width - Me.width
End If

'end pos

SetStatus "Loaded Window"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
modSpaceGame.GameBotFormLoaded = False
Call FormLoad(Me, True, False)
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
End Select

End Sub

Private Sub sldrCol_Change(Index As Integer)
sldrCol_Click Index
End Sub

Private Sub sldrCol_Click(Index As Integer)
Dim RGBCol As ptRGB
Dim lCol As Long

lCol = RGB(sldrCol(0).Value, sldrCol(1).Value, sldrCol(2).Value)
RGBCol = modSpaceGame.RGBDecode(lCol)

If RGBCol.Red < 150 And RGBCol.Blue < 150 And RGBCol.Green < 150 Then
    SetStatus "Colour is Too Dark", True
Else
    picBotColour.BackColor = lCol
    If lblStatus.Caption = "Status: Colour is Too Dark" Then
        SetStatus "Colour Accepted"
    End If
End If

End Sub

Private Sub sldrCol_Scroll(Index As Integer)
sldrCol_Click Index
End Sub

Private Sub sldrReaction_Change()

Const Cap As String = "Bot Reaction Time (ms) - "

lblAIReaction.Caption = Cap & CStr(sldrReaction.Value)

frmGame.AI_Sample_Rate = sldrReaction.Value

End Sub

Private Sub sldrReaction_Click()
sldrReaction_Change
End Sub

Private Sub sldrReaction_Scroll()
sldrReaction_Change
End Sub

Private Sub SetStatus(ByVal T As String, Optional ByVal Red As Boolean = False)
Const K As String = "Status: "
lblStatus.Caption = K & T
If Red Then
    lblStatus.ForeColor = vbRed
Else
    lblStatus.ForeColor = &HFF0000
End If
lblStatus.Refresh
End Sub

Private Sub tmrRefresh_Timer()
Dim i As Integer
Dim ListI As Integer

ListI = lstBots.ListIndex
lstBots.Clear

For i = 0 To frmGame.NumPlayers - 1
    If Player(i).IsBot Then
        lstBots.AddItem Trim$(Player(i).Name)
    End If
Next i

On Error Resume Next
lstBots.ListIndex = ListI
End Sub
