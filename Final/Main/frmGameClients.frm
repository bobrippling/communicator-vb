VERSION 5.00
Begin VB.Form frmGameClients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Client List"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraForce 
      Caption         =   "ForceTeam/Kicking"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.PictureBox picForce 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   2895
         TabIndex        =   1
         Top             =   240
         Width           =   2895
         Begin VB.CommandButton cmdKick 
            Caption         =   "Kick Player"
            Height          =   375
            Left            =   1200
            TabIndex        =   5
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdForce 
            Caption         =   "ForceTeam"
            Height          =   375
            Left            =   1200
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
            TabIndex        =   6
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Red"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optnTeam 
            Alignment       =   1  'Right Justify
            Caption         =   "Neutral"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
      End
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   3000
      Left            =   2520
      Top             =   2160
   End
   Begin projMulti.ScrollListBox lstMain 
      Height          =   1215
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   3375
      _ExtentX        =   9340
      _ExtentY        =   4683
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "frmGameClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetStatus(ByVal T As String)
Const K As String = "Status: "
lblStatus.Caption = K & T
lblStatus.Refresh
End Sub

Private Sub cmdForce_Click()

Dim PlayerID As Integer, i As Integer
Dim vTeam As eTeams
Dim Txt As String

If cmdForce.Enabled Then
    Txt = cmdKick.Caption
    
    If LenB(Txt) Then
        On Error GoTo EH
        PlayerID = CInt(Right$(Txt, Len(Txt) - InStrRev(Txt, Space$(1), , vbTextCompare)))
        
        If PlayerID <> frmGame.MyID Then
            For i = eTeams.Neutral To eTeams.Blue '0 to 2
                If optnTeam(i).Value Then
                    vTeam = i
                    Exit For
                End If
            Next i
            
            
            modWinsock.SendPacket frmGame.socket, _
                Player(frmGame.FindPlayer(PlayerID)).ptSockAddr, _
                sForceTeams & CStr(vTeam)
            
            SetStatus "ForceTeam'd " & Trim$(Player(frmGame.FindPlayer(PlayerID)).Name) _
                & " to the " & GetTeamStr(vTeam) & " team"
            
        End If
    End If
    
End If

EH:
cmdForce.Enabled = False
End Sub

Private Sub cmdKick_Click()

Dim ID As Integer, Playeri As Integer
Dim Txt As String

If cmdKick.Enabled Then
    Txt = cmdKick.Caption
    
    If LenB(Txt) Then
        On Error GoTo EH
        ID = CInt(Right$(Txt, Len(Txt) - InStrRev(Txt, Space$(1), , vbTextCompare)))
        If ID Then
            Playeri = frmGame.FindPlayer(ID)
            modWinsock.SendPacket frmGame.socket, Player(Playeri).ptSockAddr, modSpaceGame.sKicks
        End If
    End If
End If

EH:
cmdKick.Enabled = False

End Sub

Private Sub lstMain_Click()
Const K As String = "Kick Player "
Dim Txt As String
Dim ID As Integer, i As Integer

If modSpaceGame.SpaceServer Then
    Txt = lstMain.Text
    
    If LenB(Txt) Then
        ID = CInt(Right$(Txt, Len(Txt) - InStrRev(Txt, Space$(1), , vbTextCompare)))
        
        i = frmGame.FindPlayer(ID)
        
        If i = -1 Then
            SetStatus "Error - Player Not Found"
            lstMain.RemoveItem lstMain.ListIndex
        Else
            
            If ID <> 0 And Not Player(i).IsBot Then 'modSpaceGame.IsBotID(ID) Then
                cmdKick.Caption = K & CStr(ID)
                cmdKick.Enabled = True
            Else
                cmdKick.Caption = Left$(K, Len(K) - 1)
                cmdKick.Enabled = False
            End If
            
        End If
        
    End If
Else
    cmdKick.Enabled = False
    cmdKick.Caption = Left$(K, Len(K) - 1)
End If

cmdForce.Enabled = cmdKick.Enabled

End Sub

Private Sub Form_Load()

cmdKick.Enabled = False
cmdForce.Enabled = False
Me.fraForce.Enabled = modSpaceGame.SpaceServer

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

tmrRefresh_Timer

SetStatus "Loaded Player List"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
modSpaceGame.GameClientsFormLoaded = False
Call Space_FormLoad(Me, True)
End Sub

Private Sub tmrRefresh_Timer()
Dim i As Integer

lstMain.Clear

For i = LBound(modSpaceGame.Player) To UBound(modSpaceGame.Player)
    lstMain.AddItem "Name: " & Trim$(modSpaceGame.Player(i).Name) & Space$(3) _
        & "ID: " & modSpaceGame.Player(i).ID
        '& "Bot: " & CStr(modSpaceGame.Player(i).IsBot) & Space$(3)
Next i

End Sub
