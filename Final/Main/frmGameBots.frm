VERSION 5.00
Begin VB.Form frmGameBots 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bots"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Bot"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Bot"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin projMulti.ScrollListBox lstBots 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "frmGameBots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Dim i As Integer

'For i = 1 To frmGame.NumPlayers - 1
    'If Player(i).IsBot Then
        'SetStatus "Can't add more than one bot at the moment"
        'Exit Sub
    'End If
'Next i

Call frmGame.AddBot

SetStatus "Added Bot " & CStr(i)

Call ListBots

End Sub

Private Sub cmdRemove_Click()
Dim i As Integer

cmdRemove.Enabled = False

On Error GoTo EH
i = lstBots.Text
If i <> 0 And i <> -1 Then
    RemoveBot i
End If

EH:
Call ListBots
End Sub

Private Sub Form_Load()
Call FormLoad(Me, , False)
ListBots
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True, False)
End Sub

Private Sub ListBots()
Dim i As Integer

lstBots.Clear

For i = 0 To frmGame.NumPlayers - 1
    If Player(i).IsBot Then
        lstBots.AddItem Player(i).ID
    End If
Next i

End Sub

Private Sub lstBots_Click()
cmdRemove.Enabled = (Len(lstBots.Text) > 0)
End Sub

Private Sub RemoveBot(ByVal i As Integer)
frmGame.RemovePlayer i
SetStatus "Removed Bot " & CStr(i)
End Sub

Private Sub SetStatus(ByVal S As String)
lblStatus.Caption = "Status: " & S
lblStatus.Refresh
End Sub
