VERSION 5.00
Begin VB.Form frmStickClients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stick Game Client List"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKick 
      Caption         =   "Kick Player"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   3000
      Left            =   2520
      Top             =   1320
   End
   Begin projMulti.ScrollListBox lstMain 
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   960
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
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "frmStickClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetStatus(ByVal T As String)
Const k As String = "Status: "
lblStatus.Caption = k & T
lblStatus.Refresh
End Sub

Private Sub cmdKick_Click()
Dim ID As Integer, Sticki As Integer
Dim Txt As String

cmdKick.Enabled = False

Txt = lstMain.Text

ID = CInt(Right$(Txt, Len(Txt) - InStrRev(Txt, Space$(1), , vbTextCompare)))

If LenB(Txt) Then
    On Error GoTo EH
    Sticki = frmStickGame.FindStick(ID)
    
    If Sticki <> -1 Then
        If Stick(Sticki).IsBot Then
            SetStatus "Error - Can't kick a bot"
        ElseIf Sticki = 0 Then
            SetStatus "Error - Can't kick self"
        Else
            modWinsock.SendPacket frmStickGame.lSocket, Stick(Sticki).ptsockaddr, sKicks
        End If
    End If
End If

EH:
End Sub

Private Sub lstMain_Click()
cmdKick.Enabled = (LenB(lstMain.Text) And modStickGame.StickServer)
End Sub

Private Sub Form_Load()

cmdKick.Enabled = False

Me.Top = frmStickGame.Top + frmStickGame.height / 2 - Me.height / 2

'pos
Me.Left = frmStickGame.Left + frmStickGame.width '/ 2 - Me.Width / 2
If (Me.Left + Me.width) > Screen.width Then
    Me.Left = frmStickGame.Left - Me.width
End If
If Me.Left < 0 Then
    Me.Left = Screen.width - Me.width 'frmGame.Left + frmGame.width - Me.width
End If
'end pos

Call Stick_FormLoad(Me)


tmrRefresh_Timer

SetStatus "Loaded Stick List"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Stick_FormLoad(Me, True)
End Sub

Private Sub tmrRefresh_Timer()
Dim i As Integer

lstMain.Clear

For i = 0 To modStickGame.NumSticks - 1
    lstMain.AddItem "Name: " & Trim$(modStickGame.Stick(i).Name) & Space$(3) _
        & "ID: " & modStickGame.Stick(i).ID
Next i

End Sub
