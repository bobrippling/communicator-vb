VERSION 5.00
Begin VB.Form frmStealth 
   Caption         =   "Untitled - Notepad"
   ClientHeight    =   5265
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7950
   Icon            =   "frmStealthT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMain 
      Height          =   1335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileReturn 
         Caption         =   "Return"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Send Message..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileRD 
         Caption         =   "Ranger Danger Window"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "&Edit"
      Begin VB.Menu mnuHost 
         Caption         =   "Host"
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "Manual Connect..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuConnectList 
         Caption         =   "List Connect..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuDisco 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGame 
         Caption         =   "Game Lobby"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmStealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Cap As String = "Untitled - Notepad"
Private Const CapN As String = "Untitled* - Notepad"
Public Form2Loaded As Boolean

Private Sub Form_Load()
Me.Caption = Cap

frmStealth.mnuGame.Enabled = (Status = Connected)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If modVars.StealthMode Then
    ExitProgram
End If
End Sub

Private Sub Form_Resize()
txtMain.width = Me.ScaleWidth
txtMain.height = Me.ScaleHeight
End Sub

Private Sub mnuConnect_Click()
frmManual.Show , Me
End Sub

Private Sub mnuConnectList_Click()
frmUDP.Show
End Sub

Private Sub mnuDisco_Click()
frmMain.CleanUp True
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileRD_Click()
frmMain.mnuFileRD_Click
End Sub

Private Sub mnuFileReturn_Click()
StealthMode = False
End Sub

Public Sub AddText(ByVal Text As String, ByVal NewLine As Boolean, Optional ByVal DoStar As Boolean = True)

With txtMain
    .Sellength = 0
    .Selstart = Len(.Text)
    If Asc(Text) = 8 Then
        .Selstart = Len(.Text) - 1
        .Sellength = 1
        .SelText = vbNullString
    Else
        .SelText = IIf(NewLine, vbNewLine, vbNullString) & Text
    End If
End With

If IsForegroundWindow(Me.hWnd) Then
    Me.Caption = Cap
ElseIf DoStar Then
    Me.Caption = CapN
End If

End Sub

Private Sub mnuFileSave_Click()
frmMain.mnuFileSaveCon_Click
End Sub

Public Sub mnuFileSend_Click()

If Form2Loaded Then
    Unload frmStealth2
Else
    DoEvents
    Load frmStealth2
    On Error Resume Next
    frmStealth2.txtIn.SetFocus
End If

End Sub

Private Sub mnuGame_Click()

mnuGame.Enabled = (Status = Connected)

If mnuGame.Enabled Then frmMain.mnuOptionsMessagingLobby_Click

End Sub

Private Sub mnuHost_Click()
frmMain.Listen False
End Sub

Private Sub txtMain_Click()
Me.Caption = Cap
End Sub

'Private Sub txtMain_KeyPress(KeyAscii As Integer)
'
'Dim StrOut As String * 2 'for strout=crlf
''Dim CurrentLine As String
''Dim Tmp As String
''Dim iTmp As Integer
'
'If KeyAscii <> 8 Then
'    txtMain.Selstart = Len(txtMain.Text)
'    StrOut = Chr$(KeyAscii)
'    txtMain.Sellength = Len(StrOut)
'ElseIf KeyAscii = 8 Then
'    If Len(txtMain.SelText) <> 0 Then
'        StrOut = vbNewLine
'    End If
''Else
'    'iTmp = Asc(LCase$(Chr$(KeyAscii)))
'    'If Asc("a") <= iTmp And iTmp <= Asc("z") Then
'    'StrOut = Chr$(KeyAscii)
'    'Else
'        'Exit Sub
'    'End If
'End If
'
'If Server Then
'    Call modMessaging.DistributeMsg(eCommands.matrixMessage & Str(TxtForeGround) & "#" & StrOut, -1)
'Else
'    SendData eCommands.matrixMessage & Str(TxtForeGround) & "#" & StrOut
'    'MidText StrOut, IIf(mnuOptionsMessagingColours.Checked, TxtForeGround, TxtSent)
'End If
'
'AddText StrOut, False
'
'End Sub
