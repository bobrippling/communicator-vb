VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmDev 
   Caption         =   "Developer Form"
   ClientHeight    =   8130
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11160
   Begin VB.TextBox txtDevLog 
      Height          =   3615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1920
      Width           =   11175
   End
   Begin VB.Frame fraStates 
      Caption         =   "States"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.Timer tmrStates 
         Interval        =   1000
         Left            =   10080
         Top             =   1320
      End
      Begin VB.VScrollBar vsSockets 
         Height          =   1455
         Left            =   10560
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblStates 
         AutoSize        =   -1  'True
         Caption         =   "Socket States Label"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1440
      End
   End
   Begin RichTextLib.RichTextBox RtfData 
      Height          =   2175
      Left            =   0
      TabIndex        =   4
      Top             =   5520
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3836
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDev.frx":0000
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuMainExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

modLoadProgram.frmDev_Loaded = True
Me.Visible = False
Call FormLoad(Me, , False)
tmrStates_Timer

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
    ShowForm False
    Cancel = True
Else
    modLoadProgram.frmDev_Loaded = False
    FormLoad Me, True, False
End If

End Sub

Public Sub ShowForm(bShow As Boolean)

modImplode.AnimateAWindow Me.hWnd, aRandom, Not bShow

If bShow Then
    Me.Show vbModeless, frmMain
Else
    Me.Hide
    Me.RtfData.Text = vbNullString
    'Me.txtDevLog.Text = vbNullString
End If

End Sub

Private Sub Form_Resize()

If Me.WindowState = vbMinimized Then Exit Sub

On Error Resume Next

RtfData.width = Me.ScaleWidth
RtfData.height = Me.ScaleHeight - RtfData.Top

txtDevLog.width = Me.ScaleWidth
fraStates.width = Me.ScaleWidth - fraStates.Left - 100
vsSockets.Left = fraStates.width - vsSockets.width - 200

End Sub

Public Sub AddDev(ByVal Text As String, ByVal socket As Integer, ByVal Arrival As Boolean)

Dim From As String

If InStr(1, Text, vbCrLf, vbTextCompare) Then Text = Replace$(Text, vbCrLf, "[CRLF]", , , vbTextCompare)

On Error Resume Next
If socket = -1 Then
    From = frmMain.SckLC.RemoteHost
Else
    From = frmMain.SockAr(socket).RemoteHost
End If
On Error GoTo 0

From = IIf(From = vbNullString, vbNullString, " (" & From & ")")

If Arrival Then
    DevAddText "Arrived: " & Text & vbNewLine & "From: " & socket & From & vbNewLine
Else
    DevAddText "Sent: " & Text & vbNewLine & "To: " & socket & From & vbNewLine
End If

End Sub

Private Sub DevAddText(ByVal Text As String)

RtfData.Selstart = Len(RtfData.Text)
RtfData.SelText = vbNewLine & Text

End Sub

Private Sub mnuMainExit_Click()
SetFocus2 frmMain
Unload Me
End Sub

Private Sub tmrStates_Timer()
Dim i As Integer
Dim St As String
Dim state As Integer

state = frmMain.ucVoiceTransfer.iCurSockStatus
St = "SckVoice: " & GetState(state) & " (" & state & ")" & vbNewLine

state = frmMain.ucFileTransfer.iCurSockStatus
St = St & "SckFT: " & GetState(state) & " (" & state & ")" & vbNewLine

state = frmMain.SckLC.state
St = St & "SckLC: " & GetState(state) & " (" & state & ")" & IP(frmMain.SckLC) & vbNewLine

On Error GoTo Nex

For i = 0 To (frmMain.SockAr.Count - 1)
    
    state = frmMain.SockAr(i).state
    
    St = St & "Socket " & i & ": " & GetState(state) & " (" & state & ")" & IP(frmMain.SockAr(i)) & vbNewLine
    
Nex:
Next i

lblStates.Caption = St

If i >= 6 Then
    vsSockets.Enabled = True
    vsSockets.Max = i * 207
Else
    vsSockets.Enabled = False
    lblStates.Top = 250
End If

End Sub

Private Function IP(socket As Winsock) As String
Dim Tmp As String

If socket.state = sckConnected Then
    Tmp = vbSpace & socket.RemoteHostIP
Else
    Tmp = vbNullString
End If

IP = Tmp

End Function

Private Sub vsSockets_Change()
vsSockets_Scroll
End Sub

Private Sub vsSockets_Scroll()
lblStates.Top = -vsSockets.Value + 250
End Sub
