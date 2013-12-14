VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPrivate 
   Caption         =   "Private Comm Channel - "
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox rtfIn 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2990
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmPrivate.frx":0000
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtPrivateOut 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private pSendTo As String
Private pSendToSock As Integer
Private currentLogFile As String
Private bTellOnClose As Boolean

Public Property Get SendToSock() As Integer
SendToSock = pSendToSock
End Property

Public Property Let SendToSock(ByVal iSock As Integer)
Dim i As Integer

pSendToSock = iSock

For i = 0 To UBound(Clients)
    If Clients(i).iSocket = iSock Then
        UpdateCaption Clients(i).sName
        Exit Property
    End If
Next i

UpdateCaption "Unknown"

End Property

Private Sub UpdateCaption(sTxt As String)

If LenB(currentLogFile) = 0 Then
    currentLogFile = frmMain.GetCurrentLogFolder() & sTxt & " Private " & frmMain.MakeTimeFile() & ".txt"
End If

Me.Caption = PvtCap & sTxt

End Sub

'Public Property Get SendTo() As String
'SendTo = pSendTo
'End Property
'
'Public Property Let SendTo(ByVal St As String)
'
'pSendTo = St
'Me.Caption = PvtCap & St
'
'End Property

Private Sub cmdSend_Click()
Dim sTxt As String

sTxt = Trim$(txtPrivateOut.Text)

If LenB(sTxt) Then
    SendPrivateMessage frmMain.LastName & modMessaging.MsgNameSeparator & sTxt
End If

txtPrivateOut.Text = vbNullString

SetFocus2 txtPrivateOut

'msg = eDevCmd & WhoTo & # & From & @ & Text
End Sub

Private Sub SendPrivateMessage(ByVal sTxt As String, Optional bLog As Boolean = True)
Dim Msg As String

Msg = eCommands.Prvate & pSendToSock & "#" & modMessaging.MySocket & "@" & sTxt

AddPvtText sTxt, frmMain.txtOut.ForeColor, False, , bLog

If Server Then
    DistributeMsg Msg, -1
Else
    SendData Msg
End If

End Sub

Private Sub Form_DblClick()
bTellOnClose = Not bTellOnClose

If bTellOnClose Then
    modSpeech.Say "Closing will now send a close message"
Else
    modSpeech.Say "Closing will now not send a close message"
End If

End Sub

Private Sub Form_Load()
bTellOnClose = True

Call FormLoad(Me)
'frmMain.txtName.Enabled = False
modVars.nPrivateChats = modVars.nPrivateChats + 1

Show vbModeless, frmMain

If Me.Left < 10 Then Me.Left = 10
If Me.Top < 10 Then Me.Top = 10
If Me.Left > Screen.width Then Me.Left = 10
If Me.Top > Screen.height Then Me.Top = 10

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If bTellOnClose Then
    SendPrivateMessage InfoStart & frmMain.LastName & " closed the private chat" & InfoEnd, False
End If

modVars.nPrivateChats = modVars.nPrivateChats - 1

'If modVars.nPrivateChats = 0 Then frmMain.txtName.Enabled = True

Call FormLoad(Me, True)

'frmMain.LogPrivate
End Sub

Private Sub Form_Resize()
On Error Resume Next
txtPrivateOut.width = Me.ScaleWidth - txtPrivateOut.Left - cmdSend.width - 100
cmdSend.Left = txtPrivateOut.Left + txtPrivateOut.width + 50
rtfIn.width = Me.ScaleWidth - rtfIn.Left * 2
rtfIn.height = Me.ScaleHeight - rtfIn.Top - 100
End Sub

Private Sub txtPrivateOut_Change()
Dim b As Boolean

b = LenB(txtPrivateOut.Text)

cmdSend.Enabled = b
cmdSend.Default = b

End Sub

Public Sub AddPvtText(ByVal Text As String, ByVal Colour As Long, Optional ByVal bFlash As Boolean = True, _
    Optional ByVal iSockFrom As Integer = 0, Optional bLog As Boolean = True)

Dim i As Integer


With rtfIn
    .Selstart = Len(.Text)
    .SelColor = Colour
    .SelText = vbNewLine & Text
End With

If bLog Then UpdateLog Text

If bFlash Then
    FlashWin Me.hWnd
    
    If modSpeech.sReceived Then
        If modSpeech.sSayName Then
            modSpeech.Say Text
        Else
            modSpeech.Say Mid$(Text, 2 + InStr(1, Text, modMessaging.MsgNameSeparator, vbTextCompare))
        End If
    End If
End If

If iSockFrom <> 0 Then
    i = FindClient(iSockFrom)
    
    If i > -1 Then
        UpdateCaption Clients(i).sName
    End If
End If

End Sub

Private Sub UpdateLog(sMessage As String)
Dim f As Integer

f = FreeFile()

On Error GoTo EH
Open currentLogFile For Append As #f
    Print #f, sMessage
Close #f

Exit Sub
EH:
Close #f
AddText "Error Logging File for " & GetFileName(currentLogFile) & " - " & Err.Description, TxtError, True
End Sub
