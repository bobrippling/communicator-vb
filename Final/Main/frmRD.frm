VERSION 5.00
Begin VB.Form frmRD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RD"
   ClientHeight    =   1575
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7755
   Icon            =   "frmRD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMain 
      Interval        =   1000
      Left            =   3240
      Top             =   1080
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox txtInt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   5520
         TabIndex        =   4
         Text            =   "1000"
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chkCheck 
         Caption         =   "Enable Checking"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.CheckBox chkExitProg 
         Caption         =   "Exit Communicator when they connect (otherwise, hide)"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   4455
      End
      Begin VB.Label lblInt 
         Alignment       =   2  'Center
         Caption         =   "Scan Interval - 1000ms"
         Height          =   255
         Left            =   5520
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label lblInfo 
      Caption         =   "Status: Safe"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   7335
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuPopupCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu mnuPopupSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CmdLine As String = "netstat -a -o" '-n
Private Const ZStr As String = "zentek"

Private Connected As Boolean

'str = "TCP    PC3095:1039            localhost:1038         ESTABLISHED     512" & vbnewline & "TCP    PC3095:1041            zenteck.localhost:1042         ESTABLISHED     544" & vbnewline & "TCP    PC3095:1038            localhost:1039         ESTABLISHED     512"
Private Function GetLine(Str As String, ByVal j As Integer) As String
Dim i As Integer
Dim PrevPart As String
Dim NextPart As String

For i = j To 1 Step -1
    If Mid$(Str, i, 2) = vbNewLine Then
        PrevPart = Mid$(Str, i + 2, j - i - 2)
        Exit For
    End If
Next i


For i = j To Len(Str)
    If Mid$(Str, i, 2) = vbNewLine Then
        NextPart = Mid$(Str, j, i - j)
        Exit For
    End If
Next i

GetLine = PrevPart & NextPart

End Function

Private Sub Check()
Dim Str As String
Dim Ero As Boolean
Dim i As Integer, j As Integer
Dim Line As String

modCmd.ExecAndCapture CmdLine, Str, Ero


i = InStr(1, Str, "zentek", vbTextCompare)

''find line + print (may be multiple lines)
''If i <> 0 Then
''    Line =
''
''    i = InStr(i + 1, Str, ZStr, vbTextCompare)
''End If

'##########################################################

If i <> 0 Then
    
    Call OnConnect(GetLine(Str, i))
    Connected = True
    
ElseIf Not Connected Then
    
    Connected = False
    
End If


If Ero Then
    lblInfo.Caption = "Error!"
End If

End Sub

Private Sub OnConnect(ByVal Line As String)

If Not Connected Then
    
    If chkExitProg.Value = 1 Then
        modLoadProgram.ExitProgram
        Exit Sub
    Else
        Me.Hide
        frmMain.ShowForm False, False
        frmSystray.ShowBalloonTip "RD!", "RD", NIIF_ERROR, 1000, True
    End If
    
    lblInfo.Caption = Line
End If


End Sub

Private Sub chkCheck_Click()
tmrMain.Enabled = CBool(chkCheck.Value)

If tmrMain.Enabled = False Then
    lblInfo.Caption = "WARNING! Not scanning for ranger could be detrimental to your account"
Else
    lblInfo.Caption = "Status: Safe"
End If

End Sub

Private Sub tmrMain_Timer()
Call Check
End Sub

'##############

Public Sub ShowForm()

Call FormLoad(Me)
Me.Visible = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Me.Visible Then Call FormLoad(Me, True)
 
If UnloadMode = vbFormControlMenu Then
    Me.Hide
    Cancel = -1
End If

End Sub


Private Sub txtInt_Change()

With txtInt
    If LenB(.Text) > 0 Then
        On Error GoTo EH
        tmrMain.Interval = Val(.Text)
        
        lblInt.Caption = "Scan Interval - " & tmrMain.Interval & "ms"
        
    End If
End With

Exit Sub
EH:
txtInt.Text = vbNullString
MsgBox "Please enter a valid interval (in milliseconds)", vbExclamation, "Error"
End Sub

Private Sub txtInt_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
    If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
End If
End Sub
