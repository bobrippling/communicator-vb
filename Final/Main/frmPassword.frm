VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter a Password"
   ClientHeight    =   1245
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735.588
   ScaleMode       =   0  'User
   ScaleWidth      =   4844.96
   Begin VB.CheckBox chkShowChars 
      Caption         =   "Show Characters"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3855
   End
   Begin VB.Timer tmrLoad 
      Interval        =   100
      Left            =   1680
      Top             =   720
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   780
   End
   Begin VB.TextBox txtPassword 
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   3885
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Prompt"
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bNumeric As Boolean, sPassChar As String
Private Const Norm_Height = 1620, Check_Height = 1725

Public Property Let MaxLen(nLen As Integer)
txtPassword.MaxLength = nLen
End Property

Public Property Let PassChar(nVal As String)
txtPassword.PasswordChar = nVal
sPassChar = nVal

If LenB(nVal) Then
    chkShowChars.Visible = True
    Me.height = Check_Height
Else
    Me.height = Norm_Height
End If

End Property

Public Property Let Default(nVal As String)
txtPassword.Text = nVal
txtPassword.Selstart = Len(nVal)
End Property

Public Property Let Numeric(nVal As Boolean)
bNumeric = nVal
End Property

Private Sub chkShowChars_Click()
If chkShowChars.Value = 1 Then
    txtPassword.PasswordChar = vbNullString
Else
    txtPassword.PasswordChar = sPassChar
End If
End Sub

Private Sub cmdCancel_Click()
modVars.uPassword = vbNullString
Unload Me
End Sub

Public Property Let Prompt(ByVal P As String)
lblPrompt.Caption = P
End Property

Private Sub cmdOK_Click()
modVars.uPassword = txtPassword.Text
Unload Me
End Sub

Private Sub Form_Load()
txtPassword.MaxLength = 0
chkShowChars.Visible = False
tmrLoad.Enabled = True

'formload done elsewhere
'or not
'or yes
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'modVars.uPassword = vbNullString
Call FormLoad(Me, True)
End Sub

Private Sub tmrLoad_Timer()
tmrLoad.Enabled = False

SetFocus2 Me.txtPassword

If LenB(txtPassword.Text) = 0 Then
    If bNumeric Then
        modDisplay.ShowBalloonTip txtPassword, "Greetings", "Enter some numbers, if you don't mind"
    Else
        modDisplay.ShowBalloonTip txtPassword, "Greetings", "Enter some text, if you don't mind"
    End If
End If

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If bNumeric Then
    If KeyAscii <> 8 Then
        If Not IsNumeric(Chr$(KeyAscii)) Then
            KeyAscii = 0
            Beep
        End If
    End If
End If
End Sub
