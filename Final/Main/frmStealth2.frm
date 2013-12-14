VERSION 5.00
Begin VB.Form frmStealth2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkMatch 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtIn 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblFind 
      Caption         =   "Fi&nd What:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmStealth2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = frmStealth.Icon
Me.Move (frmStealth.Left + (frmStealth.width / 2)) - (Me.width / 2), _
    (frmStealth.Top + (frmStealth.height / 2)) - (Me.height / 2)

Me.Show vbModeless, frmStealth
frmStealth.Form2Loaded = True
End Sub

Private Sub cmdFind_Click()
If Status = Connected Then
    frmMain.cmdSend_Click
    txtIn.Text = frmMain.txtOut.Text 'aka ""
Else
    frmStealth.AddText "Error - Not Connected", True
    txtIn.Text = vbNullString
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmStealth.Form2Loaded = False
End Sub

Private Sub txtIn_Change()
frmMain.txtOut.Text = txtIn.Text
cmdFind.Enabled = LenB(txtIn.Text)
cmdFind.Default = cmdFind.Enabled
End Sub
