VERSION 5.00
Begin VB.Form frmDevData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Developer Procedure Calls"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExe 
      Caption         =   "Execute"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton optnType 
      Caption         =   "Send Data"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton optnType 
      Caption         =   "Data Arrival"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtIndex 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblInfo 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label lblIndex 
      Caption         =   "Index (Optional):"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblData 
      Caption         =   "Data:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmDevData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExe_Click()

Dim sIndex As String, Data As String
Dim iIndex As Integer
Dim bIndexPresent As Boolean

Data = Trim$(txtData.Text)
sIndex = txtIndex.Text

bIndexPresent = CBool(LenB(sIndex))

If optnType(0).Value Then 'arrival
    If bIndexPresent Then
        iIndex = val(sIndex)
        Call DataArrival(Data, iIndex)
    Else
        Call DataArrival(Data)
    End If
Else 'send
    If bIndexPresent Then
        iIndex = val(sIndex)
        Call SendData(Data, iIndex)
    Else
        Call SendData(Data)
    End If
End If

lblInfo.Caption = IIf(optnType(0).Value, "Recieved", "Sent") & ": '" & Data & "'" & IIf(bIndexPresent, " to " & sIndex, vbNullString) & "."

End Sub

Private Sub Form_Load()
modDev.bDevDataFormLoaded = True

Call FormLoad(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
modDev.bDevDataFormLoaded = False
End Sub

Private Sub optnType_Click(Index As Integer)
cmdExe.Enabled = True
End Sub

Private Sub txtIndex_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr$(KeyAscii)) Then
    KeyAscii = 0
    Beep
End If
End Sub
