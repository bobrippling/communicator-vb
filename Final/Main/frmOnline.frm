VERSION 5.00
Begin VB.Form frmOnline 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Online Status"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdOffline 
      Caption         =   "Set Offline"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdOnline 
      Caption         =   "Set Online"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin projMulti.ScrollListBox lstOnline 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3413
   End
End
Attribute VB_Name = "frmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Sep As String = "#"

Private Sub cmdRefresh_Click()
Call RefreshList
End Sub

Private Sub Form_Load()
Call FormLoad(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub

Private Sub RefreshList()
Dim Str As String, Name As String
Dim i As Integer, Online As Boolean
Dim Ar() As String

cmdRefresh.Enabled = False

Str = modFTP.GetOnlineStr()

Ar = Split(Str, vbNewLine, , vbTextCompare)

For i = LBound(Ar) To UBound(Ar)
    If Len(Ar(i)) > 0 Then
        On Error Resume Next
        Name = Left$(Ar(i), InStr(1, Ar(i), Sep, vbTextCompare) - 1)
        Online = CBool(Right$(Ar(i), 1))
        
        If Err.Number = 0 Then
            lstOnline.AddItem Name & Space$(3) & "Online: " & CStr(Online)
        End If
    End If
Next i

cmdRefresh.Enabled = True
End Sub
