VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DX7 Tester"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   5295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   9495
   End
   Begin VB.CommandButton Command 
      Caption         =   "Start Test"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_Click()
modDXSound.DXSound_Init Me.hWnd
End Sub

Public Sub AddToLog(ByVal sTxt As String)

With txtLog
    .SelStart = Len(.Text)
    .SelText = sTxt & vbNewLine
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
modDXSound.DXSound_Terminate
End Sub
