VERSION 5.00
Begin VB.Form frmIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Addresses"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   2970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClip 
      Caption         =   "Copy IP to Clipboard"
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblRemote 
      AutoSize        =   -1  'True
      Caption         =   "Remote"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   555
   End
   Begin VB.Label lblLocal 
      AutoSize        =   -1  'True
      Caption         =   "Local"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   390
   End
End
Attribute VB_Name = "frmIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const klIP As String = "Local IP Address: "
Private Const krIP As String = "Remote IP Address: "

Private Sub cmdClip_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText modVars.rIP
End Sub

Private Sub Form_Load()
Dim brIP As Boolean

Call FormLoad(Me)

lblLocal.Caption = klIP
lblRemote.Caption = krIP

Screen.MousePointer = vbHourglass

If Len(lIP) = 0 Then lIP = frmMain.SckLC.LocalIP
If Len(rIP) = 0 Then rIP = frmMain.GetIP()

rIP = frmMain.GetIP()

brIP = Not CBool(InStr(1, rIP, "Error:", vbTextCompare))

If brIP = False Then
    rIP = vbNullString
End If

lblLocal.Caption = klIP & lIP
lblRemote.Caption = IIf(Len(rIP) > 0, krIP & rIP, "Error fetching Remote IP," & vbNewLine & "Please wait a minute before retrying")

Screen.MousePointer = vbNormal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub
