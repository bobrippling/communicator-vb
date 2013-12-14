VERSION 5.00
Begin VB.Form frmIPChooser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IP Chooser"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin projMulti.ScrollListBox lstIPs 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   7011
      _ExtentY        =   1931
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   "Select an IP to send to..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4935
   End
End
Attribute VB_Name = "frmIPChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbChooseSocket As Boolean

Private Sub cmdCancel_Click()
modVars.pIP_Choice = vbNullString
Unload Me
End Sub

Private Sub cmdOk_Click()
modVars.pIP_Choice = GetSelectedIP()

Unload Me
End Sub

Private Function GetSelectedIP() As String
Dim sTxt As String
Dim i As Integer

sTxt = lstIPs.Text

If LenB(sTxt) Then
    i = InStr(1, sTxt, vbSpace)
    If i Then
        GetSelectedIP = Left$(sTxt, i - 1)
    Else
        GetSelectedIP = sTxt
    End If
End If

End Function

Private Sub cmdRefresh_Click()
Dim i As Integer

lstIPs.Clear

If pbChooseSocket Then
    For i = 0 To UBound(Clients)
        If Clients(i).iSocket <> 0 And Clients(i).iSocket <> modMessaging.MySocket Then
            lstIPs.AddItem Clients(i).iSocket & " - " & IIf(LenB(Clients(i).sName), Clients(i).sName, "[Unknown Name]")
        End If
    Next i
Else
    For i = UBound(Clients) To 0 Step -1 'ensures that Client0 (i.e. Me) is at the bottom
        If LenB(Clients(i).sIP) > 0 Then
            
            lstIPs.AddItem Clients(i).sIP & _
                IIf(LenB(Clients(i).sName), " - " & Clients(i).sName, vbNullString)
            
        End If
    Next i
End If

End Sub

Private Sub Form_Load()
FormLoad Me, , , , True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub

Private Sub lstIPs_Click()
cmdOk.Enabled = CBool(LenB(lstIPs.Text))
End Sub

'################################################################################
'socket chooser

Public Property Let bChooseSocket(b As Boolean)
pbChooseSocket = b
cmdRefresh_Click
End Property
