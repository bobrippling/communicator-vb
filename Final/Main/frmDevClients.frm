VERSION 5.00
Begin VB.Form frmDevClients 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clients"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5130
   Begin VB.Timer tmrRefresh 
      Interval        =   3000
      Left            =   120
      Top             =   240
   End
   Begin VB.CommandButton cmdCloseSock 
      Caption         =   "Disconnect Socket"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin projMulti.ScrollListBox lstName 
      Height          =   2295
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4048
   End
   Begin projMulti.ScrollListBox lstIP 
      Height          =   2295
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   4048
   End
   Begin projMulti.ScrollListBox lstSocket 
      Height          =   2295
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   4048
   End
   Begin VB.Label lblInfo 
      Caption         =   "Note: The method to list clients differers from the normal client list"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblTitles 
      Caption         =   "Name                                           IP                         Socket"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "frmDevClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCloseSock_Click()
Dim i As Integer
Dim Nme As String

Nme = lstName.Text

With cmdCloseSock
    i = Right$(.Caption, Len(.Caption) - InStrRev(.Caption, Space$(1), , vbTextCompare))
    
    If i <> -1 And LenB(CStr(i)) <> 0 Then
        If LenB(Nme) Then
            If Nme = "?" Then
                Nme = "Unknown Name"
            End If
        End If
        
        frmMain.Kick i, Nme
        tmrRefresh_Timer
    End If
    
    .Enabled = False
    
End With

End Sub

Private Sub Form_Load()
cmdCloseSock.Enabled = False
'txtIP.Enabled = Server 'And (Status = Connected)
Call FormLoad(Me)
tmrRefresh_Timer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub

Private Sub lstIP_Click()
Call ListClick(lstIP.ListIndex)
End Sub

Private Sub lstName_Click()
Call ListClick(lstName.ListIndex)
End Sub

Private Sub lstSocket_Click()
Const K As String = "Disconnect Socket "

cmdCloseSock.Enabled = ((lstSocket.ListIndex <> -1) And Server)  'And (lstSocket.Text <> "-"))

If cmdCloseSock.Enabled Then
    cmdCloseSock.Enabled = True
    cmdCloseSock.Caption = K & Trim$(lstSocket.Text)
End If

Call ListClick(lstSocket.ListIndex)

End Sub

Private Sub tmrRefresh_Timer()
Dim i As Integer, j As Integer
Dim Nme As String, IP As String

If Status <> Connected Then
    If UBound(Clients) > 0 Then
        ReDim Clients(0)
    End If
    lstName.Clear
    lstSocket.Clear
    lstIP.Clear
Else
    lstName.Clear
    lstSocket.Clear
    lstIP.Clear
    
    If Server Then
        With frmMain
            For i = 1 To .SockAr.UBound '.Count - 1
                If .SockAr(i).State = sckConnected Then
                    
                    j = FindClient(i)
                    
                    If j = -1 Then
                        IP = "?"
                        Nme = "?"
                    Else
                        Nme = Clients(j).sName
                        IP = .SockAr(Clients(j).iSocket).RemoteHostIP
                    End If
                    
                    'For j = LBound(Clients) To UBound(Clients)
                        'On Error Resume Next
                        'If Clients(j).iSocket = i Then
                            'Nme = Clients(i).sName
                            'IP = .SockAr(Clients(i).iSocket).RemoteHostIP
                            'Exit For
                        'End If
                    'Next j
                    
                    
                    lstName.AddItem Nme
                    lstIP.AddItem IP
                    lstSocket.AddItem CStr(i)
                    
                End If
                
            Next i
        End With
    End If
End If

End Sub


Private Sub ListClick(ByVal i As Integer)
Static Doing As Boolean

If i <> -1 And Not Doing Then
    Doing = True
    On Error Resume Next
    lstName.ListIndex = i
    lstIP.ListIndex = i
    lstSocket.ListIndex = i
    
    tmrRefresh.Enabled = False
    tmrRefresh.Enabled = True 'reset it
    
    Doing = False
End If
End Sub
