VERSION 5.00
Begin VB.Form frmNetwork 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Computers"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   11460
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "User Details >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin projMulti.ScrollListBox lstComputer 
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3201
   End
   Begin projMulti.ScrollListBox lstComment 
      Height          =   1815
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3201
   End
   Begin projMulti.ScrollListBox lstUser 
      Height          =   1815
      Left            =   7920
      TabIndex        =   5
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3201
   End
End
Attribute VB_Name = "frmNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
Dim IP As String

cmdConnect.Enabled = False

IP = lstComputer.Text

If LenB(IP) Then
    
    Unload Me
    
    frmMain.Connect IP
    
End If

End Sub

Private Sub cmdRefresh_Click()
cmdRefresh.Enabled = False
cmdRefresh.Caption = "Refreshing..."
Me.Refresh
RefreshList
cmdRefresh.Caption = "Refresh"
cmdRefresh.Enabled = True
End Sub

Private Sub RefreshList()
Dim sError As String

If frmMain.RefreshNetwork(sError, lstComputer, lstComment, True) = False Then
    
    MsgBoxEx "Error - " & sError, "An error occured listing the network PCs", _
                    vbExclamation, "Network Error"
    
    
End If

End Sub

Private Sub cmdUser_Click()
Dim strPC As String, sError As String
Dim UserList As modNetwork.ListOfUserExt
Dim i As Integer

cmdUser.Enabled = False
cmdUser.Caption = "Wait a sec..."

strPC = lstComputer.Text

If LenB(strPC) Then
    UserList = modNetwork.LongEnumUsers(strPC, sError)
    
    If UserList.Init Then
        
        For i = 1 To UBound(UserList.List)
            
            lstUser.AddItem UserList.List(i).Name & _
                IIf(LenB(UserList.List(i).Comment), " - " & UserList.List(i).Comment, vbNullString)
            
        Next i
        
    Else
        If LenB(sError) Then
            MsgBoxEx "Error - " & sError, "An error occured listing the users", vbExclamation, _
                    "Network Error"
            
        Else
            MsgBoxEx "Error Code " & CStr(UserList.LastErr) & " occured", "An error occured listing the users", _
                    vbExclamation, "Network Error"
            
        End If
    End If
End If

cmdUser.Caption = "User Details >>"

End Sub

Private Sub Form_Load()

Me.Left = frmMain.Left - Me.width
If Me.Left < 10 Then
    Me.Left = 10
End If

Me.Top = frmMain.Top + frmMain.height / 2 - Me.height / 2

If frmMain.mnuOptionsAdvDisplayVistaControls.Checked Then
    modDisplay.SetButtonIcon cmdConnect.hWnd, frmMain.GetCommandIconHandle()
End If

FormLoad Me

Me.Show vbModeless, frmMain

cmdRefresh_Click

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub

Private Sub lstComment_Click()
Call listClick(lstComment.ListIndex)
End Sub

Private Sub lstComputer_Click()
Call listClick(lstComputer.ListIndex)
End Sub

Private Sub listClick(i As Integer)
On Error Resume Next

lstComputer.ListIndex = i
lstComment.ListIndex = i

cmdConnect.Enabled = (i <> -1)
cmdUser.Enabled = cmdConnect.Enabled

End Sub

Private Sub lstComputer_DblClick()
cmdConnect_Click
End Sub

Private Sub lstComment_DblClick()
lstComputer_DblClick
End Sub
