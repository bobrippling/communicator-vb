VERSION 5.00
Begin VB.Form frmFTPServers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTP Servers"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstRoot 
      Height          =   1620
      Left            =   5400
      TabIndex        =   6
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Use Server"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdPass 
      Caption         =   "View Passwords"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox lstPass 
      Height          =   1620
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.ListBox lstUser 
      Height          =   1620
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.ListBox lstHost 
      Height          =   1620
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      Caption         =   "                   Host                                 Username                  Password                               File Path"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   7935
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "status"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   195
      Width           =   3015
   End
End
Attribute VB_Name = "frmFTPServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbPassAnswered As Boolean

Private Function getSelectedServer() As Integer
Dim sHost As String
Dim iServer As Integer

sHost = lstHost.Text
For iServer = LBound(modFTP.FTP_Details) To UBound(modFTP.FTP_Details)
    If modFTP.FTP_Details(iServer).FTP_Host_Name = sHost Then
        Exit For
    End If
Next iServer

If iServer = UBound(modFTP.FTP_Details) + 1 Then
    iServer = 0
End If

getSelectedServer = iServer
End Function

Private Sub cmdSelect_Click()
Dim iServer As Integer

cmdSelect.Enabled = False

iServer = getSelectedServer()
frmMain.mnuOnlineFTPServerAr_Click iServer

If modFTP.iCurrent_FTP_Details = iServer Then lblStatus.Caption = "Server Selected"

End Sub

Private Sub cmdPass_Click()
Dim sPass As String

sPass = modVars.Password("Enter the master password", Me, "FTP Server Password View", , , 10)

If sPass = ":mmoC:" Then
    cmdPass.Enabled = False
    pbPassAnswered = True
    loadServers
    lblStatus.Caption = "Password Correct"
ElseIf LenB(sPass) Then
    lblStatus.Caption = "Password Incorrect"
Else
    lblStatus.Caption = vbNullString
End If

End Sub

Private Sub Form_Load()
cmdPass.Enabled = True
cmdSelect.Enabled = False
lblStatus.Caption = "Double click to edit"

pbPassAnswered = False
loadServers

FormLoad Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FormLoad Me, True
End Sub

Private Sub listClick(i As Integer)
On Error Resume Next
lstHost.ListIndex = i
lstUser.ListIndex = i
lstPass.ListIndex = i
lstRoot.ListIndex = i

cmdSelect.Enabled = True

If LenB(lblStatus.Caption) Then lblStatus.Caption = vbNullString
End Sub

Private Sub loadServers()
Dim i As Integer

lstHost.Clear
lstUser.Clear
lstPass.Clear
lstRoot.Clear

For i = LBound(modFTP.FTP_Details) To UBound(modFTP.FTP_Details)
    lstHost.AddItem modFTP.FTP_Details(i).FTP_Host_Name
    lstUser.AddItem modFTP.FTP_Details(i).FTP_User_Name
    If pbPassAnswered Then
        lstPass.AddItem modFTP.FTP_Details(i).FTP_Password
    Else
        lstPass.AddItem "-"
    End If
    lstRoot.AddItem IIf(LenB(modFTP.FTP_Details(i).FTP_Root), modFTP.FTP_Details(i).FTP_Root, ".")
Next i

End Sub

Private Sub lstHost_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
listClick lstHost.ListIndex
End Sub
Private Sub lstUser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
listClick lstUser.ListIndex
End Sub
Private Sub lstPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
listClick lstPass.ListIndex
End Sub
Private Sub lstRoot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
listClick lstRoot.ListIndex
End Sub

Private Sub lstHost_DblClick()
Dim sHost As String
Dim i As Integer

i = getSelectedServer()

sHost = lstHost.List(i)
sHost = modVars.Password("Enter a host", Me, "FTP Host", sHost, False, 40)

If LenB(sHost) Then
    modFTP.FTP_Details(i).FTP_Host_Name = sHost
    loadServers
End If

End Sub

Private Sub lstUser_DblClick()
Dim sUser As String
Dim i As Integer

i = getSelectedServer()

sUser = lstUser.List(i)
sUser = modVars.Password("Enter a User", Me, "FTP User", sUser, False, 40)

If LenB(sUser) Then
    modFTP.FTP_Details(i).FTP_User_Name = sUser
    loadServers
End If

End Sub

Private Sub lstPass_DblClick()
Dim sPass As String
Dim i As Integer

If Not pbPassAnswered Then
    lblStatus.Caption = "Can't view passwords"
Else
    i = getSelectedServer()
    
    sPass = lstPass.List(i)
    sPass = modVars.Password("Enter a Password", Me, "FTP Password", sPass, False, 30)
    
    If LenB(sPass) Then
        modFTP.FTP_Details(i).FTP_Password = sPass
        loadServers
    End If
End If

End Sub

Private Sub lstRoot_DblClick()
Dim sRoot As String
Dim i As Integer

i = getSelectedServer()

sRoot = lstRoot.List(i)
sRoot = modVars.Password("Enter a Path", Me, "FTP Path", sRoot, False, 60)

If LenB(sRoot) Then
    modFTP.FTP_Details(i).FTP_Root = sRoot
    loadServers
End If

End Sub
