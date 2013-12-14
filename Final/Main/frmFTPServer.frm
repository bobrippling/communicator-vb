VERSION 5.00
Begin VB.Form frmFTPServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom FTP Server"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkChars 
      Caption         =   "Show Characters"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Okay"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtRoot 
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtPass 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtUser 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtHost 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblRoot 
      Caption         =   "Root Path"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblPass 
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblUser 
      Caption         =   "Username"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblHost 
      Caption         =   "Host Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmFTPServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Dim i As Integer
Dim sRootTmp As String

If LenB(txtHost.Text) = 0 Or _
        LenB(txtUser.Text) = 0 Or _
        LenB(txtPass.Text) = 0 Or _
        LenB(txtRoot.Text) = 0 Then
   
    MsgBoxEx "More data needed!", "You need to make sure all the fields are filled", vbExclamation, "Error", , , , , Me.hWnd
    
Else
    
    sRootTmp = txtRoot.Text
    If Right$(sRootTmp, 1) = "/" Then sRootTmp = Left$(sRootTmp, Len(sRootTmp) - 1)
    
    If modFTP.FTP_iCustomServer > -1 Then
        With modFTP.FTP_Details(modFTP.FTP_iCustomServer)
            .FTP_File_Ext = modVars.FileExt
            .FTP_Host_Name = txtHost.Text
            .FTP_Password = txtPass.Text
            .FTP_Root = sRootTmp
            .FTP_User_Name = txtUser.Text
        End With
    Else
        modFTP.Add_FTP_Server_Details txtHost.Text, txtUser.Text, txtPass.Text, sRootTmp, modVars.FileExt
        
        modFTP.FTP_iCustomServer = UBound(modFTP.FTP_Details)
    End If
    
    Unload Me
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim sHost As String, sUser As String, sPass As String, sRoot As String, sFileExt As String
Dim iLast As Integer

txtPass.PasswordChar = "*"

If modFTP.FTP_iCustomServer > -1 Then
    iLast = modFTP.iCurrent_FTP_Details
    modFTP.iCurrent_FTP_Details = FTP_iCustomServer
    
    modFTP.GetFTPDetails sHost, sUser, sPass, sRoot, sFileExt
    
    modFTP.iCurrent_FTP_Details = iLast
Else
    sRoot = "/"
End If

txtHost.Text = sHost
txtUser.Text = sUser
txtPass.Text = sPass
txtRoot.Text = sRoot


FormLoad Me

modVars.bModalFormShown = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FormLoad Me, True
modVars.bModalFormShown = False
End Sub

Private Sub chkChars_Click()
txtPass.PasswordChar = IIf(chkChars.Value = 1, vbNullString, "*")
End Sub
