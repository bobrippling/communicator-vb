VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Profile"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4560
   Begin VB.Frame fraStats 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4335
      Begin VB.Label lblLastDate 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label lblIP 
         Caption         =   "IP"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create User"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame fraCmds 
      Caption         =   "Maintenence"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   4335
      Begin VB.PictureBox picCmds 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   4095
         TabIndex        =   12
         Top             =   360
         Width           =   4095
         Begin VB.CommandButton cmdDel 
            Caption         =   "Delete User"
            Height          =   375
            Left            =   2160
            TabIndex        =   14
            Top             =   0
            Width           =   1935
         End
         Begin VB.CommandButton cmdNewPassword 
            Caption         =   "New Password"
            Height          =   375
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1935
         End
      End
   End
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Logout"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status: "
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblName 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EncryptionKey As String
'constants ^

Private pbLoggedIn As Boolean
Private DlbClickFra1 As Boolean, DlbClickFra2 As Boolean, DlbClickForm As Boolean

Private Type ptLoginData
    'sUserName As String
    sPassword As String
    
    dLastLogin As Date
    
    IP As String
End Type

Private Const Split_Sep = "|"

Private Property Get FTPFileRoot() As String
FTPFileRoot = modFTP.FTP_Root_Location & "/IP Data/Users/"
End Property

Private Function MakeLoginData() As String

MakeLoginData = modEncrypt.CryptString(txtPass.Text, EncryptionKey) & Split_Sep & _
            CStr(Date$) & Split_Sep & _
            modWinsock.RemoteIP & Split_Sep


End Function

Private Function GetLoginData(sData As String) As ptLoginData
Dim Parts() As String

Parts = Split(sData, Split_Sep, , vbTextCompare)

With GetLoginData
    '.sUserName = Parts(0)
    .sPassword = modEncrypt.CryptString(Parts(0), EncryptionKey)
    .dLastLogin = CDate(Parts(1))
    .IP = Parts(2)
End With

Erase Parts

End Function

'###############################################################

Private Sub LoginCmds(ByVal bLoggedIn As Boolean)
Const Cap As String = "User Profile - Logged "

cmdLogin.Enabled = False
cmdLogout.Enabled = bLoggedIn
txtName.Enabled = Not bLoggedIn
txtPass.Enabled = Not bLoggedIn
cmdNewPassword.Enabled = bLoggedIn
cmdCreate.Enabled = False
cmdDel.Enabled = bLoggedIn

If Not bLoggedIn Then
    'txtName.Text = vbNullString
    txtPass.Text = vbNullString
End If

pbLoggedIn = bLoggedIn

Me.Caption = Cap & IIf(bLoggedIn, "In", "Out")

End Sub

Private Sub SetStatus(ByVal T As String)
Const Max = 4290 'textwidth(string$(26, "W"))

'If TextWidth(T) > Max Then
    'T = Left$(T, 23)
'End If

lblStatus.Caption = "Status: " & T
lblStatus.Refresh
End Sub

'###############################################################

Private Sub cmdCreate_Click()
Dim sFileStr As String, rFileName As String, sError As String
Dim eType As eFTPCustErrs

DisableExtras
cmdCreate.Enabled = False
cmdLogin.Enabled = False
txtName.Enabled = False
txtPass.Enabled = False

If TrimAndValid() Then
    SetStatus "Creating User..."
    
    rFileName = txtName.Text & "." & modVars.FileExt
    
    'check if user exists
    modFTP.GetFileStr sFileStr, eType, FTPFileRoot & rFileName, cmdCreate, sError, True
    'dl current user's file
    
    If eType = cFileNotFoundOnServer Then
        'create user
        sFileStr = MakeLoginData()
        
        
        If modFTP.PutFileStr(sFileStr, FTPFileRoot & rFileName, cmdCreate, sError, True) Then
            SetStatus "User Created Successfully"
        Else
            SetStatus "Error" & IIf(LenB(sError), " (" & sError & ")", vbNullString)
        End If
        
        LoginCmds False
        
    ElseIf eType = cSuccess Then
        SetStatus "Error - User already exists"
        LoginCmds False
        
    ElseIf eType = cOther Or eType = cFileNotFoundOnLocal Then
        SetStatus "Error Checking For Existing User" & IIf(LenB(sError), " (" & sError & ")", vbNullString)
        LoginCmds False
        
    End If
    
'    Else
'        Files = Split(List, vbNewLine, , vbTextCompare)
'
'        For i = 0 To UBound(Files)
'            If LenB(Files(i)) <> 0 Then
'                If Files(i) = rFileName Then
'                    SetStatus "Error - User Exists"
'                    LoginCmds False
'                    Exit Sub
'                End If
'            End If
'        Next i
'
'
'
'    End If
End If

End Sub

Private Sub cmdDel_Click()
Dim sError As String

If MsgBoxEx("Delete " & txtName.Text & vbNewLine & _
            "Are you sure?", "Say bye bye to your score (and other stats)", _
            vbYesNo Or vbQuestion, "Delete User", , , frmMain.Icon) = vbYes Then
    
    
    If modFTP.DelFTPFile(FTPFileRoot & txtName.Text & "." & modVars.FileExt, sError) Then
        SetStatus "User Deleted"
        LoginCmds False
    Else
        SetStatus "Error Deleting User - " & sError
    End If
    
End If

End Sub

Private Sub cmdLogin_Click()

cmdLogin.Enabled = False
txtName.Enabled = False
txtPass.Enabled = False
cmdCreate.Enabled = False

If TrimAndValid() Then
    Call Login(txtName.Text, txtPass.Text)
Else
    txtName.Enabled = True
    txtPass.Enabled = True
End If


'do login ftp stuff

'lblKills.Caption = ftpkills + modSpaceGame.Kills
'?
End Sub

Private Sub cmdLogout_Click()
Dim sFile As String, sError As String

cmdLogout.Enabled = False
DisableExtras

SetStatus "Logging Out..."

sFile = MakeLoginData()

If modFTP.PutFileStr(sFile, FTPFileRoot & txtName.Text & "." & modVars.FileExt, cmdLogout, sError, False) Then
    SetStatus "Logged Out Successfully"
Else
    SetStatus "Error Uploading Stats - Logged Out" & IIf(LenB(sError), " (" & sError & ")", vbNullString)
End If

LoginCmds False

End Sub

Private Sub cmdNewPassword_Click()
Dim NP As String

NP = modVars.Password("Enter a New Password" & vbNewLine & "30 Characters Max.", frmLogin, "Communicator Account Password", , , 30)

If LenB(NP) Then
    txtPass.Text = NP
    cmdLogin.Enabled = False
    SetStatus "Password Changed (Logout to Save It)"
End If

End Sub

'########################################################################################

Private Sub Form_Load()
Dim sTmp As String

CreateEncryptionKey

LoginCmds False
DisableExtras

sTmp = modWinsock.RemoteIP
lblIP.Caption = "-" 'IIf(LenB(sTmp), sTmp, "Unknown IP")
lblLastDate.Caption = "-"

SetStatus "Logged Out"

Me.Left = frmMain.Left - Me.width
Me.Top = frmMain.Top + frmMain.height / 2 - Me.height / 2

If Me.Left < 10 Then Me.Left = 10

If frmMain.mnuOptionsAdvDisplayVistaControls.Checked Then
    modDisplay.SetButtonIcon cmdLogin.hWnd, frmMain.GetCommandIconHandle()
    modDisplay.SetButtonIcon cmdLogout.hWnd, frmMain.GetCommandIconHandle()
End If

Call FormLoad(Me, , , False)

'AddConsoleText "Loaded frmLogin", , True, , True
'AddConsoleText "frmLogin ThreadID: " & App.ThreadID


DlbClickFra1 = False
DlbClickFra2 = False
DlbClickForm = False
End Sub

Private Sub CreateEncryptionKey()
Const TempKey As String = "116 105 109 109 121 95 119 97 115 95 101 114 101" '"timmy_was_ere"
Dim KeyChars() As String
Dim i As Integer

KeyChars = Split(TempKey, vbSpace)

For i = 0 To UBound(KeyChars)
    EncryptionKey = EncryptionKey & Chr$(KeyChars(i))
Next i

Erase KeyChars

'CREATING THE 'TEMPKEY'...
'Const sKey As String = "timmy_was_ere"
'Dim sCode As String, i As Integer
'For i = 1 To Len(sKey)
'    sCode = sCode & " " & CStr(Asc(Mid$(sKey, i, 1)))
'Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If pbLoggedIn Then
    lblStatus.ForeColor = vbRed
    cmdLogout_Click
End If

Call FormLoad(Me, True)
End Sub

'########################################################################################

Private Sub Login(ByVal Name As String, ByVal Pass As String)

Dim sFile As String, sError As String
Dim eType As eFTPCustErrs
Dim ptData As ptLoginData

SetStatus "Logging In..."
Me.Refresh

modFTP.GetFileStr sFile, eType, FTPFileRoot & Name & "." & modVars.FileExt, cmdLogin, sError, True

Select Case eType
    Case eFTPCustErrs.cSuccess
        
        On Error GoTo EH
        ptData = GetLoginData(sFile)
        
        If Pass = ptData.sPassword Then
            
            If LenB(ptData.IP) Then
                lblIP.Caption = ptData.IP
            Else
                lblIP.Caption = "IP Unknown"
            End If
            lblLastDate.Caption = CStr(ptData.dLastLogin)
            
            SetStatus "Logged In Successfully"
            LoginCmds True
            
        Else
            
            If DlbClickFra1 And DlbClickFra2 And DlbClickForm And bDevMode Then
                SetStatus "Incorrect Password (" & ptData.sPassword & ")"
            Else
                SetStatus "Incorrect Password"
            End If
            
            txtName.Enabled = True
            txtPass.Enabled = True
            
            SetFocus2 txtPass
        End If
        
        
    Case eFTPCustErrs.cFileNotFoundOnServer
        SetStatus "User Doesn't Exist"
        LoginCmds False
        
        SetFocus2 txtName
        
        txtName.Selstart = 0
        txtName.Sellength = Len(txtName.Text)
        
    Case eFTPCustErrs.cFileNotFoundOnLocal
        SetStatus "Error - " & Err.Description
        LoginCmds False
        
    Case eFTPCustErrs.cOther
        SetStatus "Error" & IIf(LenB(sError), " (" & sError & ")", vbNullString)
        LoginCmds False
        
End Select

Exit Sub
EH:
'MsgBoxEx "Error Getting Login Data, Try Again", "Duh. There was an error logging in. Try again or don't", vbExclamation, "Error", , , , , Me.hWnd
SetStatus "Error Getting Login Data, User could be corrupt"
LoginCmds False
End Sub

Private Sub fraCmds_DblClick()
DlbClickFra2 = True
Call CheckDblBeep
End Sub

Private Sub fraStats_DblClick()
DlbClickFra1 = True
Call CheckDblBeep
End Sub

Private Sub Form_DblClick()
DlbClickForm = True
Call CheckDblBeep
End Sub

Private Sub CheckDblBeep()
If DlbClickForm And DlbClickFra1 And DlbClickFra2 And bDevMode Then
    Beep
End If
End Sub

Private Sub txtName_Change()
txtPass_Change
End Sub

Private Sub txtPass_Change()
cmdLogin.Enabled = (LenB(txtPass.Text) <> 0) And (LenB(txtName.Text) <> 0)
cmdLogin.Default = cmdLogin.Enabled
cmdCreate.Enabled = cmdLogin.Enabled
End Sub

Private Sub txtPass_GotFocus()
txtPass.Selstart = 0
txtPass.Sellength = Len(txtPass.Text)
End Sub

Private Sub DisableExtras()
cmdNewPassword.Enabled = False
cmdDel.Enabled = False
End Sub

Private Function TrimAndValid() As Boolean
Dim b As Boolean

'txtPass.Text = Trim$(txtPass.Text)
txtName.Text = Trim$(txtName.Text)

b = False

If LenB(txtName.Text) Then
    If LenB(txtPass.Text) Then
        b = True
    End If
End If

If b = False Then
    SetStatus "Please enter a username/password"
End If

TrimAndValid = b

End Function
