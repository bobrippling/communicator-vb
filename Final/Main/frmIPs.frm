VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmIPs 
   Caption         =   "Who's Online...?"
   ClientHeight    =   4305
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Myself"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdGetDetails 
      Caption         =   "Get User Details >>"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdConnectI 
      Caption         =   "Connect (Internal IP)"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdConnectE 
      Caption         =   "Connect (External IP)"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvOnline 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvUser 
      Height          =   3135
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblRightClick 
      Caption         =   "Right click the list for a menu"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Main"
      Begin VB.Menu mnuPopupGet 
         Caption         =   "Get Details"
      End
      Begin VB.Menu mnuPopupRefresh 
         Caption         =   "Refresh List"
      End
      Begin VB.Menu mnuPopupSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupAdd 
         Caption         =   "Add Myself"
      End
      Begin VB.Menu mnuPopupRemove 
         Caption         =   "Remove User"
      End
   End
End
Attribute VB_Name = "frmIPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const minX = 804, minY = 255 'px
Private Const naStr = "N/A"


Private Sub cmdAdd_Click()
cmdAdd.Enabled = False
mnuPopupAdd_Click
cmdAdd.Enabled = Not modLogin.bUploaded_User()
End Sub

Private Sub Form_Load()
Dim lH As Long

mnuPopup.Visible = False

With lvOnline.ColumnHeaders
    .Add , , "Name"
    '.Add , , "Internal IP"
    '.Add , , "External IP"
    .Add , , "Login Time"
    '.Add , , "PC Name"
End With
With lvUser.ColumnHeaders
    .Add , , "Name"
    .Add , , "Login Time" '(Local)"
    .Add , , "Version"
    .Add , , "PC Name"
    .Add , , "User Name"
    .Add , , "Internal IP"
    .Add , , "External IP"
End With

Me.width = ScaleX(minX, vbPixels, vbTwips) 'must be after column header settings, due to resize
Me.height = ScaleY(minY, vbPixels, vbTwips)

If frmMain.mnuOptionsAdvDisplayVistaControls.Checked Then
    lH = frmMain.GetCommandIconHandle()
    modDisplay.SetButtonIcon cmdConnectI.hWnd, lH
    modDisplay.SetButtonIcon cmdConnectE.hWnd, lH
End If


ConnectCmds False
cmdAdd.Enabled = Not modLogin.bUploaded_User()

If modLoadProgram.bIsIDE = False Then
    modSubClass.SubclassAuto Me
End If

lblStatus.Caption = "Loaded Window"

FormLoad Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If modLoadProgram.bIsIDE = False Then
    modSubClass.SubclassAuto Me, False
End If

FormLoad Me, True
End Sub

Private Sub Form_Resize()
Dim i As Integer, j As Integer
Const iGap As Integer = 120

On Error Resume Next

cmdRefresh.Top = Me.ScaleHeight - cmdRefresh.height - iGap
cmdGetDetails.Top = cmdRefresh.Top
cmdConnectI.Top = cmdRefresh.Top
cmdConnectE.Top = cmdRefresh.Top
cmdAdd.Top = cmdRefresh.Top


lblStatus.width = Me.ScaleWidth


lvUser.Left = Me.ScaleWidth / 3 + iGap / 2
lvUser.width = Me.ScaleWidth - lvUser.Left
lvUser.height = cmdRefresh.Top - lvOnline.Top - iGap

lvOnline.width = lvUser.Left - iGap / 2
lvOnline.height = lvUser.height


'cmdRefresh.Left = Me.ScaleWidth / 4 - cmdRefresh.width - iGap
cmdAdd.Left = cmdRefresh.Left + cmdRefresh.width + iGap
cmdGetDetails.Left = Me.lvUser.Left 'Me.ScaleWidth / 2 - cmdGetDetails.width + iGap
cmdConnectI.Left = lvUser.Left + lvUser.width / 2 - cmdConnectI.width / 2 'Me.ScaleWidth * 3 / 4 - cmdConnectI.width - iGap
cmdConnectE.Left = Me.ScaleWidth - cmdConnectE.width - iGap



'###############################

j = lvOnline.ColumnHeaders.Count
lvOnline.ColumnHeaders(1).width = lvOnline.width / j - iGap / 2
For i = 2 To j
    lvOnline.ColumnHeaders(i).width = lvOnline.ColumnHeaders(1).width
Next i


j = lvUser.ColumnHeaders.Count
lvUser.ColumnHeaders(1).width = lvUser.width / j - iGap / 6
For i = 2 To j
    lvUser.ColumnHeaders(i).width = lvUser.ColumnHeaders(1).width
Next i


End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Dim MMI As MINMAXINFO


If uMsg = WM_GETMINMAXINFO Then
    
    CopyMemory MMI, ByVal lParam, LenB(MMI)
    
    'set the MINMAXINFO data to the
    'minimum and maximum values set
    'by the option choice
    
    With MMI
        .ptMinTrackSize.X = minX
        .ptMinTrackSize.Y = minY
        '.ptMaxTrackSize.X = maxX
        '.ptMaxTrackSize.Y = maxY
    End With
    
    CopyMemory ByVal lParam, MMI, LenB(MMI)
    WindowProc = 0
Else
    WindowProc = modSubClass.CallWindowProc(GetProp(hWnd, WndProcStr), hWnd, uMsg, wParam, lParam)
End If

End Function

'################################################################################################
'################################################################################################
'################################################################################################

Private Sub ConnectCmds(bEn As Boolean)
If bEn = False Then
    cmdConnectE.Enabled = False
    cmdConnectI.Enabled = False
    cmdGetDetails.Enabled = False
End If

'lvOnline.Enabled = bEn
lvUser.Enabled = bEn
lvUser.ListItems.Clear
End Sub

Private Sub cmdRefresh_Click()
cmdRefresh.Enabled = False
Refresh_List
cmdRefresh.Enabled = True
End Sub

Private Sub lvOnline_Click()
cmdGetDetails.Enabled = Not (lvOnline.SelectedItem Is Nothing)
'If Not (lvOnline.SelectedItem Is Nothing) Then
'    If lvOnline.SelectedItem.Text = No_One_Online_Str Then
'        If lvOnline.SelectedItem.SubItems(1) = No_One_Online_Str2 Then
'            cmdGetDetails.Enabled = False
'        Else
'            cmdGetDetails.Enabled = True
'        End If
'    Else
'        cmdGetDetails.Enabled = True
'    End If
'End If
End Sub

Private Sub lvOnline_DblClick()
mnuPopupGet_Click
End Sub

Private Sub lvOnline_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    lblRightClick.Visible = False
    mnuPopupGet.Enabled = Not (lvOnline.SelectedItem Is Nothing)
    mnuPopupAdd.Enabled = Not modLogin.bUploaded_User()
    mnuPopupRemove.Visible = bDevMode
    mnuPopupRemove.Enabled = mnuPopupGet.Enabled
    
    PopupMenu mnuPopup, , , , mnuPopupGet
End If
End Sub

Private Sub lvUser_Click()

If Not (lvUser.SelectedItem Is Nothing) Then
    cmdConnectE.Enabled = (lvUser.SelectedItem.SubItems(6) <> naStr)
    cmdConnectI.Enabled = True
Else
    cmdConnectE.Enabled = False
    cmdConnectI.Enabled = False
End If

End Sub

Private Sub lvUser_DblClick()
cmdConnectE_Click
End Sub

Private Sub cmdGetDetails_Click()

cmdGetDetails.Enabled = False
Me.Refresh

If Not (lvOnline.SelectedItem Is Nothing) Then
    GetDetails lvOnline.SelectedItem.Text & Dot & FileExt
End If
End Sub

'################

Private Sub cmdConnectE_Click()
DoConnect True
End Sub
Private Sub cmdConnectI_Click()
DoConnect False
End Sub

Private Sub DoConnect(ByVal bExternal As Boolean)
Dim sIP As String

cmdConnectI.Enabled = False
cmdConnectE.Enabled = False


If Not (lvUser.SelectedItem Is Nothing) Then
    sIP = lvUser.SelectedItem.SubItems(IIf(bExternal, 6, 5))
    
    If bExternal Then
        If sIP = naStr Then
            MsgBoxEx "Can't Connect to '" & naStr & "' - they didn't upload their IP", "Hi there", vbExclamation, "Error", , , , , Me.hWnd
            Exit Sub
        End If
    End If
    
    Me.Hide
    frmMain.Connect sIP
    Unload Me
End If

End Sub

'################

Private Sub mnuPopupGet_Click()
cmdGetDetails_Click
End Sub

Private Sub mnuPopupAdd_Click()
lblStatus.Caption = "Adding to List..."
modLogin.AddToFTPList False
Refresh_List
End Sub

Private Sub mnuPopupRemove_Click()
Dim sUser As String
Dim sError As String
Dim bSuccess As Boolean

On Error GoTo EH
sUser = lvOnline.SelectedItem.Text

If LenB(sUser) Then
    lblStatus.Caption = "Removing " & sUser & "..."
    
    If sUser = modLogin.sNameUsed() Then
        bSuccess = modLogin.RemoveFromFTPList(, False, sError)
    Else
        bSuccess = modLogin.RemoveFromFTPList(sUser, False, sError)
    End If
    
    If bSuccess Then
        Refresh_List
    End If
    
    lblStatus.Caption = sError
End If

EH:
End Sub

Private Sub mnuPopupRefresh_Click()
Refresh_List
End Sub

'################################################################################################
'################################################################################################
'################################################################################################

Private Sub Refresh_List()

Dim sRemoteFile As String, sError As String
Dim DirList() As ptFTPFile
Dim i As Integer, iCurrentSetting As Integer
Dim bPotentialOld As Boolean

iCurrentSetting = modFTP.iCurrent_FTP_Details
modFTP.iCurrent_FTP_Details = 0 'force byethost server


sRemoteFile = modLogin.Users_Path()
sRemoteFile = Left$(sRemoteFile, Len(sRemoteFile) - 1) 'knock off the trailing '/'


If modFTP.ListDir(sRemoteFile, DirList, sError) Then
    If modVars.FileArrayDimensioned(DirList) Then
        
        lvOnline.ListItems.Clear
        For i = 0 To UBound(DirList)
            lvOnline.ListItems.Add , , RemoveFileExt(DirList(i).sName)
            lvOnline.ListItems(lvOnline.ListItems.Count).SubItems(1) = CStr( _
                DateAdd("h", 1, DirList(i).dDateLastWritten))
            
'            If DateDiff("d", Date, DirList(i).dDateLastWritten) > 1 Then
'                bPotentialOld = True
'            End If
        Next i
        
        ConnectCmds True
        
        lblStatus.Caption = "Listed Online Users"
        
        '############################################################
        'remove laggers
'        If bPotentialOld Then
'            If MsgBoxEx("Someone appears to have been logged on for more than a day." & vbNewLine & _
'                        "Their Communicator may not have been able to log them out." & vbNewLine & vbNewLine & _
'                        "Log them out?", "Read above. It's all explained there", vbQuestion + vbYesNo, _
'                        "Remove Laggy Randomer from Server?", , , , , Me.hWnd) = vbYes Then
'
'
'                For i = 0 To UBound(DirList)
'                    If DateDiff("d", Date, DirList(i).dDateLastWritten) > 1 Then
'                        lblStatus.Caption = "Removing " & RemoveFileExt(DirList(i).sName) & "..."
'
'                        'still in byethost mode, so it's ok
'                        modFTP.DelFTPFile modLogin.Users_Path() & DirList(i).sName, sError
'
'                    End If
'                Next i
'
'
'                lblStatus.Caption = "Thanks for the housekeeping. Listed Online Users"
'
'            End If
'        End If
        '############################################################
    Else
        lblStatus.Caption = "No one is online..."
        
        lvOnline.ListItems.Clear
        ConnectCmds False
    End If
Else
    lblStatus.Caption = "Error - " & sError
End If


lblRightClick.Visible = False
modFTP.iCurrent_FTP_Details = iCurrentSetting
End Sub

Private Sub GetDetails(ByVal sFileName As String)
Dim sDetails As String, sError As String
Dim eError As eFTPCustErrs
Dim DetailAr() As String
Dim i As Integer

modFTP.GetFileStr sDetails, eError, modLogin.Users_Path() & sFileName, cmdRefresh, sError, True

If eError = cSuccess Then
    'name|pcname|internal|external
    DetailAr = Split(sDetails, modLogin.Detail_Sep)
    
    lvUser.ListItems.Add , , DetailAr(0)
    
    With lvUser.ListItems(lvUser.ListItems.Count)
        For i = 1 To UBound(DetailAr)
            If i = 3 Then
                .SubItems(i) = LCase$(DetailAr(i))
            Else
                .SubItems(i) = DetailAr(i)
            End If
        Next i
    End With
    
    Erase DetailAr
    
    
    lblStatus.Caption = "Retrieved User Details for " & modVars.RemoveFileExt(sFileName)
    
    
ElseIf eError = cFileNotFoundOnServer Then
    If LenB(sError) Then
        lblStatus.Caption = "Error - " & sError
    Else
        lblStatus.Caption = "Error - User has logged off"
    End If
    
ElseIf eError = cFileNotFoundOnLocal Or eError = cOther Then
    If LenB(sError) Then
        lblStatus.Caption = "Error - " & sError
    Else
        lblStatus.Caption = "Error in download...?"
    End If
Else
    lblStatus.Caption = "You should never see this text"
End If


End Sub
