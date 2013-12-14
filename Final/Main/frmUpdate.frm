VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Communicator Update"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6030
   Begin projMulti.VistaProg progUpdate 
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   397
   End
   Begin VB.CommandButton cmdPrev 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdNext 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox picInfo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label lblState 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   5775
   End
   Begin VB.Label lblInfo 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5775
   End
   Begin VB.Line lnButtons 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5564
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line lnInfo 
      X1              =   0
      X2              =   3720
      Y1              =   620
      Y2              =   620
   End
   Begin VB.Line lnButtons 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   5549
      Y1              =   2175
      Y2              =   2175
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'for zipping
Private WithEvents ZipO As clsZipExtraction
Attribute ZipO.VB_VarHelpID = -1

Private pbStealth As Boolean

Private Enum eUpdateStates
    uNone = 0
    uCheckingVersion
    uWaitingForDownloadConfirmation
    uDownloadingVersion
    uExtractingVersion
    uWaitingForRestartConfirmation
    uRestarting
End Enum
Private uCurrentState As eUpdateStates 'for command buttons


Private Function GetUpdateInfo( _
    bHaveOld As Boolean, bHaveCurrent As Boolean, obj_Disable As Object, _
    nMaj As Integer, nMin As Integer, nRev As Integer, _
    Optional sError As String) As Boolean

Dim i As Integer, j As Integer
Dim bError As Boolean
Dim oMaj As Integer, oMin As Integer, oRev As Integer
Dim VerHTMLTxt As String

modFTP.fGetVersion obj_Disable, sError, bError, VerHTMLTxt

If Not bError Then
    oMaj = App.Major
    oMin = App.Minor
    oRev = App.Revision
    
    
    'stored like: "1.13.5"
    i = InStr(1, VerHTMLTxt, ".", vbTextCompare)
    j = InStr(i + 1, VerHTMLTxt, ".", vbTextCompare)
    
    On Error GoTo EH
    nMaj = Left$(VerHTMLTxt, i - 1)
    nMin = Mid$(VerHTMLTxt, i + 1, j - i - 1)
    nRev = Mid$(VerHTMLTxt, j + 1)
    On Error GoTo 0
    
    bHaveOld = False
    
    If nMaj > oMaj Then
        bHaveOld = True
    ElseIf nMaj = oMaj Then
        If nMin > oMin Then
            bHaveOld = True
        ElseIf nMin = oMin Then
            If nRev > oRev Then
                bHaveOld = True
            ElseIf nRev = oRev Then
                bHaveCurrent = True
            Else
                bHaveCurrent = False
            End If
        End If
    End If
End If


GetUpdateInfo = Not bError

Exit Function
EH:
GetUpdateInfo = False
End Function

Public Sub Set_Progress(sgProgress As Single, bCancel As Boolean)

If uCurrentState = uDownloadingVersion Then
    lblInfo.Caption = "Downloaded " & FormatNumber$(sgProgress, 2, vbTrue, vbFalse, vbFalse) & "%..."
    progUpdate.Value = sgProgress
End If

End Sub


'####################################################################################

Public Function Show_Update_Check(Optional ByVal bStealth As Boolean = False) As Boolean

Dim bHaveOld As Boolean, bHaveCurrent As Boolean
Dim sError As String
Dim obj_Disable As Object
Dim nMaj As Integer, nMin As Integer, nRev As Integer


SetInfo "Checking for an update..."
lblInfo.Caption = "Checking latest version..."

If Not bStealth Then
    ShowWindow
    Me.Refresh
End If
pbStealth = bStealth

Set obj_Disable = frmMain.mnuOnline

uCurrentState = uCheckingVersion

If GetUpdateInfo(bHaveOld, bHaveCurrent, obj_Disable, nMaj, nMin, nRev, sError) Then
    Show_Update_Check = True
    
    'reset our check thing
    modSettings.LastUpdate = Date
    
    If bHaveOld Then
        If bStealth Then ShowWindow 'show now
        
        lblInfo.Caption = "Update found - " & nMaj & Dot & nMin & Dot & nRev & vbNewLine & "Select Next to download"
        SetInfo "Newer Version Found"
        uCurrentState = uWaitingForDownloadConfirmation
        cmdNext.Enabled = True
        cmdNext.Default = True
        SetcmdPrevWidth False
        cmdPrev.Caption = "Cancel"
        cmdPrev.Visible = True
        cmdPrev.Enabled = True
        lblState.Caption = vbNullString
    Else
        If bStealth Then
            uCurrentState = uNone
            ForceUnload
        Else
            uCurrentState = uWaitingForDownloadConfirmation
            lblInfo.Caption = "No newer versions available (Latest Version: " & nMaj & Dot & nMin & Dot & nRev & ")" & vbNewLine & _
                "Select Next to download anyway."
            
            If bHaveCurrent Then
                lblState.Caption = vbNullString
            Else
                lblState.Caption = "You have a newer version, how is this possible?!"
            End If
            SetInfo "Communicator is up to date"
            
            SetcmdPrevWidth False
            cmdPrev.Caption = "Finish"
            cmdPrev.Visible = True
            cmdPrev.Enabled = True
            cmdPrev.Default = True
            cmdNext.Caption = "Next"
            cmdNext.Enabled = True
        End If
    End If
Else
    If LenB(sError) Then
        If modFTP.bCancelFTP Then
            lblInfo.Caption = sError
        Else
            lblInfo.Caption = "Error - " & sError
        End If
    Else
        lblInfo.Caption = "Error - The website may be offline...?"
        'IIf(LenB(VerHTMLTxt) > 0, " (Version Received: '" & VerHTMLTxt & "')", vbNullString), TxtError, True
    End If
    
    SetInfo "Couldn't Check for an Update"
    
    lblState.Caption = vbNullString
    
    cmdNext.Caption = "Finish"
    cmdNext.Enabled = True
    cmdNext.Default = True
    uCurrentState = uNone
End If

Set obj_Disable = Nothing

End Function

Private Sub Start_Download()
Dim sError As String, LFile As String

cmdPrev.Visible = False

SetInfo "Downloading Latest Version, Please Wait..."

uCurrentState = uDownloadingVersion
ShowprogUpdate

lblInfo.Caption = "Downloaded 0%"


cmdPrev.Caption = "Cancel"
SetcmdPrevWidth False
cmdPrev.Enabled = True
cmdPrev.Visible = True

If modFTP.DownloadLatest(frmMain.mnuOnline, sError) Then
    SetInfo "Download Successful"
    progUpdate.Visible = False
    
    lblInfo.Caption = "Extracting Communicator..."
    lblState.Caption = vbNullString
    
    uCurrentState = uExtractingVersion
    
    On Error GoTo EH
    LFile = AppPath() & modFTP.Communicator_File 'new zip
    FileCopy modFTP.FTP_Comm_Exe_File, LFile 'move from d/l location to ^
    
    
    If ZipFileExtract(LFile) Then 'extract new one
        lblInfo.Caption = "To complete the update, Communicator must be restarted." & vbNewLine & _
            "Restart Now?"
        
        uCurrentState = uWaitingForRestartConfirmation
        cmdPrev.Enabled = True
        cmdPrev.Caption = "Restart Communicator"
        SetcmdPrevWidth True
        cmdPrev.Visible = True
        cmdPrev.Default = True
        cmdNext.Enabled = True
        cmdNext.Caption = "Later"
    Else
        'errors all ready displayed
        cmdNext.Enabled = True
        cmdNext.Caption = "Finish"
        cmdNext.Default = True
        uCurrentState = uNone
    End If
Else
    lblInfo.Caption = "Error in download" & IIf(LenB(sError), " - " & sError, vbNullString)
    cmdNext.Enabled = True
    cmdNext.Caption = "Finish"
    cmdNext.Default = True
    uCurrentState = uNone
    lblState.Caption = vbNullString
    SetInfo "Download Stopped"
    cmdPrev.Visible = False
End If


Exit Sub
EH:
SetInfo "Error!"
lblInfo.Caption = "Error moving the newly downloaded zip:" & vbNewLine & Err.Description
cmdPrev.Visible = False
cmdNext.Caption = "Finish"
uCurrentState = uNone
cmdNext.Enabled = True
cmdPrev.Enabled = False
cmdNext.Default = True
End Sub

Private Sub Restart_Communicator()

uCurrentState = uRestarting

'open it
If modVars.OpenNewCommunicator("/killold /forceopen") Then
    lblInfo.Caption = "Opened New Communicator"
    SetInfo "Goodbye"
    ExitProgram
Else
    GoTo New_Shell_EH
End If


Exit Sub
New_Shell_EH:
SetInfo "Error!"
lblInfo.Caption = "Error opening new Communicator" & vbNewLine & Err.Description

OpenFolder vbNormalFocus, AppPath()

uCurrentState = uNone
cmdNext.Enabled = True
cmdNext.Caption = "Finish"
cmdPrev.Enabled = False
cmdNext.Default = True
End Sub

'######################################################
'######################################################
'######################################################

Private Function ZipFileExtract(ByVal zFile As String) As Boolean
Dim sFile As String

'ZFILE = APPPATH() & COMMUNICATOR.ZIP


'############################################################
'rename current as old
On Error GoTo Rename_Old_EH
sFile = AppPath() & App.EXEName

If FileExists(sFile & " Old.exe") Then 'in case another one exists
    On Error Resume Next
    Kill sFile & " Old.exe"
End If

On Error GoTo Rename_Old_EH
Name (sFile & ".exe") As (sFile & " Old.exe") 'rename current to Old
'############################################################

If ExtractZip(zFile, AppPath()) Then
    'extract to %AppPath%\
    
    If FileExists(sFile & ".exe") Then
        ZipFileExtract = True
        
        lblInfo.Caption = "Extracted Communicator"
        
        On Error Resume Next
        Kill zFile 'kill zip
    Else
        GoTo Extract_EH
    End If
Else
    GoTo Extract_EH
End If


Exit Function
Extract_EH:
lblInfo.Caption = "Error extracting zip file" & vbNewLine & "The folder containing the zip file's been opened, see if you can extract it yourself"
SetInfo "Error!"
GoTo General_EH

Rename_Old_EH:
SetInfo "Error!"
lblInfo.Caption = "Error overwriting the current Communicator"

General_EH:
OpenFolder vbNormalFocus, AppPath()

uCurrentState = uNone
cmdNext.Enabled = True
cmdNext.Caption = "Finish"
cmdPrev.Enabled = False
End Function

'##########################################################################
'##########################################################################
'##########################################################################

'Private Sub ZipO_Status(Text As String)
'AddConsoleText "Zip Status: " & Text
'End Sub
'
'Private Sub ZipO_ZipError(Number As eZipError, Description As String)
'AddConsoleText "Zip Object Error: " & Description & vbNewLine & Space$(modConsole.IndentLevel + 18) & _
'                "Number: " & CStr(Number)
'End Sub

Private Function ExtractZip(ByVal ZipFile As String, ByVal Path As String) As Boolean

If ExtractzlibDll() Then
    
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    
    On Error GoTo EH
    Set ZipO = New clsZipExtraction
    
    On Error GoTo EH
    
    If ZipO.OpenZip(ZipFile) Then
        ExtractZip = ZipO.Extract(Path, True, True)
    Else
        ExtractZip = False
    End If
    
    ZipO.CloseZip
    
    Set ZipO = Nothing
    
Else
    ExtractZip = False
End If

Exit Function
EH:
ExtractZip = False
Set ZipO = Nothing
End Function

Private Function ExtractzlibDll() As Boolean

Dim WinPath As String
Dim f As Integer

WinPath = Environ$("windir")
If Right$(WinPath, 1) <> "\" Then
    WinPath = WinPath & "\"
End If
WinPath = WinPath & "system32\zlib.dll"

If FileExists(WinPath) = False Then
    f = FreeFile()
    
    On Error GoTo EH
    
    Open WinPath For Output As #f
        Print #f, StrConv(LoadResData(101, "CUSTOM"), vbUnicode);
    Close #f
    ExtractzlibDll = FileExists(WinPath)
    
Else
    ExtractzlibDll = True
End If


Exit Function
EH:
ExtractzlibDll = False
End Function

Private Function RemovezlibDll() As Boolean

Dim WinPath As String

WinPath = Environ$("windir")
If Right$(WinPath, 1) <> "\" Then
    WinPath = WinPath & "\"
End If
WinPath = WinPath & "system32\zlib.dll"

If FileExists(WinPath) Then
    On Error GoTo EH
    
    Kill WinPath
    
    RemovezlibDll = Not FileExists(WinPath)
Else
    RemovezlibDll = True
End If

Exit Function
EH:
RemovezlibDll = False
End Function

'#################################################
'#################################################
'#################################################

Private Sub cmdNext_Click()
cmdNext.Enabled = False

If uCurrentState = uNone Then
    ForceUnload 'Finish
ElseIf uCurrentState = uWaitingForDownloadConfirmation Then
    Start_Download
ElseIf uCurrentState = uWaitingForRestartConfirmation Then
    modSettings.AddToTodoList "killold"
    ForceUnload
End If
End Sub
Private Sub cmdPrev_Click()
cmdPrev.Enabled = False

If uCurrentState = uWaitingForRestartConfirmation Then
    Restart_Communicator
ElseIf uCurrentState = uWaitingForDownloadConfirmation Then
    ForceUnload
Else
    CancelIfDownloading
End If
End Sub
Private Sub SetcmdPrevWidth(bWide As Boolean)
If bWide Then
    cmdPrev.width = 1935
    cmdPrev.Left = 2280
Else
    cmdPrev.width = 1455
    cmdPrev.Left = 2760
End If
End Sub

Private Sub Form_Load()

uCurrentState = uNone

cmdNext.Caption = "Next >"
cmdPrev.Caption = "< Previous"
cmdNext.Enabled = False
cmdPrev.Enabled = False: cmdPrev.Visible = False

ShowprogUpdate False

FormLoad Me, , False
Me.Visible = False
End Sub

Private Sub ShowprogUpdate(Optional ByVal bShow As Boolean = True)
progUpdate.Visible = bShow
lblState.Top = IIf(bShow, 1800, progUpdate.Top)
End Sub

Private Sub ShowWindow()

If Me.Visible = False Then
    modImplode.AnimateAWindow Me.hWnd, aRandom
    Me.Show vbModeless, frmMain
End If

End Sub

Private Sub ForceUnload()
uCurrentState = uNone
Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CancelIfDownloading

If uCurrentState <> uNone And Not modVars.Closing Then
    Cancel = True
    Beep
Else
    FormLoad Me, True, Not pbStealth
End If

End Sub
Private Sub CancelIfDownloading()
If uCurrentState = uDownloadingVersion Or uCurrentState = uCheckingVersion Then
    modFTP.bCancelFTP = True
End If
End Sub

Private Sub Form_Resize()
picInfo.width = Me.ScaleWidth
lnInfo.X2 = Me.ScaleWidth
lnButtons(0).X2 = lnInfo.X2
lnButtons(1).X2 = lnInfo.X2
End Sub
Public Sub SetInfo(ByVal sText As String)
Const iIndent = 120

With picInfo
    .Cls
    .CurrentX = iIndent
    .CurrentY = .height / 2 - .TextHeight(sText) / 2 - 60
    picInfo.Print sText
End With

End Sub
