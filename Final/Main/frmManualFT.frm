VERSION 5.00
Begin VB.Form frmManualFT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual File Transfer"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDir 
      Caption         =   "Save Folder"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   4200
      TabIndex        =   13
      Top             =   480
      Width           =   3975
      Begin VB.PictureBox picDir 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   3735
         TabIndex        =   14
         Top             =   240
         Width           =   3735
         Begin VB.CommandButton cmdOpen 
            Caption         =   "Open"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2760
            TabIndex        =   19
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtDir 
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   120
            Width           =   2535
         End
         Begin VB.CommandButton cmdDir 
            Caption         =   "..."
            Height          =   255
            Left            =   3360
            TabIndex        =   17
            Top             =   120
            Width           =   300
         End
         Begin VB.Label lblDirInfo 
            Alignment       =   2  'Center
            Caption         =   "This is where files are saved when received from the sender"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   0
            TabIndex        =   18
            Top             =   550
            Width           =   2535
         End
         Begin VB.Label lblDir 
            Caption         =   "Directory:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.Frame fraFT 
      Caption         =   "Sendage"
      Height          =   1575
      Left            =   4200
      TabIndex        =   20
      Top             =   2040
      Width           =   3975
      Begin VB.PictureBox picFT 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   3735
         TabIndex        =   21
         Top             =   240
         Width           =   3735
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send File"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   25
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   255
            Left            =   3360
            TabIndex        =   24
            Top             =   120
            Width           =   300
         End
         Begin VB.TextBox txtPath 
            Height          =   285
            Left            =   720
            TabIndex        =   22
            Top             =   120
            Width           =   2535
         End
         Begin projMulti.VistaProg progBar 
            Height          =   225
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   397
         End
         Begin VB.Label lblDragInfo 
            Caption         =   "Drag a file here to insert"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label lblBrowse 
            Caption         =   "File Path:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.Frame fraCmds 
      Caption         =   "Connectage"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox picClear 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2400
         ScaleHeight     =   495
         ScaleWidth      =   1455
         TabIndex        =   10
         Top             =   2880
         Width           =   1455
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.ListBox lstTransfers 
         Height          =   840
         ItemData        =   "frmManualFT.frx":0000
         Left            =   120
         List            =   "frmManualFT.frx":0002
         TabIndex        =   9
         Top             =   2040
         Width           =   3735
      End
      Begin VB.PictureBox picCmds 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   3735
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         Begin VB.CommandButton cmdIPChooser 
            Caption         =   "..."
            Height          =   255
            Left            =   3360
            TabIndex        =   4
            Top             =   0
            Width           =   300
         End
         Begin projMulti.IPTextBox txtIP 
            Height          =   285
            Left            =   960
            TabIndex        =   3
            Top             =   0
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   503
         End
         Begin VB.CommandButton cmdListen 
            Caption         =   "Listen"
            Height          =   375
            Left            =   0
            TabIndex        =   5
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "Connect"
            Height          =   375
            Left            =   1920
            TabIndex        =   6
            Top             =   480
            Width           =   1695
         End
         Begin VB.CommandButton cmdDisconnect 
            Caption         =   "Disconnect"
            Height          =   375
            Left            =   840
            TabIndex        =   7
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lblIP 
            Caption         =   "IP Address:"
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Label lblTransfers 
         Alignment       =   2  'Center
         Caption         =   "Transfers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   3735
      End
   End
   Begin projMulti.ucFileTransfer ucFT 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Loaded Window - Ready"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   4035
   End
   Begin VB.Menu mnuTransfer 
      Caption         =   "Transfer Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuTransferOpen 
         Caption         =   "Open File"
      End
      Begin VB.Menu mnuTransferFolder 
         Caption         =   "Open Containing Folder"
      End
      Begin VB.Menu mnuTransferRemove 
         Caption         =   "Remove from List"
      End
   End
End
Attribute VB_Name = "frmManualFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'need to enable menu in Cmds()

Private pftStatus As eStatus
Private LastCDPath As String
Private lTransferStart As Long

Private Const FileTransfer_Shortcut As String = " (Ctrl+T)"

Private bCancelButton As Boolean, bCanceled As Boolean

Public Property Let ftStatus(ByVal newValue As eStatus)
pftStatus = newValue
FTCmds

'If pftStatus = Connected Then
'    EnableOLEDragDrop
'Else
'    EnableOLEDragDrop False
'End If

End Property

'Private Sub chkAutoAccept_Click()
'bFT_AutoAccept = CBool(chkAutoAccept.Value)
'End Sub

Private Sub cmdBrowse_Click()
Dim Path As String
Dim bEr As Boolean

frmMain.CommonDPath Path, bEr, "File to Send", "All Files (*.*)|*.*", LastCDPath, True

If Not bEr Then
    If FileExists(Path) Then
        
        txtPath.Text = Path
        txtPath.Selstart = Len(txtPath.Text)
        
        LastCDPath = Left$(Path, InStrRev(Path, "\") - 1)
        
    End If
End If

End Sub

Private Sub cmdClear_Click()

Me.lstTransfers.Clear

Erase modVars.TransferFilePaths
modVars.nFilePaths = 0

cmdClear.Enabled = False

End Sub

Private Sub AddTransfer(sTransferName As String, sFilePath As String, bReceived As Boolean)
Dim sInfo As String

lstTransfers.AddItem IIf(bReceived, "Received ", "Sent ") & sTransferName
If lstTransfers.ListCount > 0 Then
    lstTransfers.ListIndex = lstTransfers.ListCount - 1
End If

cmdClear.Enabled = True

sInfo = "File Transfer - " & IIf(bReceived, "received ", "sent ") & sTransferName

modSpeech.Say sInfo
AddConsoleText sInfo
modLogging.addToActivityLog sInfo

ReDim Preserve TransferFilePaths(modVars.nFilePaths)
TransferFilePaths(nFilePaths).sName = sFilePath
TransferFilePaths(nFilePaths).bReceived = bReceived
nFilePaths = nFilePaths + 1

End Sub

Private Sub cmdConnect_Click()
Dim IP As String

cmdConnect.Enabled = False

IP = txtIP.Text
If LenB(IP) Then
    ucFT.Connect IP, FTPort
    ftStatus = Connecting
    lblStatus.Caption = "Connecting... (Port " & modPorts.FTPort & ")"
Else
    ftStatus = Idle
End If

End Sub

Private Sub cmdDir_Click()
Dim Path As String

Path = modVars.BrowseForFolder(AppPath(), "File Directory", Me)

If LenB(Path) Then
    If FileExists(Path, vbDirectory) Then
        txtDir.Text = Path
        ucFT.SaveDir = Path
    Else
        txtDir.Text = vbNullString
        ucFT.SaveDir = vbNullString
    End If
End If

End Sub

Private Sub cmdDisconnect_Click()
ucFT.Disconnect
'ftStatus = Idle
End Sub

Private Sub cmdIPChooser_Click()
Dim IP As String

IP = modVars.IPChoice(Me)

If LenB(IP) Then
    txtIP.Text = IP
End If

End Sub

Private Sub cmdListen_Click()
Dim sMsg As String

ftStatus = Idle

If ucFT.Listen(FTPort) Then
    ftStatus = Listening
    sMsg = frmMain.LastName & " is listening for a file transfer, port " & modPorts.FTPort & FileTransfer_Shortcut
    
    SendInfoMessage sMsg
    AddText sMsg, , True
    
    lblStatus.Caption = "Awaiting Connection..."
End If

End Sub

Private Sub cmdOpen_Click()
OpenFolder vbNormalFocus, txtDir.Text

modDisplay.ShowBalloonTip txtDir, "Directory Opened", _
    "Save directory has been opened"

End Sub

Public Sub cmdSend_Click()
Dim File As String, fileName As String
Dim lStart As Long

If bCancelButton Then
    ucFT.Disconnect
    'FTCmds is called by Disconnect event
    
    setCancelButton False
    bCanceled = True
    
Else
    
    If pftStatus = Connected Then
        
        If Me.ucFT.iCurStatus = tReceiving Then
           lblStatus.Caption = "Can't send while receiving"
           Exit Sub
        End If
        
        
        setCancelButton True
        
        txtPath_Enabled False
        EnableOLEDragDrop False
        cmdBrowse.Enabled = False
        'cmdCancel.Enabled = True
        
        File = Trim$(txtPath.Text)
        If FileExists(File) Then
            
            fileName = Mid$(File, InStrRev(File, "\") + 1)
            
            'modDisplay.ShowBalloonTip txtPath.hWnd, "Sending File", _
                "The file is being sent..."
            
            lStart = GetTickCount()
            ucFT.SendFile File, fileName
            
            If bCanceled Then
                lblStatus.Caption = "Canceled"
                bCanceled = False
            Else
                lblStatus.Caption = "File Sent in " & _
                    FormatNumber$((GetTickCount() - lStart) / 1000, 2, vbTrue, vbFalse, vbFalse) & " seconds"
                
            End If
            
        Else
            lblStatus.Caption = "Error - File Doesn't Exist"
        End If
        
        txtPath_Enabled True
        cmdBrowse.Enabled = True
        EnableOLEDragDrop True
        'cmdCancel.Enabled = False
    Else
        lblStatus.Caption = "Error - Not Connected"
    End If
    
    setCancelButton False
End If

End Sub

Private Sub setCancelButton(bCancel As Boolean)

If bCancelButton = bCancel Then Exit Sub

If bCancel Then
    cmdSend.Caption = "Cancel"
Else
    cmdSend.Caption = "Send File"
End If
cmdSend.Enabled = True

bCancelButton = bCancel
End Sub

Private Sub Form_Load()
Dim Path As String
Dim i As Integer

For i = 0 To modVars.nFilePaths - 1
    lstTransfers.AddItem IIf(modVars.TransferFilePaths(i).bReceived, "Received ", "Sent ") & _
        GetFileName(modVars.TransferFilePaths(i).sName)
    
Next i

If lstTransfers.ListCount > 0 Then
    lstTransfers.ListIndex = lstTransfers.ListCount - 1
End If

modLoadProgram.frmManualFT_Loaded = True

ftStatus = Idle

cmdSend.Left = cmdBrowse.Left + cmdBrowse.width - cmdSend.width

Path = frmMain.FT_Path

'If FileExists(Path, vbDirectory) = False Then
'    On Error Resume Next
'    MkDir Path
'    lblStatus.Caption = "Created Save Directory"
'End If

ucFT.SaveDir = Path
ucFT.CloseOnReceived = False
txtDir.Text = ucFT.SaveDir 'vbnullstring if it doesn't exist

lblStatus.Caption = "Hi there"

txtIP.Text = IIf(Server, frmMain.SckLC.LocalHostName, frmMain.SckLC.RemoteHostIP)

'Me.chkAutoAccept.Value = Abs(CInt(bFT_AutoAccept))

EnableOLEDragDrop True

Call FormLoad(Me)


'for noobs
If Server Then
    cmdListen_Click
ElseIf LenB(txtIP.Text) Then
    cmdConnect_Click
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
ucFT.Disconnect
modLoadProgram.frmManualFT_Loaded = False
End Sub

Private Sub txtPath_Enabled(b As Boolean)
Const kDisabled As String = "Disabled"

txtPath.Enabled = b
If Not b Then
    If LenB(txtPath.Text) = 0 Then txtPath.Text = kDisabled
    
ElseIf txtPath.Text = kDisabled Then
    txtPath.Text = vbNullString
End If

End Sub
Private Sub FTCmds()
Const Cap = "File Transfer"

'txtPath.Text = vbNullString
setCancelButton False
bCanceled = False

'SendMessageByLong Me.hWnd, WM_SETICON, ICON_SMALL, _
    CLng(frmSystray.img16x16.ListImages(CInt(pftStatus) + 1).Picture.Handle)


If pftStatus = Connected Then
    cmdConnect.Enabled = False
    cmdListen.Enabled = False
    cmdDisconnect.Enabled = True
    txtIP.Enabled = False
    fraFT.Enabled = True
    cmdIPChooser.Enabled = False
    'cmdSend.Enabled = cbool(lenb
    txtPath_Change
    txtPath_Enabled True
    cmdBrowse.Enabled = True
    
    Me.Caption = Cap & " - Status: Connected"
    
ElseIf pftStatus = Connecting Then
    cmdConnect.Enabled = False
    cmdListen.Enabled = False
    cmdDisconnect.Enabled = True
    txtIP.Enabled = False
    fraFT.Enabled = False
    cmdIPChooser.Enabled = False
    cmdSend.Enabled = False
    txtPath_Enabled False
    cmdBrowse.Enabled = False
    
    Me.Caption = Cap & " - Status: Connecting"
    
ElseIf pftStatus = Idle Then
    cmdConnect.Enabled = LenB(txtIP.Text)
    cmdListen.Enabled = True
    cmdDisconnect.Enabled = False
    txtIP.Enabled = True
    fraFT.Enabled = False
    cmdIPChooser.Enabled = True
    cmdSend.Enabled = False
    txtPath_Enabled False
    cmdBrowse.Enabled = False
    
    lblStatus.Caption = "Disconnected"
    
    Me.Caption = Cap & " - Status: Idle"
    
ElseIf pftStatus = Listening Then
    cmdConnect.Enabled = False
    cmdListen.Enabled = False
    cmdDisconnect.Enabled = True
    txtIP.Enabled = False
    fraFT.Enabled = False
    cmdIPChooser.Enabled = False
    txtIP.Text = frmMain.SckLC.LocalHostName
    cmdSend.Enabled = False
    txtPath_Enabled False
    cmdBrowse.Enabled = False
    
    Me.Caption = Cap & " - Status: Listening"
    
End If

End Sub

Private Sub lstTransfers_DblClick()
mnuTransferOpen_Click
End Sub

Private Sub lstTransfers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sFile As String

If Button = vbRightButton Then
    sFile = lstTransfers.Text
    
    If LenB(sFile) Then
        mnuTransferOpen.Caption = "Open " & Mid$(sFile, InStr(1, sFile, vbSpace) + 1)
        'mnuTransferOpen.Enabled = Not Left$(sFile, 4) = "Sent"
        
        PopupMenu Me.mnuTransfer, , , , Me.mnuTransferOpen
    End If
End If

End Sub

Private Sub mnuTransferFolder_Click()
Dim sPath As String

sPath = GetSelectedFilePath()

If LenB(sPath) Then OpenFolder vbNormalFocus, , sPath

End Sub

Private Sub mnuTransferOpen_Click()
Dim sPath As String

sPath = GetSelectedFilePath()

If LenB(sPath) Then
    If OpenURL(sPath) <= 32 Then
        MsgBox "Error opening file", vbExclamation, "Error"
    End If
End If

End Sub

Private Function GetSelectedFilePath() As String
Dim sFile As String, sPath As String
Dim i As Integer

On Error GoTo EH

sFile = lstTransfers.Text
If Left$(sFile, 4) = "Sent" Then
    sPath = TransferFilePaths(lstTransfers.ListIndex).sName
Else
    sPath = ucFT.SaveDir & Mid$(sFile, InStr(1, sFile, vbSpace) + 1)
End If

GetSelectedFilePath = sPath

EH:
End Function

Private Sub mnuTransferRemove_Click()
Dim i As Integer, j As Integer
Dim sText As String

i = lstTransfers.ListIndex

If i > -1 Then
    sText = lstTransfers.Text
    lstTransfers.RemoveItem lstTransfers.ListIndex
    
    If lstTransfers.ListCount = 0 Then
        cmdClear.Enabled = False
        Erase modVars.TransferFilePaths
        modVars.nFilePaths = 0
    Else
        sText = "*" & Mid$(sText, InStr(1, sText, vbSpace) + 1)
        
        j = -1
        For i = 0 To modVars.nFilePaths - 1
            If modVars.TransferFilePaths(i).sName Like sText Then
                j = i
                Exit For
            End If
        Next i
        
        If j > -1 Then
            'found, remove
            
            For i = j To modVars.nFilePaths - 2
                modVars.TransferFilePaths(i) = modVars.TransferFilePaths(i + 1)
            Next i
            
            ReDim Preserve modVars.TransferFilePaths(modVars.nFilePaths - 2)
            modVars.nFilePaths = modVars.nFilePaths - 1
        End If
    End If
End If

End Sub

Private Sub txtDir_Change()
cmdOpen.Enabled = CBool(LenB(txtDir.Text))
End Sub

Private Sub txtIP_Change()
cmdConnect.Enabled = LenB(txtIP.Text) And (pftStatus = Idle)

txtIP.ShowIPBalloonTip
End Sub

Private Sub txtPath_Change()
Dim Path As String

If pftStatus = Connected Then
    Path = Trim$(txtPath.Text)
    
    
    'txtPath.Text = Trim$(Path)
    'txtPath.Selstart = Len(Path)
    
    setCancelButton False
    
    If FileExists(Path, vbNormal) Then
        cmdSend.Enabled = CBool(LenB(Path))
        cmdSend.Caption = "Send File"
    Else
        cmdSend.Enabled = False
        cmdSend.Caption = "File not found"
    End If
    
'        'If Len(Path) > 4 Then
'            'If Mid$(Path, Len(Path) - 3, 1) = Dot Then
'
'                'modDisplay.ShowBalloonTip txtPath, "File exists", _
'                    "The file exists, feel free to send"
'
'            'End If
'        'End If
'        If LenB(Path) Then
'            modDisplay.ShowBalloonTip txtPath, "File doesn't exist", _
'                "Are you trying to raise an error/commotion around here?", TTI_WARNING
'
'        End If
'    End If
End If

End Sub

Private Sub ucFT_Connected(IP As String)
ftStatus = Connected
lblStatus.Caption = "Connected to " & IP

modSpeech.Say "File Transfer Connection Established"
ResetProgbar
End Sub

Private Sub ucFT_ConnectionRequest(IP As String, bAccept As Boolean)
'Dim Ans As VbMsgBoxResult

'If chkAutoAccept.Value = 1 Then
    'Ans = vbYes
'Else
    'Ans = MsgBoxEx("Incomming Connection Request from " & IP & vbNewLine & _
        "Accept Connection?", IP & " is trying to connect to you, to start a file transfer" & _
        "Do you want to allow this connection?", _
        vbYesNo + vbQuestion, "Accept Connection?" _
        , , , frmMain.Icon.Handle)
'End If

'If Ans = vbYes Then
    bAccept = True
'Else
    'bAccept = False
'End If

ResetProgbar

End Sub

Private Sub ucFT_Diconnected()
ftStatus = Idle
ResetProgbar
End Sub

Private Sub ResetProgbar(Optional iVal As Single = 0)
progBar.Value = iVal
End Sub

Private Sub ucFT_Error(Description As String, ErrNo As eFTErrors)
ftStatus = Idle
lblStatus.Caption = Description
ResetProgbar
End Sub

Private Sub ucFT_ReceivedFile(sFileName As String)
Dim sTime As String, sName As String

sName = GetFileName(sFileName)
sTime = FormatNumber$((GetTickCount() - lTransferStart) / 1000, 2, vbTrue, vbFalse, vbFalse) & " seconds"


lblStatus.Caption = "Received File in " & sTime

AddTransfer sName, sFileName, True


lTransferStart = 0
ResetProgbar 100
End Sub

Private Sub ucFT_ReceivingFile(sFileName As String, ByVal BytesReceived As Long, ByVal BytesRemaining As Long, ByVal lTotalBytes As Long)
Dim Percent As Single

If lTransferStart = 0 Then lTransferStart = GetTickCount()

On Error Resume Next
'setCancelButton True
'only sender can cancel

Percent = 100 * BytesReceived / lTotalBytes

progBar.Value = Percent
lblStatus.Caption = "Receiving " & GetFileName(sFileName) & _
        " (" & FormatNumber$(BytesReceived / 1024, 2, vbTrue, vbFalse, vbFalse) & " KB) - " & _
        FormatNumber$(Percent, 2, vbTrue, vbFalse, vbFalse) & "%"


'lblStatus.Caption = "Receiving - " & GetFileName(sFileName) & _
        " (" & FormatNumber$(BytesReceived / 1024, 2, vbTrue, vbFalse, vbFalse) & " KB)"
'lblStatus.Caption = "Receiving - " & GetFileName(sFileName) & _
    " (" & Format$(BytesReceived / 1024, "0.00") & " KB)"


lblStatus.Refresh
'Me.Refresh

End Sub

Private Sub ucFT_SendingFile(sFileName As String, ByVal BytesSent As Long, ByVal BytesRemaining As Long, ByVal lTotalBytes As Long)
Dim Percent As Single

On Error Resume Next

Percent = 100 * (lTotalBytes - BytesRemaining) / lTotalBytes

progBar.Value = Percent
lblStatus.Caption = "Sending " & GetFileName(sFileName) & _
    " - " & FormatNumber$(Percent, 2, vbTrue, vbFalse, vbFalse) & "%"

End Sub

Private Sub ucFT_SentFile(sFileName As String)
'above is done after ucFT exits the SendFile() function
'buh ^^ ??

AddTransfer GetFileName(sFileName), sFileName, False

ResetProgbar 100
End Sub

Public Property Let FilePath(sPath As String)
Dim sMsg As String

txtPath.Text = sPath
txtPath.Selstart = Len(txtPath.Text)
cmdSend.Enabled = (pftStatus = Connected) And FileExists(sPath)
setCancelButton False
lblStatus.Caption = "File ready to be sent"


If pftStatus <> Connected Then ' if THIS FILE TRANSFER isn't connected...
    If Not Server Then
        sMsg = frmMain.LastName & " requests a file transfer host, port " & modPorts.FTPort & FileTransfer_Shortcut
        SendInfoMessage sMsg
        AddText sMsg, , True
    'Else
        'server, wait for the user to press host
    End If
End If

End Property

Private Sub OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

'If pftStatus = Connected Then
    If Data.Files.Count > 1 Then
        MsgBoxEx "You can only send one file at a time", _
            "Only one file may be sent at once - you dragged two or more files onto the area", _
            vbExclamation, "Error"
    Else
        'txtPath.Text = Data.Files(1)
        'txtPath.Selstart = Len(txtPath.Text)
        FilePath = Data.Files(1)
    End If
    
    txtPath_Enabled (pftStatus = Connected)
'End If

End Sub

Private Sub lblDragInfo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub picFT_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub fraFT_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub cmdBrowse_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub txtPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub lblBrowse_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub cmdSend_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub EnableOLEDragDrop(Optional bEn As Boolean = True)
Dim i As Integer

i = Abs(bEn)

Me.OLEDropMode = i
lblDragInfo.DragMode = i
picFT.OLEDropMode = i
fraFT.OLEDropMode = i
cmdBrowse.OLEDropMode = i
txtPath.OLEDropMode = i
lblBrowse.OLEDropMode = i
cmdSend.OLEDropMode = i

End Sub
