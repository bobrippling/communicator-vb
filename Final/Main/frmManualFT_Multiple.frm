VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManualFT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual File Transfer"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   16875
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrStatus 
      Interval        =   10
      Left            =   480
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckTransfer 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstTransfers 
      Height          =   1035
      ItemData        =   "frmManualFT.frx":0000
      Left            =   120
      List            =   "frmManualFT.frx":0002
      TabIndex        =   13
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Frame fraServer 
      Caption         =   "Receive"
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   10440
      TabIndex        =   9
      Top             =   3240
      Width           =   4095
      Begin VB.PictureBox picServer 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   120
         ScaleHeight     =   1935
         ScaleWidth      =   3855
         TabIndex        =   10
         Top             =   240
         Width           =   3855
         Begin VB.CommandButton cmdCloseClients 
            Caption         =   "Disconnect Senders"
            Height          =   375
            Left            =   1920
            TabIndex        =   26
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "Open"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2760
            TabIndex        =   23
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtDir 
            Height          =   285
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   120
            Width           =   2535
         End
         Begin VB.CommandButton cmdDir 
            Caption         =   "..."
            Height          =   255
            Left            =   3360
            TabIndex        =   21
            Top             =   120
            Width           =   300
         End
         Begin VB.CommandButton cmdListen 
            Caption         =   "Listen"
            Height          =   375
            Left            =   0
            TabIndex        =   11
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label lblDirInfo 
            Alignment       =   2  'Center
            Caption         =   "This is where files are saved when received from the sender"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   0
            TabIndex        =   25
            Top             =   555
            Width           =   2535
         End
         Begin VB.Label lblDir 
            Caption         =   "Directory:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   120
            Width           =   735
         End
      End
   End
   Begin VB.Frame fraClient 
      Caption         =   "Send"
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   10440
      TabIndex        =   2
      Top             =   840
      Width           =   3975
      Begin VB.PictureBox picClient 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   3735
         TabIndex        =   3
         Top             =   360
         Width           =   3735
         Begin VB.CommandButton cmdIPChooser 
            Caption         =   "..."
            Height          =   255
            Left            =   3360
            TabIndex        =   17
            Top             =   0
            Width           =   300
         End
         Begin VB.Timer tmrSendFile 
            Enabled         =   0   'False
            Left            =   0
            Top             =   1320
         End
         Begin VB.TextBox txtPath 
            Height          =   285
            Left            =   960
            TabIndex        =   6
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   255
            Left            =   3360
            TabIndex        =   5
            Top             =   360
            Width           =   300
         End
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send File"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   4
            Top             =   840
            Width           =   1335
         End
         Begin projMulti.IPTextBox txtIP 
            Height          =   285
            Left            =   960
            TabIndex        =   18
            Top             =   0
            Width           =   2295
            _extentx        =   4048
            _extenty        =   503
         End
         Begin projMulti.VistaProg vprog 
            Height          =   225
            Left            =   0
            TabIndex        =   20
            Top             =   1440
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   397
         End
         Begin VB.Label lblIP 
            Caption         =   "IP Address:"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lblBrowse 
            Caption         =   "File Path:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblDragInfo 
            Caption         =   "Drag a file here to insert"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraDir 
      Caption         =   "Save Directory"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      Begin VB.PictureBox picDir 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   3735
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
   End
   Begin MSComctlLib.ListView lstConnections 
      Height          =   2100
      Left            =   0
      TabIndex        =   15
      Top             =   3960
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   3704
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Socket"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Connection State"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Remote"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "File Name"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Progress"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label lblConnections 
      Alignment       =   2  'Center
      Caption         =   "Connections"
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
      Left            =   0
      TabIndex        =   16
      Top             =   3720
      Width           =   4095
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
      TabIndex        =   14
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Loaded Window - Ready"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4200
      TabIndex        =   12
      Top             =   3360
      Width           =   3915
   End
End
Attribute VB_Name = "frmManualFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type tClient
    sFileName As String
    lFileSize As Long
    lBytesReceived As Long
    
    iFileNum As Integer
End Type
Private Clients() As tClient

'##################################################################

'general vars
Private LastSendPath As String, LastSavePath As String

'sender vars
Private Const PacketSize As Long = 1024& * 4&
Private lTransferStart As Long
Private iSendFileNum As Integer
Private bFTServer As Boolean, bSending As Boolean

Private Sub cmdBrowse_Click()
    Dim Path As String
    Dim bEr As Boolean
    
    frmMain.CommonDPath Path, bEr, "File to Send", "All Files (*.*)|*.*", LastSendPath, True
    
    If Not bEr Then
        If FileExists(Path) Then
            txtPath.Text = Path
            
            LastSendPath = Left$(Path, InStrRev(Path, "\") - 1)
        End If
    End If
End Sub

Private Sub cmdDir_Click()
    Dim Path As String
    
    Path = modVars.BrowseForFolder(LastSavePath, "File Directory", Me)
    
    If LenB(Path) Then
        If FileExists(Path, vbDirectory) Then
            If Right$(Path, 1) <> "\" Then Path = Path & "\"
            txtDir.Text = Path
            
            LastSavePath = Path
        Else
            txtDir.Text = LastSavePath
            'lblStatus.Caption = "Path doesn't exist"
        End If
    End If
    
End Sub

Private Sub lstConnections_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub txtDir_Change()
    cmdOpen.Enabled = CBool(LenB(txtDir.Text))
End Sub

'########################################################################
'########################################################################

Private Sub cmdIPChooser_Click()
    Dim IP As String
    
    IP = modVars.IPChoice(Me)
    
    If LenB(IP) Then
        txtIP.Text = IP
    End If
    
End Sub

Private Sub cmdOpen_Click()
    OpenFolder vbNormalFocus, txtDir.Text
    
    'modDisplay.ShowBalloonTip txtDir, "Directory Opened", _
    "Save directory has been opened"
    
End Sub

'#################################################################################
'#################################################################################

Private Sub Form_Load()
    Dim Path As String
    
    modLoadProgram.frmManualFT_Loaded = True
    
    lstConnections.ListItems.Add , , "0"
    FitTextInListView lstConnections, 0
    
    Erase Clients
    tmrSendFile.Enabled = False
    LastSendPath = vbNullString
    LastSavePath = vbNullString
    lTransferStart = 0
    iSendFileNum = 0
    bFTServer = False
    bSending = False
    
    Path = frmMain.FT_Path
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    
    
    If LenB(Path) <= 1 Then
        MsgBoxEx "Error Creating Received Files Folder", "Couldn't create the folder which files are received into", _
            vbExclamation, "Error", , , , , Me.hWnd
        
        Unload Me
    Else
        txtDir.Text = Path
        
        
        lblStatus.Caption = "Loaded Window" & vbNewLine & "Ready"
        
        
        txtIP.Text = IIf(Server, frmMain.SckLC.LocalHostName, frmMain.SckLC.RemoteHostIP)
        
        EnableOLEDragDrop True
        
        Call FormLoad(Me)
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call FormLoad(Me, True)
    modLoadProgram.frmManualFT_Loaded = False
End Sub

'##############################################################################
'##############################################################################

Private Sub txtIP_Change()
    cmdConnect.Enabled = LenB(txtIP.Text) And (bFTServer = False)
    
    txtIP.ShowIPBalloonTip
End Sub

Private Sub txtPath_Change()
    Dim Path As String
    
    If sckTransfer(0).State = sckConnected Then
        Path = Trim$(txtPath.Text)
        
        
        txtPath.Text = Trim$(Path)
        txtPath.Selstart = Len(Path)
        cmdSend.Enabled = CBool(LenB(Path))
        
        If FileExists(Path, vbNormal) Then
            
            If Len(Path) > 4 Then
                If Mid$(Path, Len(Path) - 3, 1) = Dot Then
                    
                    modDisplay.ShowBalloonTip txtPath, "File exists", _
                        "The file exists, feel free to send"
                    
                End If
            End If
        ElseIf LenB(Path) Then
            modDisplay.ShowBalloonTip txtPath, "File doesn't exist", _
            "It doesn't exist, you crazy man", TTI_WARNING
            
        End If
        
    End If
    
End Sub

Private Sub ResetProgbar(Optional iVal As Single = 0)
    vprog.Value = iVal
End Sub

Public Property Let FilePath(sPath As String)
    txtPath.Text = sPath
    cmdSend.Enabled = FileExists(sPath) And sckTransfer(0).State = sckConnected
    lblStatus.Caption = "File ready to be sent"
End Property

Private Sub OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'If pftStatus = Connected Then
    If Data.Files.Count > 1 Then
        MsgBoxEx "You can only send one file at a time", _
        "Only one file may be sent at once - you dragged two or more files onto the area", _
            vbExclamation, "Error"
    Else
        txtPath.Text = Data.Files(1)
        txtPath.Selstart = Len(txtPath.Text)
    End If
    
    txtPath.Enabled = (sckTransfer(0).State = sckConnected)
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

'################################################################################################
'################################################################################################
'################################################################################################
'################################################################################################

Private Sub cmdListen_Click()
    Dim sMsg As String
    
    sckTransfer(0).Close
    If sckTransfer(0).State <> sckListening Then
        sckTransfer(0).LocalPort = modPorts.FTPort
        sckTransfer(0).RemotePort = 0
        
        On Error GoTo EH
        sckTransfer(0).Listen
        
        lblStatus.Caption = "Awaiting Connection..."
        cmdListen.Caption = "Stop Listening"
        bFTServer = True
        cmdDisconnect.Enabled = True
    Else
        'listening, stop
        
        lblStatus.Caption = "Stopped Listening"
        cmdListen.Caption = "Listen"
        bFTServer = False
        cmdDisconnect.Enabled = False
    End If
    
    Exit Sub
EH:
    lblStatus.Caption = Err.Description
    bFTServer = False
    cmdDisconnect.Enabled = False
End Sub
Private Sub cmdConnect_Click()
    Dim IP As String
    
    cmdConnect.Enabled = False
    cmdConnect.Caption = "Connecting..."
    
    IP = txtIP.Text
    bFTServer = False
    
    If LenB(IP) = 0 Then
        tmrSendFile.Enabled = False
        
        sckTransfer(0).Close
        DoEvents
        
        sckTransfer(0).RemoteHost = IP
        sckTransfer(0).RemotePort = modPorts.FTPort
        
        On Error GoTo EH
        sckTransfer(0).Connect
        cmdDisconnect.Enabled = True
        cmdConnect.Enabled = False
    Else
        cmdDisconnect.Enabled = False
        cmdConnect.Enabled = True
    End If
    
    Exit Sub
EH:
    lblStatus.Caption = Err.Description
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
End Sub
Private Sub cmdDisconnect_Click()
    'close all connections
    Dim i As Integer
    
    For i = 0 To sckTransfer.UBound
        sckTransfer_Close i
        
        If i > 0 Then
            Unload sckTransfer(i)
        End If
    Next i
    
    bFTServer = False
    
End Sub

Private Sub cmdSend_Click()
    Dim sFile As String, sFileName As String
    
    If sckTransfer(0).State = sckConnected Then
        
        
        If bSending Then
            'cancel
            tmrSendFile.Enabled = False
            tmrSendFile.Interval = 0
            
            Close #iSendFileNum ' close file
            iSendFileNum = 0 ' set file number to 0, timer will exit if another timer event
            
            sckTransfer(0).Close
            
            cmdSend.Caption = "Send"
            bSending = False
        Else
            'start to send
            txtPath.Enabled = False
            EnableOLEDragDrop False
            cmdBrowse.Enabled = False
            
            sFile = Trim$(txtPath.Text)
            If FileExists(sFile) Then
                
                bSending = True
                lTransferStart = GetTickCount()
                
                SendStartOfFile GetFileName(sFile)
                
            Else
                lblStatus.Caption = "Error - File Doesn't Exist"
            End If
        End If
    Else
        lblStatus.Caption = "Error - Not Connected"
        cmdSend.Enabled = False
    End If
    
End Sub
Private Sub SendStartOfFile(sFileName As String)
    Dim Buffer() As Byte, P As Long
    
    iSendFileNum = FreeFile()
    Open sFileName For Binary Access Read Lock Write As #iSendFileNum
    
    ReDim Buffer(lngMIN(LOF(iSendFileNum), PacketSize) - 1)
    
    Get iSendFileNum, , Buffer ' read data
    
    sckTransfer(0).SendData CStr(LOF(iSendFileNum)) & ","   ' send the file size
    sckTransfer(0).SendData GetFileName(sFileName) & ":"       ' send the file name
    sckTransfer(0).SendData Buffer                             ' send first packet
    
    Erase Buffer
    
    lTransferStart = GetTickCount()
End Sub

Private Sub sckTransfer_Close(Index As Integer)
    On Error Resume Next
    
    sckTransfer(Index).Close
    
    If bFTServer Then
        If Clients(Index).iFileNum > 0 Then
            Close #Clients(Index).iFileNum
        End If
        
        
        If Clients(Index).lBytesReceived < Clients(Index).lFileSize Then
            Kill LastSavePath & Clients(Index).sFileName
            
            Me.lstConnections.ListItems(Index + 1).SubItems(4) = "Incomplete - File Deleted"
        Else
            Me.lstConnections.ListItems(Index + 1).SubItems(4) = "Transfer Complete"
        End If
        FitTextInListView Me.lstConnections, 4, , Index + 1
        
        ResetClient Index
        
    ElseIf Index = 0 Then
        cmdConnect.Caption = "Connect"
        cmdConnect.Enabled = True
    End If
    
End Sub

Private Sub sckTransfer_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim K As Integer
    
    
    For K = 1 To sckTransfer.UBound
        If sckTransfer(K).State = sckClosed Then Exit For
    Next K
    
    If K = sckTransfer.UBound + 1 Then
        'need a new socket
        Load sckTransfer(K)
        ReDim Preserve Clients(K)
        
        lstConnections.ListItems.Add , , CStr(K)
    End If
    
    sckTransfer(K).Close
    sckTransfer(K).accept requestID
    
    If LenB(sckTransfer(K).RemoteHost) = 0 Then
        Me.lstConnections.ListItems(K + 1).SubItems(2) = sckTransfer(K).RemoteHostIP
    Else
        Me.lstConnections.ListItems(K + 1).SubItems(2) = sckTransfer(K).RemoteHost
    End If
    
    FitTextInListView Me.lstConnections, 2, , K + 1
End Sub

Private Sub sckTransfer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim sData As String, lPos As Long, lPos2 As Long
    
    sckTransfer(Index).GetData sData, vbString, bytesTotal
    
    If Clients(Index).lFileSize = 0 Then
        If InStr(1, sData, ":") > 0 Then
            lPos = InStr(1, sData, ",")
            
            On Error GoTo start_EH
            Clients(Index).lFileSize = val(Left$(sData, lPos - 1))
            
            lPos2 = InStr(lPos, sData, ":")
            Clients(Index).sFileName = Mid$(sData, lPos + 1, (lPos2 - lPos) - 1)
            
            
            On Error GoTo save_EH
            Clients(Index).iFileNum = FreeFile()
            Open LastSavePath & Clients(Index).sFileName For Binary Access Write Lock Write As #Clients(Index).iFileNum
            
            sData = Mid$(sData, lPos2 + 1)
            
            Me.lstConnections.ListItems(Index + 1).SubItems(3) = Clients(Index).sFileName
            FitTextInListView Me.lstConnections, 3, , Index + 1
        End If
    End If
    
    
    If LenB(sData) > 0 Then
        Clients(Index).lBytesReceived = Clients(Index).lBytesReceived + Len(sData)
        Put #Clients(Index).iFileNum, , sData
        
        Me.lstConnections.ListItems(Index + 1).SubItems(4) = Format$(Clients(Index).lBytesReceived / Clients(Index).lFileSize * 100#, "#0.00") & "%"
        FitTextInListView Me.lstConnections, 4, , Index + 1
        
        If Clients(Index).lBytesReceived >= Clients(Index).lFileSize Then
            sckTransfer_Close Index
        End If
    End If
    
    
    Exit Sub
    
start_EH:
    ResetClient Index
    Me.lstConnections.ListItems(Index + 1).SubItems(4) = "Error Starting"
    sckTransfer_Close Index
    Exit Sub
    
save_EH:
    ResetClient Index
    Me.lstConnections.ListItems(Index + 1).SubItems(4) = "Error Saving File"
    sckTransfer_Close Index
    Exit Sub
    
End Sub

Private Sub sckTransfer_Connect(Index As Integer)
    If Not bFTServer Then
        cmdConnect.Caption = "Connected"
    End If
End Sub
Private Sub SckSendFile_SendComplete()
    If bSending Then
        ' can't call SendData here, so enable timer to do it...
        tmrSendFile.Enabled = False
        tmrSendFile.Interval = 1
        tmrSendFile.Enabled = True
    End If
End Sub

Private Sub sckTransfer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckTransfer_Close Index
    If Index = 0 Then
        If bSending Then
            lblStatus.Caption = "Error - " & Description
        End If
    Else
        ResetClient Index
    End If
End Sub

Private Sub tmrSendFile_Timer()
    Dim Buffer() As Byte, BuffSize As Long
    
    tmrSendFile.Enabled = False
    If iSendFileNum <= 0 Or sckTransfer(0).State <> sckConnected Then Exit Sub
    
    
    If Loc(iSendFileNum) >= LOF(iSendFileNum) Then ' FILE COMPLETE
        Close #iSendFileNum ' close file
        iSendFileNum = 0 ' set file number to 0, timer will exit if another timer event
    Else
        'if the remaining size in the file is smaller then PacketSize, the read only whatever is left
        BuffSize = lngMIN(LOF(iSendFileNum) - Loc(iSendFileNum), PacketSize)
        
        ReDim Buffer(BuffSize - 1) ' resize buffer
        Get iSendFileNum, , Buffer ' read data
        sckTransfer(0).SendData Buffer ' send data
        
        'Show progress
        lblStatus.Caption = Format$(Loc(iSendFileNum) / CDbl(LOF(iSendFileNum)), "%") & " Complete"
        vprog.Value = 100 * Loc(iSendFileNum) / CDbl(LOF(iSendFileNum))
    End If
    
End Sub
Private Function lngMIN(ByVal L1 As Long, ByVal L2 As Long) As Long
    If L1 < L2 Then
        lngMIN = L1
    Else
        lngMIN = L2
    End If
End Function

Private Sub ResetClient(Index As Integer)
    
    Clients(Index).iFileNum = 0
    Clients(Index).lBytesReceived = 0
    Clients(Index).lFileSize = 0
    Clients(Index).sFileName = vbNullString
    
End Sub

Private Sub tmrStatus_Timer()
    Dim K As Long, TmpStr As String
    
    For K = 0 To sckTransfer.UBound
        TmpStr = GetSckState(sckTransfer(K).State)
        
        If Me.lstConnections.ListItems(K + 1).SubItems(1) <> TmpStr Then
            Me.lstConnections.ListItems(K + 1).SubItems(1) = TmpStr
            FitTextInListView Me.lstConnections, 1, , K + 1
        End If
    Next K
End Sub

Private Function GetSckState(iState As MSWinsockLib.StateConstants) As String
    
    GetSckState = Choose(CInt(iState) + 1, "Closed", "Open", "Listening", "Connection pending", _
        "Resolving host", "Host resolved", "Connecting", "Connected", "Server is disconnecting", "Error")
    
End Function

Private Sub FitTextInListView(LV As ListView, ByVal Column As Integer, Optional ByVal Text As String, Optional ByVal ItemIndex As Long = -1)
    Dim TLen As Single, CapLen As Single
    
    CapLen = Me.TextWidth(LV.ColumnHeaders(Column + 1).Text) + 195
    
    If ItemIndex >= 0 Then
        If ItemIndex = 0 Then
            TLen = Me.TextWidth(LV.ListItems(ItemIndex).Text)
        Else
            TLen = Me.TextWidth(LV.ListItems(ItemIndex).SubItems(Column))
        End If
    Else
        TLen = Me.TextWidth(Text)
    End If
    
    TLen = TLen + 195
    
    If CapLen > TLen Then TLen = CapLen
    
    If LV.ColumnHeaders(Column + 1).width < TLen Then LV.ColumnHeaders(Column + 1).width = TLen
End Sub

