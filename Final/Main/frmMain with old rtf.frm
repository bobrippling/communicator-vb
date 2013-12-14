VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Communicator"
   ClientHeight    =   7320
   ClientLeft      =   75
   ClientTop       =   765
   ClientWidth     =   9435
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   9435
   Begin MSWinsockLib.Winsock SckLC 
      Left            =   4800
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrLog 
      Interval        =   60000
      Left            =   7080
      Top             =   5400
   End
   Begin VB.Timer tmrInactive 
      Interval        =   30000
      Left            =   6480
      Top             =   5400
   End
   Begin VB.Frame fraDev 
      Height          =   735
      Left            =   3360
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtSendTo 
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Text            =   "Send to: "
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cboDevCmd 
         Height          =   315
         ItemData        =   "frmMain.frx":0CCE
         Left            =   2160
         List            =   "frmMain.frx":0CD0
         TabIndex        =   23
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Timer tmrCanShake 
      Interval        =   5000
      Left            =   8280
      Top             =   4800
   End
   Begin VB.CommandButton cmdShake 
      Caption         =   "Shake"
      Height          =   375
      Left            =   8640
      TabIndex        =   21
      Top             =   3960
      Width           =   735
   End
   Begin VB.ComboBox cboWidth 
      Height          =   315
      Left            =   2040
      TabIndex        =   19
      Text            =   "1"
      Top             =   4440
      Width           =   735
   End
   Begin VB.PictureBox picColour 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   18
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton cmdCls 
      Caption         =   "Clear Board"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Timer tmrShake 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7680
      Top             =   4800
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "No"
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "Yes"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrHost 
      Interval        =   60000
      Left            =   7080
      Top             =   4800
   End
   Begin VB.TextBox txtDev 
      Height          =   285
      Left            =   3360
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton cmdDevSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog Cmdlg 
      Left            =   4200
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer tmrList 
      Interval        =   10000
      Left            =   6480
      Top             =   4800
   End
   Begin VB.ListBox lstComputers 
      Height          =   2790
      ItemData        =   "frmMain.frx":0CD2
      Left            =   120
      List            =   "frmMain.frx":0CD4
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Connection"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.ListBox lstConnected 
      Height          =   2790
      ItemData        =   "frmMain.frx":0CD6
      Left            =   1800
      List            =   "frmMain.frx":0CD8
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Connect"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   3195
      TabIndex        =   10
      Top             =   4920
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Host"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtOut 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Top             =   3960
      Width           =   4575
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   3960
      Width           =   615
   End
   Begin RichTextLib.RichTextBox rtfIn 
      Height          =   3495
      Left            =   3360
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0CDA
   End
   Begin MSWinsockLib.Winsock SockAr 
      Index           =   0
      Left            =   5400
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblColour 
      Caption         =   "Colour"
      Height          =   165
      Left            =   1440
      TabIndex        =   20
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label lblTyping 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   25
      Width           =   5775
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Communicator"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveCon 
         Caption         =   "Save Conversation..."
      End
      Begin VB.Menu mnuFileSaveDraw 
         Caption         =   "Save Drawing..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveSettings 
         Caption         =   "Save Settings"
      End
      Begin VB.Menu mnuFileLoadSettings 
         Caption         =   "Load Settings"
      End
      Begin VB.Menu mnuFileDelSettings 
         Caption         =   "Delete Settings"
      End
      Begin VB.Menu mnuFileSaveExit 
         Caption         =   "Save Setting On Exit"
         Checked         =   -1  'True
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileManual 
         Caption         =   "Manual Connect"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileRefresh 
         Caption         =   "Refresh Network List"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileClient 
         Caption         =   "Client Window"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFileInvite 
         Caption         =   "Invite"
         Shortcut        =   ^I
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOptionsWindow 
         Caption         =   "Options Window"
      End
      Begin VB.Menu Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsWindow2 
         Caption         =   "Window"
         Begin VB.Menu mnuOptionsFlashMsg 
            Caption         =   "Flash When Message Recieved"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsFlashInvert 
            Caption         =   "Invert Mode"
         End
         Begin VB.Menu mnuOptionsWindow2Sep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsWindow2Animate 
            Caption         =   "Animate Window"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsWindow2SingleClick 
            Caption         =   "Single Click Tray Icon"
         End
      End
      Begin VB.Menu mnuOptionsMessaging 
         Caption         =   "Messaging"
         Begin VB.Menu mnuOptionsBalloonMessages 
            Caption         =   "Balloon Messages"
         End
         Begin VB.Menu mnuOptionsMessagingLog 
            Caption         =   "Log Conversations"
         End
         Begin VB.Menu mnuOptionsMessagingSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsTimeStamp 
            Caption         =   "TimeStamp Messages"
         End
         Begin VB.Menu mnuOptionsTimeStampInfo 
            Caption         =   "Timestamp Information"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuOptionsMessagingSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsMessagingColours 
            Caption         =   "Allow Different Colours"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsMatrix 
            Caption         =   "Matrix Chat Mode"
         End
      End
      Begin VB.Menu mnuOptionsAdv 
         Caption         =   "Advanced"
         Begin VB.Menu mnuOptionsAdvPreset 
            Caption         =   "Preset Settings"
            Begin VB.Menu mnuOptionsAdvPresetServer 
               Caption         =   "Default Server Settings"
            End
            Begin VB.Menu mnuOptionsAdvPresetManual 
               Caption         =   "Default Manual Settings"
            End
            Begin VB.Menu mnuOptionsAdvPresetReset 
               Caption         =   "Reset Settings"
            End
         End
         Begin VB.Menu mnuOptionsXP 
            Caption         =   "XP Style Mode"
         End
         Begin VB.Menu mnuOptionsAdvInactive 
            Caption         =   "Inactivity Timer"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsHost 
            Caption         =   "Host Mode"
         End
         Begin VB.Menu mnuOptionsStartup 
            Caption         =   "Startup"
         End
      End
      Begin VB.Menu mnuOptionsSystray 
         Caption         =   "Systray"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuDev 
      Caption         =   "DevMode"
      Begin VB.Menu mnuDevForms 
         Caption         =   "Forms"
         Begin VB.Menu mnuDevForm 
            Caption         =   "DevForm"
         End
         Begin VB.Menu mnuDevDataForm 
            Caption         =   "Dev Data Form"
         End
      End
      Begin VB.Menu mnuDevDataCmds 
         Caption         =   "Data Commands"
         Begin VB.Menu mnuDevShowCmds 
            Caption         =   "Show Recieved Dev Commands"
         End
         Begin VB.Menu mnuDevShowAll 
            Caption         =   "Show All Unknown Data"
         End
         Begin VB.Menu mnuDevConsole 
            Caption         =   "Console/CommandLine Commands"
            Shortcut        =   {F3}
         End
      End
      Begin VB.Menu mnuDevAdvCmds 
         Caption         =   "Advanced Commands"
         Begin VB.Menu mnuDevPause 
            Caption         =   "Pause Timers"
         End
         Begin VB.Menu mnuDevSubClass 
            Caption         =   "SubClass"
         End
         Begin VB.Menu mnuDevEndOnClose 
            Caption         =   "End On Close"
         End
         Begin VB.Menu mnuDevAdvNullChar 
            Caption         =   "Use NullChar Seperator"
         End
      End
      Begin VB.Menu mnuDevMaintenance 
         Caption         =   "Maintenance"
         Begin VB.Menu mnuDevEnable 
            Caption         =   "Enable Menu"
         End
         Begin VB.Menu mnuOptionsMessagingClearTypeList 
            Caption         =   "Clear Typing List"
         End
      End
      Begin VB.Menu mnuDevCmds 
         Caption         =   "Cmds (Procedure)"
         Begin VB.Menu mnuDevCmdsP 
            Caption         =   "Idle"
            Index           =   0
         End
         Begin VB.Menu mnuDevCmdsP 
            Caption         =   "Connected"
            Index           =   1
         End
         Begin VB.Menu mnuDevCmdsP 
            Caption         =   "Connecting"
            Index           =   2
         End
         Begin VB.Menu mnuDevCmdsP 
            Caption         =   "Listening"
            Index           =   3
         End
      End
      Begin VB.Menu mnuDevSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDevOff 
         Caption         =   "Turn Off"
      End
      Begin VB.Menu mnuDevHelp 
         Caption         =   "Command Help"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuRtfPopup 
      Caption         =   "RtfPopup"
      Begin VB.Menu mnuRtfPopupSaveAs 
         Caption         =   "Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuRtfPopupCls 
         Caption         =   "Clear Screen"
      End
      Begin VB.Menu mnuRtfPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRtfPopupCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuRtfPopupDelSel 
         Caption         =   "Delete Selected Text"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuConsole 
      Caption         =   "Console"
      Begin VB.Menu mnuConsoleType 
         Caption         =   "Single Command"
      End
      Begin VB.Menu mnuConsoleTypeLots 
         Caption         =   "Mutiple Commands"
      End
      Begin VB.Menu mnuConsoleOff 
         Caption         =   "Turn Off"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private QuestionReply As Byte
Private CanShake As Boolean
'Private LastWndState As FormWindowStateConstants
Private Questioning As Boolean

Private InActiveTmr As Integer

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal Msg As Long, _
    wParam As Any, lParam As Any) As Long
    
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private LastName As String

Private Sub cboDevCmd_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboWidth_Change()
Dim sTmp As String
Dim nTmp As Integer

sTmp = Trim$(cboWidth.Text)

If Len(sTmp) > 0 Then
    
    On Error Resume Next
    nTmp = CInt(sTmp)
    On Error GoTo 0
    
    If (1 <= nTmp) And (nTmp <= 50) Then
        picDraw.DrawWidth = nTmp
    Else
        AddText "Please enter a Width between 1 and 50", TxtError, True
        cboWidth.Text = Trim$(Str(picDraw.DrawWidth))
    End If
    
End If
End Sub

Private Sub cboWidth_Click()
cboWidth_Change
End Sub

Private Sub cboWidth_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
If Connect() Then
    AddText "To save time, double click the listbox instead", , True
End If
End Sub

Public Function Connect(Optional ByVal Name As String = vbNullString) As Boolean
Dim sRemoteHost As String, Text As String

Connect = True

On Error GoTo EH

Call CleanUp

Cmds Connecting

sRemoteHost = Trim$(IIf(Name = vbNullString, lstComputers.List(lstComputers.ListIndex), Name))

SckLC.RemoteHost = sRemoteHost

If SckLC.RemoteHost = vbNullString Then
    AddText "Please select a computer to connect to", TxtError, True
    Cmds Idle
    Connect = False
    Exit Function
End If


SckLC.RemotePort = RPort

SckLC.LocalPort = LPort

Text = "Connecting to " & sRemoteHost & ":" & RPort & "..."

AddText Text, , True

AddConsoleText Text

SckLC.Connect 'try to connect


Exit Function
EH:
Call ErrorHandler(Err.Description, Err.Number, , True)
Connect = False
End Function

Public Sub CleanUp()
Dim n As Integer

Server = False

SckLC_Close 'we close it in case it was trying to connect

lstConnected.Clear
picDraw.Cls
lblTyping.Caption = vbNullString

ReDim Clients(0)

On Error Resume Next
'close and unload all previous sockets
For n = 1 To (SockAr.Count - 1)
    SockAr(n).Close
    Unload SockAr(n)
Next n

'Cmds Idle - no need - done in scklc_close

SocketCounter = 0

AddConsoleText "Cleaned Up"

'frmSystray.ShowBalloonTip "All Connections Closed", "Communicator", NIIF_INFO

End Sub

Public Sub cmdClose_Click()
Call CleanUp
AddText "Connection Closed", , True
End Sub

Private Sub cmdCls_Click()
picDraw.Cls
If Server Then
    DistributeMsg eCommands.Draw & "cls", -1
Else
    SendData eCommands.Draw & "cls"
End If
End Sub

Private Sub cmdDevSend_Click()
Dim dMsg As String, SendTo As String

On Error Resume Next
SendTo = Trim$(Right$(txtSendTo.Text, Len(txtSendTo.Text) - 9))
On Error GoTo 0

If Len(SendTo) <= 0 And Left$(cboDevCmd.Text, 1) <> "0" Then
    AddText "Please Select a computer to send to", TxtError, True
    Exit Sub
End If

If Left$(cboDevCmd.Text, 1) <> "0" Then
    dMsg = eCommands.DevSend & SendTo & "#" & Trim$(txtName.Text) & "@" & Left$(cboDevCmd.Text, 1) & txtDev.Text
    
    AddDevText vbNewLine & "DevMode, Sent:" & vbNewLine & _
        "To: " & SendTo & vbNewLine & _
        "Command: " & Right$(cboDevCmd.Text, Len(cboDevCmd.Text) - 4) & vbNewLine & _
        "Parameter(s): " & txtDev.Text & vbNewLine, True
Else
    dMsg = txtDev.Text
    
    AddDevText vbNewLine & "DevMode, Sent:" & vbNewLine & _
        "To: " & SendTo & vbNewLine & _
        "Message: " & txtDev.Text & vbNewLine, True
End If

If Server Then
    DistributeMsg dMsg, -1
Else
    SendData dMsg
End If

txtDev.Text = vbNullString
txtDev.SetFocus

'dmsg = eDevCmd & WhoTo & # & From & @ & Command

End Sub

Public Sub cmdListen_Click()
Call Listen
End Sub

Public Sub Listen(Optional ByVal DoEH As Boolean = True)

Call CleanUp

On Error GoTo EH

Cmds Listening

'SckLC(0) is the name of our Winsock ActiveX Control

SckLC.Close 'we close it in case it listening before


'txtPort is the textbox holding the Port number
SckLC.LocalPort = RPort  'set the port we want to listen to
                              '( the client will connect on this port too)
SckLC.RemotePort = LPort

On Error Resume Next
SckLC.Listen                'Start Listening

If SckLC.State <> sckListening Then GoTo EH

AddText "Listening...", , True

'frmSystray.ShowBalloonTip "Listening...", , NIIF_INFO, 1000

Server = True

Exit Sub
EH:
Call ErrorHandler(Err.Description, Err.Number, DoEH)

End Sub

Private Sub cmdReply_Click(Index As Integer)
QuestionReply = Index
cmdReply(0).Visible = False
cmdReply(1).Visible = False
End Sub

Private Sub cmdShake_Click()

If CanShake = False Then
    AddText "You cannot shake that often", TxtError, True
    Exit Sub
End If

If Server Then
    DistributeMsg eCommands.Shake, -1
Else
    SendData eCommands.Shake
End If

AddText "Shake Sent", TxtSent, True

CanShake = False
tmrCanShake.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If Shift = 1 Then
    If KeyCode = 223 Then
        If ConsoleShown Then
            ShowConsole False
        Else
            ShowConsole
        End If
        Pause 10
        
        KeyCode = 0
        Shift = 0
        'prevent beep
        
    End If
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If (UnloadMode = vbAppWindows) Or (UnloadMode = vbAppTaskManager) _
   Or (UnloadMode = vbFormOwner) Or (UnloadMode = vbFormMDIForm) Then Closing = True

If Closing Then
    If ConsoleShown Then
        ShowConsole False
    End If
    
    If mnuFileSaveExit.Checked Then modSettings.SaveSettings
    If InTray Then
        DoSystray False
    End If
    If Me.Visible Then ImplodeFormToMouse Me.hWnd
Else
    Cancel = True
    ShowForm False
End If

End Sub

Public Function Question(ByVal sMessage As String, Caller As Object) As VbMsgBoxResult

If Questioning Then
    Question = vbRetry
    AddText "Please Answer the Previous Question first", TxtError, True
    Exit Function
End If

Caller.Enabled = False
Questioning = True

AddText sMessage, TxtQuestion, True

QuestionReply = 3

If Me.Visible = False Then Me.ShowForm

cmdReply(0).Visible = True
cmdReply(1).Visible = True

cmdReply(1).Default = True
cmdReply(1).SetFocus

Do
    Pause 100
    
    If QuestionReply <> 3 Then
        Select Case QuestionReply
            Case 1
                Question = vbYes
            Case 0
                Question = vbNo
        End Select
    End If
    
Loop While QuestionReply = 3 And Not Closing

'if questionreply = 3

Caller.Enabled = True
Questioning = False

End Function

Private Sub Form_Unload(Cancel As Integer)
Dim Frm As Form

If bDevMode Then frmDev.CloseMe = True

For Each Frm In Forms
    If Frm.Name <> "frmMain" Then Unload Frm
    'Set Frm = Nothing
Next Frm


Dim f As Integer

f = FreeFile()
Open modVars.SafeFile For Output As #f
    Print #f, Str(SafeConfirm)
Close #f


'End
End Sub

Private Sub lstComputers_DblClick()
Call Connect
End Sub

'Private Sub lstComputers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = vbRightButton Then
    'addtext "Name: " & lstcomputers.Text & " Description: " & lstcomputers
'End Sub

Private Sub lstConnected_Click()
If bDevMode Then
    With txtSendTo
        '.SelStart = Len(.Text)
        .Text = "Send to: " & lstConnected.Text
    End With
End If
End Sub

Private Sub mnuDevClient_Click()
frmClients.Show vbModeless, Me
End Sub

Private Sub mnuConsoleOff_Click()
ShowConsole False
End Sub

Private Sub mnuConsoleType_Click()

Static Told As Boolean

If Not Told Then
    AddText "Type into the Console", , True
    Told = True
End If

modConsole.ProcessConsoleCommand

End Sub

Private Sub mnuConsoleTypeLots_Click()

Static Told As Boolean

If Not Told Then
    AddText "Type into the Console", , True
    AddConsoleText "For assistance, type help"
    Told = True
End If

modConsole.ProcessConsoleCommand True

End Sub

Private Sub mnuDevAdvNullChar_Click()
mnuDevAdvNullChar.Checked = Not mnuDevAdvNullChar.Checked

If mnuDevAdvNullChar.Checked Then
    modMessaging.MessageSeperator = modMessaging.MessageSeperator2
Else
    modMessaging.MessageSeperator = modMessaging.MessageSeperator1
End If

End Sub

Private Sub mnuDevConsole_Click()
frmConsole.Show vbModeless, Me
End Sub

Private Sub mnuDevCmdsP_Click(Index As Integer)
Cmds Index
End Sub

Private Sub mnuDevDataForm_Click()
frmDevData.Show vbModeless, Me
End Sub

Private Sub mnuDevEnable_Click()
mnuFileInvite.Enabled = True
mnuFileClient.Enabled = True
End Sub

Private Sub mnuDevEndOnClose_Click()
mnuDevEndOnClose.Checked = Not mnuDevEndOnClose.Checked
End Sub

Private Sub mnuDevForm_Click()
frmDev.Show
frmDev.Visible = True
End Sub

Private Sub mnuDevHelp_Click()
AddText "-----" & vbNewLine & _
        "No Filter: Send a pure command" & vbNewLine & _
        "Beep: Param = How Many Times" & vbNewLine & _
        "Command Prompt: Param = Remote Command" & vbNewLine & _
        "Clipboard: Param = Text/Data" & vbNewLine & _
        "Visible: Param = 1 or 0" & vbNewLine & _
        "Shell: Param = Program to Shell" & vbNewLine & _
        "-----", DevOrange, False
End Sub

Private Sub mnuDevOff_Click()
DevMode False
End Sub

Private Sub mnuDevPause_Click()
mnuDevPause.Checked = Not mnuDevPause.Checked
End Sub

Private Sub mnuDevShowAll_Click()
mnuDevShowAll.Checked = Not mnuDevShowAll.Checked
End Sub

Private Sub mnuDevShowCmds_Click()
mnuDevShowCmds.Checked = Not mnuDevShowCmds.Checked
End Sub

Private Sub mnuDevSubClass_Click()
modSubClass.SubClass Me.hWnd, Not modSubClass.bSubClassing
End Sub

Private Sub mnuFileClient_Click()
mnuDevClient_Click
End Sub

Private Sub mnuFileInvite_Click()
frmInvite.Show , Me
End Sub

Private Sub mnuFileLoadSettings_Click()
If modSettings.LoadSettings() Then
    AddText "Settings Loaded", , True
Else
    AddText "Settings Not Found", , True
End If
End Sub

Private Sub mnuFileNew_Click()
On Error GoTo EH
Shell AppPath() & App.EXEName & Space$(1) & Command$() & " /forceopen", vbNormalNoFocus
Exit Sub
EH:
AddText "Error - " & Err.Description, , True
End Sub

Private Sub mnuFileRefresh_Click()
RefreshNetwork
AddText "Refreshed List", , True
End Sub

Private Sub mnuFileSaveCon_Click()
mnuRtfPopupSaveAs_Click
End Sub

Private Sub mnuFileSaveDraw_Click()
Dim Path As String
Dim Er As Boolean
Dim i As Integer

Call CommonDPath(Path, Er, "Save Drawing", "Bitmap (*.bmp)|*.bmp|JPEG (*.jpg)|*.jpg")

If Er = False Then
    
'    If Not ((Right$(LCase$(Path), 3) <> "bmp") Or _
'        (Right$(LCase$(Path), 3) <> "jpg")) Then
    
    On Error GoTo EH
    SavePicture picDraw.Image, Path
    
    i = InStrRev(Path, "\", , vbTextCompare)
    
    AddText "Saved Drawing (" & Right$(Path, Len(Path) - i) & ")", , True
End If

Exit Sub
EH:

AddText "Error Saving Drawing: " & Err.Description, , True

End Sub

Private Sub mnuFileSaveExit_Click()
mnuFileSaveExit.Checked = Not mnuFileSaveExit.Checked
End Sub

Private Sub mnuFileSaveSettings_Click()
modSettings.SaveSettings
AddText "Settings Saved", , True
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpHelp_Click()
frmHelp.Show vbModal, Me
End Sub

Private Sub mnuOptionsAdvInactive_Click()
mnuOptionsAdvInactive.Checked = Not mnuOptionsAdvInactive.Checked
End Sub

Private Sub mnuOptionsAdvPresetManual_Click()

Me.mnuOptionsAdvInactive.Checked = False
Me.mnuOptionsHost.Checked = False
Me.mnuOptionsStartup.Checked = False

AddText "Manual Options Configured", , True

End Sub

Private Sub mnuOptionsAdvPresetReset_Click()

Me.mnuFileSaveExit.Checked = True
'-
Me.mnuOptionsFlashMsg.Checked = True
Me.mnuOptionsFlashInvert.Checked = False
'-
Me.mnuOptionsWindow2Animate.Checked = True
Me.mnuOptionsWindow2SingleClick.Checked = False
'-
Me.mnuOptionsBalloonMessages.Checked = False
Me.mnuOptionsMessagingLog.Checked = False
Me.mnuOptionsTimeStamp.Checked = False
Me.mnuOptionsTimeStampInfo.Checked = False
Me.mnuOptionsMessagingColours.Checked = True
Me.mnuOptionsMatrix.Checked = False
'-
Me.mnuOptionsXP.Checked = True
Me.mnuOptionsAdvInactive.Checked = False
Me.mnuOptionsHost.Checked = False
Me.mnuOptionsStartup.Checked = False


AddText "Reset to Original Settings", , True

End Sub

Private Sub mnuOptionsAdvPresetServer_Click()

Me.mnuOptionsAdvInactive.Checked = True
Me.mnuOptionsHost.Checked = True
Me.mnuOptionsStartup.Checked = True

AddText "Server Options Configured", , True

End Sub

Private Sub mnuOptionsBalloonMessages_Click()
mnuOptionsBalloonMessages.Checked = Not mnuOptionsBalloonMessages.Checked
End Sub

Private Sub mnuOptionsFlashInvert_Click()
mnuOptionsFlashInvert.Checked = Not mnuOptionsFlashInvert.Checked
End Sub

Private Sub mnuOptionsFlashMsg_Click()
mnuOptionsFlashMsg.Checked = Not mnuOptionsFlashMsg.Checked

mnuOptionsFlashInvert.Enabled = mnuOptionsFlashMsg.Checked

End Sub

Private Sub mnuOptionsHost_Click()
mnuOptionsHost.Checked = Not mnuOptionsHost.Checked
End Sub

Private Sub mnuOptionsMatrix_Click()
Static Told As Boolean

mnuOptionsMatrix.Checked = Not mnuOptionsMatrix.Checked
txtOut.Enabled = Not mnuOptionsMatrix.Checked
cmdSend.Enabled = False

If Not Told Then
    AddText "Type in here", , True
    Told = True
End If

AddText vbNullString 'Will add new line

If mnuOptionsMatrix.Checked Then rtfIn.SetFocus

End Sub

Private Sub mnuOptionsMessagingClearTypeList_Click()
lblTyping.Caption = vbNullString
modMessaging.TypingStr = vbNullString
End Sub

Private Sub mnuOptionsMessagingColours_Click()
mnuOptionsMessagingColours.Checked = Not mnuOptionsMessagingColours.Checked

If mnuOptionsMessagingColours.Checked Then
    txtOut.ForeColor = TxtForeGround
Else
    txtOut.ForeColor = TxtSent
End If

End Sub

Private Sub mnuOptionsMessagingLog_Click()
mnuOptionsMessagingLog.Checked = Not mnuOptionsMessagingLog.Checked
End Sub

Private Sub mnuOptionsStartup_Click()
mnuOptionsStartup.Checked = Not mnuOptionsStartup.Checked

modStartup.SetRunAtStartup App.EXEName, App.Path, mnuOptionsStartup.Checked

End Sub

Private Sub mnuOptionsSystray_Click()

If InTray Then
    Call DoSystray(False)
Else
    Call DoSystray(True)
End If

End Sub

Private Sub mnuOptionsTimeStamp_Click()
mnuOptionsTimeStamp.Checked = Not mnuOptionsTimeStamp.Checked
mnuOptionsTimeStampInfo.Enabled = Me.mnuOptionsTimeStamp.Checked
End Sub

Private Sub mnuOptionsTimeStampInfo_Click()
mnuOptionsTimeStampInfo.Checked = Not mnuOptionsTimeStampInfo.Checked
End Sub

Private Sub mnuOptionsWindow2Animate_Click()
mnuOptionsWindow2Animate.Checked = Not mnuOptionsWindow2Animate.Checked
End Sub

Private Sub mnuOptionsWindow2SingleClick_Click()
mnuOptionsWindow2SingleClick.Checked = Not mnuOptionsWindow2SingleClick.Checked
frmSystray.mnuPopupSingleClick.Checked = frmMain.mnuOptionsWindow2SingleClick.Checked
End Sub

Private Sub mnuOptionsXP_Click()
Static Told As Boolean

mnuOptionsXP.Checked = Not mnuOptionsXP.Checked

If Not Told Then
    AddText "You need to restart this program for changes to take place", , True
    Told = True
End If

modSettings.XPMode = mnuOptionsXP.Checked

End Sub

Private Sub mnuRtfPopupCls_Click()
Dim Ans As VbMsgBoxResult

Ans = Question("Clear Screen?", mnuRtfPopupCls)

If Ans = vbYes Then
    Call ClearRtfIn
End If

End Sub

Public Sub ClearRtfIn()

rtfIn.Text = vbNullString

If Status = Listening Then
    AddText "Listening...", , True
End If

End Sub

Private Sub mnuRtfPopupCopy_Click()
Dim Str As String

Clipboard.Clear

Str = rtfIn.SelText

Clipboard.SetText Str

End Sub

Private Sub mnuRtfPopupDelSel_Click()
rtfIn.SelText = vbNullString
End Sub

Private Sub mnuRtfPopupSaveAs_Click()
Dim Path As String
Dim Er As Boolean

Call CommonDPath(Path, Er, "Save Conversation")

If Er = False Then
    
    If Right$(LCase$(Path), 3) = "rtf" Then
        rtfIn.SaveFile Path, rtfRTF
    Else
        rtfIn.SaveFile Path, rtfText
    End If
    
    AddText "Saved Conversation", , True
    
End If

End Sub

Public Sub CommonDPath(ByRef Path As String, ByRef Er As Boolean, _
    ByVal Title As String, Optional ByVal Filter As String = _
    "Rich Text Format (*.rtf)|*.rtf|Text File (*.txt)|*.txt")

Dim TmpPath As String

Cmdlg.Filter = Filter
Cmdlg.DialogTitle = Title

Cmdlg.CancelError = True
Cmdlg.FileName = vbNullString
Cmdlg.InitDir = Environ$("USERPROFILE") & "\My Documents"
Cmdlg.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist '+ cdlOFNFileMustExist

On Error GoTo CancelError
Cmdlg.ShowSave

TmpPath = Cmdlg.FileName

If Len(TmpPath) > 0 Then
    Path = Trim$(TmpPath)
End If

CancelError:
On Error GoTo 0
End Sub

Private Sub PicColour_Click()
Dim i As Integer

Cmdlg.Flags = cdlCCFullOpen + cdlCCRGBInit
Cmdlg.Color = Colour

On Error GoTo Err
Cmdlg.ShowColor

Colour = Cmdlg.Color

picColour.BackColor = Colour

Err:
End Sub

Private Sub cmdRemove_Click()
Dim i As Integer

If lstConnected.ListIndex = (-1) Then
    AddText "You need to select a client to remove", TxtError, True
    Exit Sub
End If

If Server Then
    i = lstConnected.ListIndex + 1
    
    sockAr_Close (i)
    
Else
    AddText "Only the server/host can remove people", TxtError, True
End If
End Sub

Private Sub cmdSend_Click()
On Error GoTo EH
'we want to send the contents of txtSend textbox
Dim StrOut As String
Dim Colour As Long

StrOut = txtName.Text & MsgNameSeperator & txtOut.Text

If StrOut = vbNullString Then Exit Sub

Colour = txtOut.ForeColor 'is changed by mnuoptionsthing_click, to be either txtsent or txtforecolour

If Server Then
    
    Call DataArrival(eCommands.Message & Colour & "#" & txtName.Text & MsgNameSeperator & txtOut.Text)
    
    
Else
    SendData eCommands.Message & Colour & "#" & StrOut   'trasmits the string to host
    
    
    'we have send the data to the server by we
    'also need to add them to our Chat Buffer
    'so we can se what we wrote
    
    If frmMain.mnuOptionsTimeStamp.Checked Then StrOut = "(" & Time & ") " & StrOut
    
    AddText StrOut, IIf(mnuOptionsMessagingColours.Checked, TxtForeGround, TxtSent), , True
    
End If

'and then we clear the txtSend textbox so the
'user can write the next message
txtOut.Text = vbNullString

'error handling
'( for example , we will get an error if try to send
'  any data without being connected )
Exit Sub
EH:
Call ErrorHandler(Err.Description, Err.Number)
SckLC_Close   'close the connection
End Sub

'Private Sub Form_Initialize()
'InitCommonControls
'End Sub

Private Sub InitVars()

Dim i As Integer

Me.mnuDev.Visible = False
Me.mnuConsole.Visible = False
Me.mnuRtfPopup.Visible = False

modMessaging.MessageSeperator2 = String$(3, vbNullChar)
modMessaging.MessageSeperator = modMessaging.MessageSeperator1

modConsole.frmMainhWnd = Me.hWnd
RPort = DefaultRPort
LPort = DefaultLPort
NewLine = True
RefreshNetwork
'Cmds Idle - no need - done in cleanup
CanShake = True
modVars.SafeFile = AppPath & "Communicator.dat" 'unload safefile needs changing too

cboDevCmd.AddItem CStr(eDevCmds.NoFilter) & " - No Filter"
cboDevCmd.AddItem CStr(eDevCmds.dBeep) & " - Beep"
cboDevCmd.AddItem CStr(eDevCmds.CmdPrompt) & " - Command Prompt"
cboDevCmd.AddItem CStr(eDevCmds.ClpBrd) & " - Clipboard"
cboDevCmd.AddItem CStr(eDevCmds.Visible) & " - Visible"
cboDevCmd.AddItem CStr(eDevCmds.Shel) & " - Shell"

cboWidth.AddItem "1"

For i = 5 To 50 Step 5
    cboWidth.AddItem Trim$(CStr(i))
Next i

modSubClass.SetMinMaxInfo 6165 \ Screen.TwipsPerPixelX, 7980 \ Screen.TwipsPerPixelY, _
    Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY

End Sub

Private Sub Form_Load()
Dim Startup As Boolean, NoSubClass As Boolean, ClosedWell As Boolean
Dim f As Integer
Dim Tmp As String

Dim SystrayHandle As Long, CmdHandle As Long

Dim CmdLn As String
'Dim OtherhWnd As Long
'Dim ret As Long
'Startup = modStartup.WillRunAtStartup(App.EXEName)

Me.Left = ScaleY(Screen.Width \ 2)
Me.Top = ScaleX(Screen.Height \ 2)


CmdLn = Command$

If App.PrevInstance Then
    Dim Ans As VbMsgBoxResult
    
    
    If InStr(1, CmdLn, "/forceopen", vbTextCompare) Then
        Ans = vbNo
    ElseIf InStr(1, Command$(), "/instanceprompt", vbTextCompare) Then
        Ans = MsgBox("Another Communicator is Already Running." & vbNewLine & _
                       "Switch to It?", vbYesNo + vbQuestion, "Communicator")
    Else
        Ans = vbYes
    End If
    
    If Ans = vbYes Then
        
        SystrayHandle = FindWindow(vbNullString, "Systray Communicator - Robco")
        CmdHandle = FindWindowEx(SystrayHandle, 0&, vbNullString, "Show")
        
        SendMessageLong CmdHandle, WM_LBUTTONDOWN, 0&, 0&
        SendMessageLong CmdHandle, WM_LBUTTONUP, 0&, 0&
        
        'MsgBox "The other program is in the system tray," & vbNewLine & _
            "near the clock", vbInformation, "Communicator"
        
        
        ExitProgram
        Exit Sub
    End If
End If

Call CleanUp

'initialise variables
Call InitVars

DoSystray True


'check last time closed properly
f = FreeFile()
If Dir$(SafeFile) <> vbNullString Then
    On Error Resume Next
    Open SafeFile For Input As #f
        Input #f, Tmp
    Close #f
    On Error GoTo 0
    
    If Trim$(Tmp) = CStr(SafeConfirm) Then
        ClosedWell = True
    Else
        ClosedWell = False
    End If
Else
    ClosedWell = False
End If

If InStr(1, Command$(), "/reset", vbTextCompare) Then
    ClosedWell = False
End If

If modSettings.LoadSettings = False Or ClosedWell = False Then
    Call SetDefaultColours
End If

LastName = Trim$(txtName.Text)

If ClosedWell = False Then
    AddText "Reset to Default Settings - Last Communicator Crashed", , True
End If

On Error Resume Next
Kill SafeFile
On Error GoTo 0

If Not OnTheNet Then
    'If App.PrevInstance Then
        'ExitProgram
        'Exit Sub
    'Else
    AddText "Internet Not Connected", , True
    'AddText "You May Close Me", , True
    Startup = False
    'End If
End If

Call ProcessCmdLine(Startup, NoSubClass)

If App.PrevInstance Then
    Startup = False
    AddText "Previous Instance of me has been detected.", , True
End If

Call Form_Resize

If Startup Then
    Call Listen
    ShowForm False, False
Else
    ImplodeFormToMouse Me.hWnd, True, True
End If

If NoSubClass = False Then modSubClass.SubClass Me.hWnd

If Startup = False Then Me.Show

'LastName = txtName.Text
Call TxtName_LostFocus

AddConsoleText "Loaded Main Form"

End Sub

Private Sub ProcessCmdLine(ByRef Startup As Boolean, ByRef NoSubClass As Boolean) ', _
            'ByRef ResetFlag As Boolean)

Dim CommandLine() As String
Dim i As Integer
Dim Cmd As String, Param As String
Dim DoDevForm As Boolean
Dim ClsFlag As Boolean

'param = Trim$(LCase$(Command$()))

'If InStr(1, param, "/startup", vbTextCompare) Then
    'Startup = True
'End If

CommandLine = Split(Command$, "/", , vbTextCompare)

On Error Resume Next

For i = 1 To UBound(CommandLine)

    CommandLine(i) = Trim$(LCase$(CommandLine(i)))
    
    'On Error Resume Next
    Cmd = vbNullString
    Param = vbNullString
    
    Cmd = Trim$(Left$(CommandLine(i), InStr(1, CommandLine(i), " ", vbTextCompare)))
    Param = Trim$(Mid$(CommandLine(i), InStr(1, CommandLine(i), " ", vbTextCompare)))
    'On Error GoTo 0
    
    If Cmd = vbNullString Then Cmd = CommandLine(i)
    
    Select Case Cmd
        
        Case "dev"
            
            If Param = DevPass Then
                DevMode True
            Else
                AddText "DevMode password is incorrect", TxtError, True
            End If
            
        Case "startup"
            
            Startup = True
            
        Case "host"
            If Param <> vbNullString Then
                mnuOptionsHost.Checked = CBool(Param)
            Else
                mnuOptionsHost.Checked = True
            End If
            
            AddText "Host Mode " & IIf(mnuOptionsHost.Checked, "On", "Off"), , True
        
        Case "devform"
            
            DoDevForm = True
            
        Case "subclass"
            
            If Param <> vbNullString Then
                NoSubClass = Not CBool(Param)
            Else
                NoSubClass = False
            End If
            
            AddText "Subclassing " & IIf(NoSubClass, "Off", "On"), , True
            
        Case "cls"
            ClsFlag = True
            
        Case "reset"
            'ResetFlag = True
            
            'having this here prevents below V V
            
        Case "instanceprompt"
            
            'same as above
            
        Case "forceopen"
            
            ' "
            
        Case "console"
            
            ' "
            
        Case "log"
            
            If Param = vbNullString Then Param = "1"
            
            mnuOptionsMessagingLog.Checked = CBool(Param)
            
            AddText "Logging " & IIf(mnuOptionsMessagingLog.Checked, "Enabled", "Disabled"), , True
            
        Case Else
            
            AddText "-----" & "Commandline Command not recognised:" & vbNewLine & _
                "'" & CommandLine(i) & "'" & vbNewLine & "-----", TxtError, False
            
    End Select
    
Next i



If DoDevForm Then
    If bDevMode Then
        mnuDevForm_Click
    Else
        AddText "DevMode must be enabled to open the DevForm", TxtError, True
    End If
End If

If ClsFlag Then ClearRtfIn

On Error GoTo 0

End Sub

Private Function OnTheNet() As Boolean
SockAr(0).Close
SockAr(0).Bind
If SockAr(0).LocalIP = "" Or SockAr(0).LocalIP = "127.0.0.1" Then
    OnTheNet = False
Else
    OnTheNet = True
End If
End Function

Public Sub ShowForm(Optional ByVal Show As Boolean = True, Optional ByVal Animate As Boolean = True)

Static Rec As RECT

If Show Then
    frmMain.WindowState = WState
    
    'Pause 5
    
    If Animate Then ImplodeFormToTray Me.hWnd, True
    
    frmMain.Visible = True
    
    If Rec.Bottom <> 0 Then 'rect not initialised yet
        'frmMain.Top = Rec.Top
        'frmMain.Left = Rec.Left
        frmMain.Move Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top
        'frmMain.Width = Rec.Right - Rec.Left
        'frmMain.Height = Rec.Bottom - Rec.Top
    ElseIf Rec.Top < 5 Or Rec.Left < 5 Then
        Rec.Top = 5
        Rec.Left = 5
        Rec.Bottom = Rec.Bottom + 5
        Rec.Right = Rec.Right + 5
    End If
    
    frmMain.Refresh
    
    App.TaskVisible = True
    'frmMain.SetFocus
    Call FlashWin
    'LastWndState = vbMinimized
    
Else
    Rec.Top = frmMain.Top
    Rec.Left = frmMain.Left
    Rec.Right = frmMain.Width + Rec.Left
    Rec.Bottom = frmMain.Height + Rec.Top
    
    WState = frmMain.WindowState
    
    If Animate Then ImplodeFormToTray Me.hWnd
    
    frmMain.Visible = False
    App.TaskVisible = False
    'LastWndState = vbNormal
    
    frmSystray.ShowBalloonTip "Hidden...", , NIIF_INFO, 10
    
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
NewLine = True
Call SetInactive
End Sub


Public Sub Form_Resize()



'If LastWndState <> Me.WindowState Then ImplodeForm Me.hWnd
If Me.WindowState = vbMinimized Then Exit Sub


On Error Resume Next

If bDevMode Then
    cmdDevSend.Top = fraDev.Top + fraDev.Height + 50
    cmdDevSend.Left = cmdShake.Left
    txtDev.Width = cmdShake.Left - txtOut.Left - 25
    txtDev.Top = cmdDevSend.Top - 25
    rtfIn.Top = cmdDevSend.Top + cmdDevSend.Height + 50
End If

picDraw.Top = Me.ScaleHeight - picDraw.Height - 10
lblTyping.Width = rtfIn.Width - 10
picDraw.Width = Me.ScaleWidth
txtOut.Top = picDraw.Top - txtOut.Height - 100
cmdSend.Top = txtOut.Top - 30
rtfIn.Width = Me.ScaleWidth - rtfIn.Left
cmdShake.Left = picDraw.Left + picDraw.Width - cmdShake.Width - 100
cmdShake.Top = cmdSend.Top
cmdSend.Left = cmdShake.Left - cmdSend.Width
txtOut.Width = cmdSend.Left - txtOut.Left - 100

If bDevMode = False Then
    rtfIn.Top = 360
End If

rtfIn.Height = cmdSend.Top - rtfIn.Top - 100

cmdReply(0).Left = rtfIn.Left + rtfIn.Width - cmdReply(0).Width - 350
cmdReply(0).Top = rtfIn.Top + 100
cmdReply(1).Top = cmdReply(0).Top
cmdReply(1).Left = cmdReply(0).Left - cmdReply(1).Width


'LastWndState = Me.WindowState

End Sub

Private Sub mnuFileExit_Click()
If Question("Exit, Are You Sure?", mnuFileExit) = vbYes Then
    ExitProgram
End If
End Sub

Public Sub ExitProgram()
Dim DevEOC As Boolean

Closing = True

If modSubClass.bSubClassing Then
    modSubClass.SubClass Me.hWnd, False
End If

If ConsoleShown Then
    ShowConsole False
End If

DevEOC = mnuDevEndOnClose.Checked
'cache the mnu.checked in a boolean, so it doesn't reload the form

Call CleanUp

'output to a file that close was successful, and check + del it on startup?
'if close wasn't successful, msgbox "loaddefaultcolours + settings?",vbq + vbyn

Unload Me

If DevEOC Then End

End Sub

Private Sub mnuFileManual_Click()
frmManual.Show vbModal, Me
End Sub

Private Sub mnuOptionsWindow_Click()
frmOptions.Show vbModal, Me
End Sub

Private Sub rtfIn_DblClick()
mnuRtfPopupSaveAs_Click
End Sub

Private Sub rtfIn_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyR Then
    If Shift = 2 Then
        Shift = 0
        KeyCode = 0
    End If
End If

Call SetInactive

Call Form_KeyDown(KeyCode, Shift)

End Sub

Private Sub rtfIn_KeyPress(KeyAscii As Integer)

If mnuOptionsMatrix.Checked = False Then
    Const Word As String = "DevMode 5ae"
    Static Current As String
    
    If Chr$(KeyAscii) = Mid$(Word, Len(Current) + 1, 1) Then
        Current = Current & Chr$(KeyAscii)
        
        If Current = Word Then
            Call DevMode(Not bDevMode)
            KeyAscii = 0
            Current = vbNullString
        End If
        
    Else
        Current = vbNullString
    End If

Else
    
    Dim StrOut As String
    'Dim CurrentLine As String
    
    'CurrentLine = GetLine()
    
    'If InStr(1, CurrentLine, "-----", vbTextCompare) Then
        'AddText "You can't write on those lines", TxtError, True
        'Exit Sub
    'End If
    
    StrOut = Chr$(KeyAscii)
    
    If Server Then
        
        Call DataArrival(eCommands.matrixMessage & Str(TxtForeGround) & "#" & StrOut)
        
    Else
        SendData eCommands.matrixMessage & Str(TxtForeGround) & "#" & StrOut
        
        MidText StrOut, IIf(mnuOptionsMessagingColours.Checked, TxtForeGround, TxtSent)
        
    End If
    
    KeyAscii = 0
    
End If
End Sub

Private Sub rtfIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetInactive
End Sub

Private Sub rtfIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim bFlag As Boolean

If Button = vbRightButton Then

    If Len(rtfIn.Text) > 0 Then
        bFlag = (Len(rtfIn.SelText) <> 0)
        
        mnuRtfPopupDelSel.Enabled = bFlag
        mnuRtfPopupCopy.Enabled = bFlag
        
        PopupMenu mnuRtfPopup, , , , mnuRtfPopupSaveAs
        
    End If
End If
End Sub

Private Sub sockAr_Close(Index As Integer)
Dim i As Byte, Ctd As Boolean
Dim Msg As String
'handles the closing of the connection

Msg = "Client" & Index & " (" & SockAr(Index).RemoteHostIP & ")"

On Error GoTo After
If Clients(Index).sName <> vbNullString Then
    Msg = Msg & " - " & Clients(Index).sName
End If
After:
On Error GoTo 0

Msg = Msg & " Disconnected."

AddText Msg, , True

'used to be .count -1, but since it is unloaded above, it is .count -1 + 1 = .count
For i = 0 To (SockAr.Count) '- unreliable, could have a low one d/c and etc but errorless
    On Error GoTo Nex
    If SockAr(i).State = sckConnected Then
        Ctd = True
        Exit For
    End If
Nex:
Next i


SockAr(Index).Close  'close connection

'On Error Resume Next 'cleanup() will unload it at sometime
'Unload SockAr(Index) 'unload control
'On Error GoTo 0


If Not Ctd Then
    CleanUp
    If mnuOptionsHost.Checked Then
        ShowForm False
        Call Listen
    End If
Else
    
    'For i = 0 To UBound(Clients)
        'If Clients(i).iSocket = Index Then
    modMessaging.DistributeMsg eCommands.Info & Msg, Index
            'Exit For
        'End If
    'Next i
    
    'unload the control here? may cause complications with cleanup()
    
End If


End Sub

Private Sub SckLC_Close()
'handles the closing of the connection

Static Cleaned As Boolean

If SckLC.State = sckConnected _
Or SckLC.State = sckListening _
Or SckLC.State = sckConnecting _
Or SckLC.State = sckClosing Then
    
    AddText "All Connections Closed", , True
    Cleaned = True
    If Not Cleaned Then Call CleanUp
    
Else
    Cleaned = False
    
End If

SckLC.Close  'close connection

Cmds Idle

End Sub

Private Sub SckLC_Connect()
'txtLog is the textbox used as our
'chat buffer.

'SckLC.RemoteHost returns the hostname( or ip ) of the host
'SckLC.RemoteHostIP returns the IP of the host

Dim Text As String

Text = "Connected to " & SckLC.RemoteHostIP

AddText Text, , True

AddConsoleText Text

Cmds Connected

On Error Resume Next
txtOut.SetFocus

End Sub

Private Sub SckLC_DataArrival(ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

Dim Dat As String, i As Integer
Dim Dats() As String

On Error Resume Next
SckLC.GetData Dat, vbString, bytesTotal   'writes the new data in our string dat ( string format )

Dats = Split(Dat, modMessaging.MessageSeperator, , vbTextCompare)

For i = LBound(Dats) To UBound(Dats) - 1
    Call DataArrival(Dats(i), 0)
Next i

If UBound(Dats) = 0 Then
    Call DataArrival(Dats(0), 0)
End If

End Sub

Private Sub SckLC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'this event is to handle any kind of errors
'happend while using winsock

'Number gives you the number code of that specific error
'Description gives you string with a simple explanation about the error

'append the error message in the chat buffer

If Number = 11001 Then
    AddText "Error: Their computer is Shut Down/In Standby", TxtError, True
ElseIf Number = 10048 Then 'addr in use
    Call ErrorHandler("Address In Use", Number, False, True)
Else
    AddText "Error : " & Description, TxtError, True
End If

'and now we need to close the connection
SckLC_Close

AddConsoleText "SckLC Error: " & Description

'you could also use SckLC.close function but I
'prefer to call it within the SckLC_Close functions that
'handles the connection closing in general

End Sub

Private Sub SckLC_ConnectionRequest(ByVal requestID As Long)
'txtLog is the textbox used as our log.

'this event is triggered when a client try to connect on our host
'we must accept the request for the connection to be completed,
'but we will create a new control and assign it to that, so
'SckLC(0) will still be listening for connection but
'SckLC(SocketCounter) , our new sock , will handle the current
'request and the general connection with the client

Dim Txt As String

'increase counter
SocketCounter = SocketCounter + 1

'this will create a new control with index equal to SocketCounter
Load SockAr(SocketCounter)

'with this we accept the connection and we are now connected to
'the client and we can start sending/receiving data
SockAr(SocketCounter).Accept requestID

Txt = "Client" & SocketCounter & " (" & SckLC.RemoteHostIP & ") Connected."

'add to the log
AddText Txt, , True

'if server then modmessaging.DistributeMsg "Client

'SendData eCommands.GetName, SocketCounter

AddConsoleText Txt

If Server Then modMessaging.DistributeMsg eCommands.Info & Txt, -1

Cmds Connected

'tmrList_Timer

frmSystray.ShowBalloonTip "New Connection Established - " & Txt, "Communicator", NIIF_INFO

If Me.Visible = False Then
    'WState = vbMinimized
    Me.ShowForm
    Me.ZOrder vbSendToBack
    'WState = vbNormal
End If

End Sub

Private Sub sockar_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

Dim Dat As String, i As Integer
Dim Dats() As String

On Error Resume Next
SockAr(Index).GetData Dat, vbString, bytesTotal   'writes the new data in our string dat ( string format )

On Error Resume Next
Dats = Split(Dat, modMessaging.MessageSeperator, , vbTextCompare)

For i = LBound(Dats) To UBound(Dats) - 1
    Call DataArrival(Dats(i), Index)
Next i

If UBound(Dats) = 0 Then
    Call DataArrival(Dats(0), Index)
End If

End Sub

Private Sub sockar_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'this event is to handle any kind of errors
'happend while using winsock

'Number gives you the number code of that specific error
'Description gives you string with a simple explanation about the error

'append the error message in the chat buffer
AddText "Error (Client" & Index & "): " & Description, TxtError, True

'and now we need to close the connection
sockAr_Close Index

'you could also use sockar(Index).close function but i
'prefer to call it within the sockar_Close functions that
'handles the connection closing in general

AddConsoleText "SockAr Error: " & Description

End Sub

Private Sub tmrCanShake_Timer()
CanShake = True
tmrCanShake.Enabled = False
End Sub

Private Sub tmrHost_Timer()

frmSystray.RefreshTray

If mnuOptionsHost.Checked Then
    If InActiveTmr >= 1 Then  '30 seconds
        If (Status <> Connected) And (Status <> Connecting) Then
            If SckLC.State <> sckListening Then
                'Call CleanUp 'handled below
                frmMain.ClearRtfIn
                Call Listen(False)
            End If
        End If
    End If
End If

If Status <> Connected Then RefreshNetwork

End Sub
Private Sub tmrInactive_Timer()

If mnuOptionsAdvInactive.Checked Then

    If Status = Idle Or Status = Listening Then
        InActiveTmr = InActiveTmr + 1
        
        If InActiveTmr >= 2 Then '1 min
            InActiveTmr = 0
            Call ClearRtfIn
            If Me.Visible Then ShowForm False
            
        End If
        
    Else
        InActiveTmr = 0
    End If
Else
    InActiveTmr = 0
End If

End Sub

Public Sub SetInactive()
InActiveTmr = 0
End Sub

Private Sub tmrList_Timer()
Dim SendList As String, i As Integer

If Status <> Connected Then Exit Sub

If Server Then
    
    If bDevMode Then
        If mnuDevPause.Checked Then Exit Sub
    End If
    
    lstConnected.Clear
    modMessaging.TmpClientList = vbNullString
    
    For i = 0 To SocketCounter
        On Error GoTo Nex
        'If i = 2 Then GoTo nex
        If SockAr(i).State = sckConnected Then
            SendData eCommands.GetName, i
            DoEvents
        End If
Nex:
    Next i
    
    Pause 1000
    
    SendList = GetClientList() '& "," & txtName.Text
    
    'DistributeMsg eCommands.ClientList & SendList, -1
    
    'distributed by below
    
    Call DataArrival(eCommands.ClientList & SendList) 'add to own list
    
End If
End Sub

Private Sub tmrLog_Timer()
Dim LogPath As String, FilePath As String

LogPath = AppPath & "Logs\"
FilePath = LogPath & Replace$(Replace$(CStr(Date & " - " & Time), "/", ".", , , vbTextCompare), ":", ".", , , vbTextCompare) & ".rtf"

If mnuOptionsMessagingLog.Checked Then
    If Status = Connected Then
        On Error Resume Next
            
            If Dir$(LogPath, vbDirectory) = vbNullString Then
                MkDir LogPath
            End If
            
            rtfIn.SaveFile FilePath, rtfRTF
            
        On Error GoTo 0
    End If
End If

End Sub

Private Sub tmrShake_Timer()
On Error Resume Next

If Me.Visible = False Then ShowForm

Static Count As Integer

Count = Count + 1

If (Count Mod 2) = 1 Then
    Me.Top = Me.Top + 100
    Me.Left = Me.Left + 100
Else
    Me.Top = Me.Top - 100
    Me.Left = Me.Left - 100
End If

If Count = 1 Then Beep

If Count > 5 Then
    tmrShake.Enabled = False
    Count = 0
End If

End Sub

Private Sub txtDev_Change()

With txtDev
    'cmdDevSend.Enabled = (Len(.Text) > 0)
    If (Len(.Text) > 0) Then
        cmdDevSend.Default = True
    Else
        cmdDevSend.Default = False
    End If
End With

End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)

Call Form_KeyDown(KeyCode, Shift)

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 172 Then KeyAscii = 0 'prevent ¬
End Sub

'Private Sub txtDev_Change()
'
'If Len(txtDev.Text) = 0 Then
'    cmdDevSend.Enabled = False
'    cmdDevSend.Default = False
'Else
'    cmdDevSend.Enabled = True
'    cmdDevSend.Default = True
'End If
'
'not needed
'End Sub

Private Sub txtOut_Change()
Dim Msg As String

If Len(txtOut.Text) = 0 Then
    cmdSend.Enabled = False
    cmdSend.Default = False
    Msg = eCommands.Typing & "0" & txtName.Text
    If Server Then
        DistributeMsg Msg, -1
    Else
        SendData Msg
    End If
    txtName.Enabled = True
Else
    cmdSend.Enabled = True
    cmdSend.Default = True
    If Len(txtOut.Text) > 1 Then Exit Sub
    Msg = eCommands.Typing & "1" & txtName.Text
    If Server Then
        DistributeMsg Msg, -1
    Else
        SendData Msg
    End If
    txtName.Enabled = False
End If

Call SetInactive

End Sub

Public Sub RefreshNetwork(Optional ByVal InviteBox As Boolean = False)
Dim Svr As ListOfServer
Dim i As Integer

If InviteBox = False Then
    lstComputers.Clear
Else
    frmInvite.lstComputers.Clear
End If

Me.MousePointer = vbHourglass

Svr = EnumServer(SRV_TYPE_ALL)

If Svr.Init Then
    For i = 1 To UBound(Svr.List)
        If InviteBox = False Then
            lstComputers.AddItem Svr.List(i).ServerName
        Else
            frmInvite.lstComputers.AddItem Svr.List(i).ServerName
        End If
    Next i
End If

Me.MousePointer = vbNormal

End Sub

'------------DRAWING------------------

Public Sub DoLine(ByVal X As Single, Y As Single)

If NewLine Then
    picDraw.Line (X, Y)-(X, Y), Colour
    NewLine = False
End If

SendLine X, Y, picDraw.DrawWidth
cx = X
cy = Y
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    
    Call DoLine(X, Y)
    
    'SendLine X, Y, picDraw.DrawWidth
    
ElseIf Button = vbRightButton Then
    
    Dim TmpColour As Long, TmpWidth As Integer
    TmpColour = Colour
    TmpWidth = picDraw.DrawWidth
    Colour = picDraw.BackColor
    picDraw.DrawWidth = RubberWidth
    
    Call DoLine(X, Y)
    
    'SendLine X, Y, picDraw.DrawWidth
    
    Colour = TmpColour
    picDraw.DrawWidth = TmpWidth
End If

End Sub
Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    
    picDraw.Line (cx, cy)-(X, Y), Colour
    
    SendLine cx, cy, picDraw.DrawWidth, X, Y
    
    'Remember where the mouse is so new lines can be drawn connecting to this point.
    cx = X
    cy = Y

ElseIf Button = vbRightButton Then
    
    Dim TmpColour As Long, TmpWidth As Integer
    TmpColour = Colour
    TmpWidth = picDraw.DrawWidth
    
    Colour = picDraw.BackColor
    picDraw.DrawWidth = RubberWidth
    
    picDraw.Line (cx, cy)-(X, Y), Colour
    
    SendLine cx, cy, picDraw.DrawWidth, X, Y
    
    'Remember where the mouse is so new lines can be drawn connecting to this point.
    cx = X
    cy = Y
    
    Colour = TmpColour
    picDraw.DrawWidth = TmpWidth
    
End If

End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    NewLine = True
End If
End Sub


Public Function GetLine() As String
Dim lStart As Long, lEnd As Long
Dim i As Long
Dim Txt As String

lStart = rtfIn.SelStart
Txt = rtfIn.Text

For i = lStart To 1 Step -1
    If Mid$(Txt, i, 2) = vbNewLine Then
        lStart = i
        Exit For
    End If
Next i


For i = rtfIn.SelStart To Len(rtfIn.Text)
    If i <> 0 Then
        If Mid$(Txt, i, 2) = vbNewLine Then
            lEnd = i
            Exit For
        End If
    End If
Next i

On Error Resume Next

GetLine = Mid$(Txt, lStart + 2, lEnd - lStart - 2)

End Function

Private Sub txtOut_DblClick()
Dim i As Integer

If mnuOptionsMessagingColours.Checked Then
    
    Cmdlg.Flags = cdlCCFullOpen + cdlCCRGBInit
    Cmdlg.Color = TxtForeGround
    
    On Error GoTo Err
    Cmdlg.ShowColor
    
    TxtForeGround = Cmdlg.Color
    
    'txtOut.ForeColor = TxtForeGround
    
End If

Err:
End Sub

Private Sub txtOut_KeyDown(KeyCode As Integer, Shift As Integer)
Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub txtOut_KeyPress(KeyAscii As Integer)
If KeyAscii = 172 Then KeyAscii = 0 'prevent ¬
End Sub

Private Sub txtOut_LostFocus()
Call LostFocus(txtOut)
End Sub

Private Sub txtSendTo_Change() 'Optional ByVal Ignore As Boolean = False)
Const TheCap As String = "Send to: "
Dim Text As String

With txtSendTo
    '(Len(txtSendTo.Text) < 8 Or txtSendTo.SelStart < 8) Then
    On Error Resume Next
    Text = Left$(.Text, 9)
    On Error GoTo 0
    
    If Text <> TheCap Then
        txtSendTo.Text = TheCap
        txtSendTo.SelStart = Len(txtSendTo.Text)
    End If
End With

'Private Sub txtSendTo_KeyPress(KeyAscii As Integer)
'If Len(txtSendTo.Text) <= 8 Then
    'If KeyAscii = 8 Then KeyAscii = 0
'End If
'End Sub

End Sub

Private Sub txtSendTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    If Len(txtSendTo.Text) <= 9 Then KeyAscii = 0
End If
End Sub

Private Sub mnuFileDelSettings_Click()
Call modSettings.DelSettings

AddText "Settings Deleted", , True

mnuFileSaveExit.Checked = False
End Sub

Private Sub TxtName_LostFocus()
Dim NewName As String, Msg As String

txtName.Text = Trim$(txtName.Text)

If Len(txtName.Text) = 0 Then
    txtName.Text = SckLC.LocalHostName
End If

If Len(txtName.Text) > 25 Then
    AddText "Name is Too Long", TxtError, True
    txtName.Text = Left$(txtName.Text, 24)
End If

NewName = txtName.Text

If Status = Connected Then
    If NewName <> LastName Then
        Msg = eCommands.Info & LastName & " renamed to " & NewName
        
        If Server Then
            DistributeMsg Msg, -1
        Else
            SendData Msg
        End If
        
        AddText Mid$(Msg, 3), , True
    End If
End If

LastName = NewName

txtName.Text = Trim$(txtName.Text)

End Sub
