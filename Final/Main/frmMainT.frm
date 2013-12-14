VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Communicator"
   ClientHeight    =   8985
   ClientLeft      =   75
   ClientTop       =   765
   ClientWidth     =   9435
   Icon            =   "frmMainT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   9435
   Begin VB.Timer tmrList 
      Interval        =   10000
      Left            =   3000
      Top             =   1680
   End
   Begin projMulti.ScrollListBox lstConnected 
      Height          =   975
      Left            =   1680
      TabIndex        =   4
      Top             =   1020
      Width           =   1455
      _extentx        =   2566
      _extenty        =   1720
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrivate 
      Caption         =   "Private Chat"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin MSScriptControlCtl.ScriptControl SC 
      Left            =   6000
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   0   'False
   End
   Begin VB.Frame fraDrawing 
      Caption         =   "Drawing"
      Height          =   1815
      Left            =   120
      TabIndex        =   39
      Top             =   3000
      Width           =   3195
      Begin VB.PictureBox picColours 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   1920
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   22
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00FF00FF&
         Height          =   255
         Index           =   6
         Left            =   1560
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   20
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Index           =   5
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   19
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   18
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H000080FF&
         Height          =   255
         Index           =   3
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   17
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   15
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   600
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   16
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   14
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkPickColour 
         Caption         =   "Pick Colour Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkStraightLine 
         Caption         =   "Straight Line Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1695
      End
      Begin VB.PictureBox picClearBoard 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   1455
         TabIndex        =   42
         Top             =   600
         Width           =   1455
         Begin VB.CommandButton cmdCls 
            Caption         =   "Clear Board"
            Height          =   375
            Left            =   360
            TabIndex        =   13
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox picColour 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   21
         Top             =   1080
         Width           =   255
      End
      Begin VB.ComboBox cboWidth 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cboRubber 
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Text            =   "1"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblLastColour 
         AutoSize        =   -1  'True
         Caption         =   "Last Colour"
         Height          =   195
         Left            =   2280
         TabIndex        =   44
         Top             =   1440
         Width           =   795
      End
      Begin VB.Label lblColour 
         Caption         =   "Colour"
         Height          =   165
         Left            =   2280
         TabIndex        =   43
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblDraw 
         Caption         =   "Draw:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblRubber 
         Caption         =   "Rubber:"
         Height          =   255
         Left            =   1560
         TabIndex        =   40
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSmile 
      Height          =   375
      Left            =   8400
      Picture         =   "frmMainT.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "Click here to see all emoticons"
      Top             =   3000
      Width           =   165
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   8730
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8043
            Text            =   "IP Panel"
            TextSave        =   "IP Panel"
            Object.ToolTipText     =   "IP Information"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Width           =   0
            Object.ToolTipText     =   "Ping Time (Milliseconds)"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8043
            Object.ToolTipText     =   "Menu Info"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      Height          =   3500
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   3435
      ScaleWidth      =   7155
      TabIndex        =   31
      Top             =   5280
      Width           =   7215
   End
   Begin MSWinsockLib.Winsock SckLC 
      Left            =   4680
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrLog 
      Interval        =   60000
      Left            =   7560
      Top             =   2400
   End
   Begin VB.Timer tmrCanShake 
      Interval        =   5000
      Left            =   8160
      Top             =   2880
   End
   Begin VB.Timer tmrShake 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8640
      Top             =   3240
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "&No"
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   35
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "&Yes"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   34
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrHost 
      Interval        =   30000
      Left            =   3960
      Top             =   4560
   End
   Begin MSComDlg.CommonDialog Cmdlg 
      Left            =   4080
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSWinsockLib.Winsock SockAr 
      Index           =   0
      Left            =   5280
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraDev 
      Height          =   735
      Left            =   3360
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtSendTo 
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Text            =   "Send to: "
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cboDevCmd 
         Height          =   315
         ItemData        =   "frmMainT.frx":203B
         Left            =   2160
         List            =   "frmMainT.frx":203D
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdShake 
      Caption         =   "Shake"
      Height          =   375
      Left            =   8640
      TabIndex        =   28
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox txtDev 
      Height          =   285
      Left            =   3360
      TabIndex        =   29
      Top             =   3480
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton cmdDevSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   8520
      TabIndex        =   30
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close Connection"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   600
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Host"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtOut 
      Height          =   285
      Left            =   3360
      TabIndex        =   26
      Top             =   3000
      Width           =   4215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   7680
      TabIndex        =   27
      Top             =   3000
      Width           =   615
   End
   Begin projMulti.smRtfFBox rtfIn 
      Height          =   2400
      Left            =   3360
      TabIndex        =   25
      Top             =   480
      Width           =   6000
      _extentx        =   10583
      _extenty        =   4233
      font            =   "frmMainT.frx":203F
      mouseicon       =   "frmMainT.frx":206B
      text            =   "0SMRTFBox"
      enabletextfilter=   -1
      selstart        =   1
      selrtf          =   $"frmMainT.frx":2089
   End
   Begin projMulti.ScrollListBox lstComputers 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   1455
      _extentx        =   2566
      _extenty        =   1720
   End
   Begin MSComctlLib.ImageList imglstIcons 
      Left            =   6720
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":20A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":4A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":73CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":9D5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":C6EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":F080
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":11A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":143A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":16D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":17A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":186EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainT.frx":193C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgStatus 
      Height          =   495
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblDevOverride 
      Caption         =   "Dev"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   48
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label lblDevOverride 
      Caption         =   "Dev"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   47
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblDevOverride 
      Caption         =   "Dev"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   46
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblDevOverride 
      Caption         =   "Dev"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   45
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblTyping 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   3360
      TabIndex        =   33
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   50
      TabIndex        =   32
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu mnuNote 
      Caption         =   "CtrlE + CtrlR + CtrlL can't be used"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Communicator"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenFolder 
         Caption         =   "Open Communicator Folder"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileRD 
         Caption         =   "Ranger Danger"
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
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSettings 
         Caption         =   "Settings"
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
         Begin VB.Menu mnuFileSettingsSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileSettingsExport 
            Caption         =   "Export Settings"
         End
         Begin VB.Menu mnuFileSettingsImport 
            Caption         =   "Import Settings"
         End
         Begin VB.Menu mnuFileSettingsSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileSettingsRMenu 
            Caption         =   "Communicator in Right Click Menu"
         End
      End
      Begin VB.Menu mnuFileConnection 
         Caption         =   "Connection/IPs"
         Begin VB.Menu mnuFileClient 
            Caption         =   "Client Window..."
            Shortcut        =   ^W
         End
         Begin VB.Menu mnuFileConnectionSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileRefresh 
            Caption         =   "Refresh Network List"
         End
         Begin VB.Menu mnuFileConnectionNetwork 
            Caption         =   "Detailed Network List"
            Shortcut        =   ^Y
         End
         Begin VB.Menu mnuFileManual 
            Caption         =   "Manual Connect..."
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuFileConnectionSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileIPs 
            Caption         =   "View IPs..."
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuFileNetIPs 
            Caption         =   "View Network IPs..."
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuFileInvite 
            Caption         =   "Invite..."
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileStealth 
         Caption         =   "Stealth Mode"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileGameMode 
         Caption         =   "Game Mode"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOptionsWindow 
         Caption         =   "Options Window..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOptionsVoice 
         Caption         =   "Voice Options"
      End
      Begin VB.Menu Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsWindow2 
         Caption         =   "Window"
         Begin VB.Menu mnuOptionsWindow2Animation 
            Caption         =   "Window Animation"
            Begin VB.Menu mnuOptionsWindow2Implode 
               Caption         =   "'Implode' Window"
            End
            Begin VB.Menu mnuOptionsWindow2Slide 
               Caption         =   "'Slide' Window"
            End
            Begin VB.Menu mnuOptionsWindow2Fade 
               Caption         =   "'Fade' Window"
            End
            Begin VB.Menu mnuOptionsWindow2All 
               Caption         =   "All Methods"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuOptionsWindow2AnimationSep 
               Caption         =   "-"
            End
            Begin VB.Menu mnuOptionsWindow2NoImplode 
               Caption         =   "Don't Animate"
            End
         End
         Begin VB.Menu mnuOptionsBalloonMessages 
            Caption         =   "Balloon Messages"
         End
         Begin VB.Menu mnuOptionsWindow2SingleClick 
            Caption         =   "Single Click Tray Icon"
         End
         Begin VB.Menu mnuOptionsWindow2BalloonInstance 
            Caption         =   "Show Balloon when second instance opens"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuOptionsMessaging 
         Caption         =   "Messaging"
         Begin VB.Menu mnuOptionsMessagingWindows 
            Caption         =   "Windows"
            Begin VB.Menu mnuOptionsMessagingPrivate 
               Caption         =   "Private Chat with..."
            End
            Begin VB.Menu mnuOptionsMessagingLobby 
               Caption         =   "Lobby Window"
               Shortcut        =   ^G
            End
         End
         Begin VB.Menu mnuOptionsMessagingDisplay 
            Caption         =   "Display"
            Begin VB.Menu mnuOptionsMessagingDisplaySysUserName 
               Caption         =   "Use System User Name as Default"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuOptionsMessagingDisplayNewLine 
               Caption         =   "Display Message on the line after the Name"
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
            Begin VB.Menu mnuOptionsFlashMsg 
               Caption         =   "Flash When Message Recieved"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuOptionsFlashInvert 
               Caption         =   "Flash Title Bar"
            End
            Begin VB.Menu mnuOptionsMessagingSep3 
               Caption         =   "-"
            End
            Begin VB.Menu mnuOptionsMessagingColours 
               Caption         =   "Allow Different Colours"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuOptionsMessagingSmilies 
               Caption         =   "Enable Smilies"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuOptionsMessagingShake 
               Caption         =   "Allow Shaking"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuOptionsMessagingDisplayShowBlocked 
               Caption         =   "Show if a blocked IP connects"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuOptionsMessagingDisplayIgnoreInvites 
               Caption         =   "Ignore All Invites (Auto-Reject)"
            End
         End
         Begin VB.Menu mnuOptionsMessagingSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsMatrix 
            Caption         =   "Matrix Chat Mode"
         End
         Begin VB.Menu mnuOptionsMessagingClearTypeList 
            Caption         =   "Clear Typing/Drawing List"
         End
         Begin VB.Menu mnuOptionsMessagingSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsMessagingIgnoreMatrix 
            Caption         =   "Ignore Matrix Messages"
         End
         Begin VB.Menu mnuOptionsMessagingDrawingOff 
            Caption         =   "Turn Off Drawing"
         End
         Begin VB.Menu mnuOptionsMessagingLog 
            Caption         =   "Log Conversations"
         End
         Begin VB.Menu mnuOptionsMessagingReplaceQ 
            Caption         =   "Replace / with ?"
         End
         Begin VB.Menu mnuOptionsMessagingEncrypt 
            Caption         =   "Encrypt Sent Messages"
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
         Begin VB.Menu mnuOptionsAdvSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsAdvInactive 
            Caption         =   "Inactivity Timer"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsAdvNoStandby 
            Caption         =   "Prevent Standby/Hibernation"
         End
         Begin VB.Menu mnuOptionsAdvSep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsHost 
            Caption         =   "Host Mode"
         End
         Begin VB.Menu mnuOptionsAdvHostMin 
            Caption         =   "Minimize to Tray When Hosting"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsAdvSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsStartup 
            Caption         =   "Startup"
         End
         Begin VB.Menu mnuOptionsAdvSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsAdvShowListen 
            Caption         =   "Show All Listen Errors"
         End
         Begin VB.Menu mnuOptionsAdvPing 
            Caption         =   "Ping"
         End
         Begin VB.Menu mnuDevConsole 
            Caption         =   "Console/CommandLine Commands"
         End
      End
      Begin VB.Menu mnuOptionsSystray 
         Caption         =   "Systray"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOptionsSocket 
         Caption         =   "Socket"
         Begin VB.Menu mnuOptionsSocketHost 
            Caption         =   "Host"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuOptionsSocketClose 
            Caption         =   "Close"
         End
      End
   End
   Begin VB.Menu mnuOnline 
      Caption         =   "Online"
      Begin VB.Menu mnuOnlineHTTP 
         Caption         =   "Use HTTP for FTP Download (Faster)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuonlinesep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnlineUpdates 
         Caption         =   "Check for Updates"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuOnlineSite 
         Caption         =   "Communicator Website"
      End
      Begin VB.Menu mnuOnlineLogin 
         Caption         =   "Login/Stats"
      End
      Begin VB.Menu mnuonlinesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnlineManual 
         Caption         =   "HTTP Download"
      End
      Begin VB.Menu mnuOnlineIPs 
         Caption         =   "View IPs"
      End
      Begin VB.Menu mnuOnlinePortForwarding 
         Caption         =   "Port Forwarding"
      End
      Begin VB.Menu mnuonlinesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnlineFileTransfer 
         Caption         =   "File Transfer"
      End
      Begin VB.Menu mnuOnlineMessages 
         Caption         =   "Messages"
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
         Begin VB.Menu mnuDevFormsClients 
            Caption         =   "Client List"
         End
      End
      Begin VB.Menu mnuDevDataCmds 
         Caption         =   "Data Commands"
         Begin VB.Menu mnuDevDataCmdsRecSent 
            Caption         =   "Show Received/Sent Data"
         End
         Begin VB.Menu mnuDevShowCmds 
            Caption         =   "Show Recieved Dev Commands"
         End
         Begin VB.Menu mnuDevDataCmdsClear 
            Caption         =   "Clear Text Box on Send"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuDevDataCmdsSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDevDataCmdsBlock 
            Caption         =   "Block Remote Commands"
         End
         Begin VB.Menu mnuDevDataCmdsSpecial 
            Caption         =   "Special"
            Begin VB.Menu mnuDevDataCmdsOverride 
               Caption         =   "Override Block"
            End
            Begin VB.Menu mnuDevDataCmdsNoReply 
               Caption         =   "Don't Reply to GetName"
            End
            Begin VB.Menu mnuDevDataCmdsSpecialOff 
               Caption         =   "Off"
            End
         End
         Begin VB.Menu mnuDevDataCmdsSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDevShowAll 
            Caption         =   "Show All Unknown Data"
         End
      End
      Begin VB.Menu mnuDevAdvCmds 
         Caption         =   "Advanced Commands"
         Begin VB.Menu mnuDevAdvCmdsDebug 
            Caption         =   "Debug Mode"
         End
         Begin VB.Menu mnuDevPause 
            Caption         =   "Pause Timers"
         End
         Begin VB.Menu mnuDevSubClass 
            Caption         =   "SubClass"
         End
         Begin VB.Menu mnuDevAdvCmdsSubrtfIn 
            Caption         =   "Unsubclass rtfIn"
         End
         Begin VB.Menu mnuDevEndOnClose 
            Caption         =   "End On Close"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuDevAdvNullChar 
            Caption         =   "Use NullChar Seperator"
         End
         Begin VB.Menu mnuDevAdvCmdsRemZLib 
            Caption         =   "Remove ZLib Dll"
         End
      End
      Begin VB.Menu mnuDevMaintenance 
         Caption         =   "Maintenance"
         Begin VB.Menu mnuDevMaintenanceTimers 
            Caption         =   "All Timers Off"
         End
      End
      Begin VB.Menu mnuDevPri 
         Caption         =   "Priority"
         Begin VB.Menu mnuDevPriRealtime 
            Caption         =   "Realtime"
         End
         Begin VB.Menu mnuDevPriHigh 
            Caption         =   "High"
         End
         Begin VB.Menu mnuDevPriAboveNormal 
            Caption         =   "Above Normal"
         End
         Begin VB.Menu mnuDevPriNormal 
            Caption         =   "Normal"
         End
         Begin VB.Menu mnuDevPriBelowNormal 
            Caption         =   "Below Normal"
         End
         Begin VB.Menu mnuDevPriLow 
            Caption         =   "Low"
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
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuDevCmdsP 
            Caption         =   "Connecting"
            Index           =   2
         End
         Begin VB.Menu mnuDevCmdsP 
            Caption         =   "Listening"
            Index           =   3
         End
         Begin VB.Menu mnuDevCmdsServer 
            Caption         =   "Server"
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
      Begin VB.Menu mnuRtfPopupCls 
         Caption         =   "Clear Screen"
      End
      Begin VB.Menu mnuRtfPopupSaveAs 
         Caption         =   "Save As..."
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
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuConsoleTypeLots 
         Caption         =   "Mutiple Commands"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuConsoleOff 
         Caption         =   "Turn Off"
      End
   End
   Begin VB.Menu mnuSB 
      Caption         =   "Statusbar"
      Begin VB.Menu mnuSBCopyrIP 
         Caption         =   "Copy External IP to Clipboard"
      End
      Begin VB.Menu mnuSBCopylIP 
         Caption         =   "Copy Internal IP to Clipboard"
      End
      Begin VB.Menu mnuSBObtain 
         Caption         =   "Obtain Remote IP"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'right click menu
Private Const RightClickExt = "*", RightClickMenuTitle = "Open Communicator"

Private Inviter As String 'who invited us? - add to convo on connect

'for the IP status bar
Private Const MinWidth As Integer = 10500

'make sure...
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Private Const QuestionTimeOut As Long = 60000

'for zipping
Private WithEvents ZipO As clsZipExtraction
Attribute ZipO.VB_VarHelpID = -1

Public pDrawDrawnOn As Boolean 'has picDraw been drawn on?

Private pHasFocus As Boolean

Private QuestionReply As Byte
Private CanShake As Boolean
'Private LastWndState As FormWindowStateConstants
Private Questioning As Boolean

Private InActiveTmr As Integer

Private Const MsMessageDelay As Long = 500 '1000
Private Const MsDevDelay As Long = 3000

'Drawing
Private PickingColour As Boolean
Private DrawingStraight As Boolean
Public Drawing As Boolean
Private SendTypeTrue As Boolean
Private SendTrueDraw As Boolean

'for "Connected in x seconds..."
Private ConnectStartTime As Long

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal Msg As Long, _
    wParam As Any, lParam As Any) As Long


Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Public LastName As String

Private Const LIPHeading As String = "Internal IP Address: "
Private Const RIPHeading As String = "External IP Address: "

Private Type PointInt
    X As Integer
    Y As Integer
End Type

Private StraightPoint1 As PointInt
Private StraightPoint2 As PointInt


Public Property Get HasFocus() As Boolean
HasFocus = pHasFocus
End Property

'--------------------------------------

Private Sub ShowSB_IP()

Dim brIP As Boolean

If LenB(lIP) = 0 Then lIP = frmMain.SckLC.LocalIP
If LenB(rIP) = 0 Then rIP = frmMain.GetIP()

brIP = Not CBool(InStr(1, rIP, "Error:", vbTextCompare))

If Len(rIP) > Len("xxx.xxx.xxx.xxx") Then
    brIP = False
ElseIf LenB(rIP) = 0 Then
    brIP = False
End If

If brIP = False Then
    rIP = vbNullString
    Me.mnuSBObtain.Visible = True 'allow them to re-obtain it
Else
    Me.mnuSBObtain.Visible = False
End If

SetPanelText LIPHeading & lIP & Space$(3) & _
    IIf(LenB(rIP), RIPHeading & rIP, _
    "Remote IP Error - Right Click Here"), 1
    '"Error fetching External IP, Please wait a minute and retry. (Right Click)"), 1

End Sub

Public Sub SetPanelText(Txt As String, ByVal PanelNo As Integer)

If Trim$(Me.sbMain.Panels(PanelNo).Text) <> Txt Then
    Me.sbMain.Panels(PanelNo).Text = Space$(12) & Txt & Space$(12)
End If

End Sub

Private Sub cboDevCmd_Change()
If Left$(cboDevCmd, 1) <> "0" Then
    txtSendTo.Enabled = True
Else
    txtSendTo.Text = vbNullString
    txtSendTo.Enabled = False
End If
End Sub

Private Sub cboDevCmd_Click()
cboDevCmd_Change
End Sub

Private Sub cboDevCmd_GotFocus()
cboDevCmd.Sellength = 0
End Sub

Private Sub cboRubber_Change()
Dim sTmp As String
Dim nTmp As Integer

sTmp = Trim$(cboRubber.Text)

If LenB(sTmp) Then
    
    On Error Resume Next
    nTmp = CInt(sTmp)
    On Error GoTo 0
    
    If (1 <= nTmp) And (nTmp <= 100) Then
        cboRubber.Text = nTmp
        modMessaging.RubberWidth = nTmp
    Else
        AddText "Please enter a Width between 1 and 50", TxtError, True
        cboRubber.Text = Trim$(Str(modMessaging.RubberWidth))
    End If
    
End If
End Sub

Private Sub cboRubber_Click()
cboRubber_Change
End Sub

Private Sub cboRubber_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If Not IsNumeric(Chr$(KeyAscii)) Then KeyAscii = 0
End Sub

Public Sub chkPickColour_Click()
PickingColour = (chkPickColour.Value = 1)
End Sub

Private Sub chkStraightLine_Click()
DrawingStraight = (chkStraightLine.Value = 1)
If DrawingStraight = False Then
    StraightPoint1.X = 0
    StraightPoint1.Y = 0
End If
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetPanelText "Connect to a computer", 3
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetPanelText "Close Current Connection", 3
End Sub

Private Sub cmdListen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetPanelText "Listen for a connection", 3
End Sub

Private Sub cmdPrivate_Click()
mnuOptionsMessagingPrivate_Click
End Sub

Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetPanelText "Kick someone from the server", 3
End Sub

Public Sub cmdScan_Click()
mnuFileNetIPs_Click
End Sub

Private Sub cmdScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetPanelText "Scan the network for other Communicators", 3
End Sub

Private Sub cmdSmile_Click()
rtfIn.CMDShowFaces_Click
End Sub

Private Sub Form_GotFocus()
pHasFocus = True
End Sub

Private Sub Form_Initialize()
modVars.SetProgress 25
End Sub

Private Sub Form_LostFocus()
pHasFocus = False
End Sub

Public Sub Form_Terminate()

'end api equivalent
If modLoadProgram.IsIDE() = False Then 'otherwise the IDE would close
    Call ExitProcess(0)
Else
    End
End If

End Sub

Private Sub fraDrawing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
frmMain.SetPanelText "Drawing Options", 3
End Sub

Private Sub lblDevOverride_Click(Index As Integer)

Static Pressed(0 To 3) As Boolean
Dim i As Integer
Dim ClearPressed As Boolean

If Index = 0 Then
    Pressed(0) = True
ElseIf Pressed(Index - 1) Then
    Pressed(Index) = True
    If Pressed(3) Then
        
        If LCase$(modVars.Password("Enter the password", Me)) = LCase$(UberDevPass) Then
            Me.mnuDevDataCmdsSpecial.Visible = True
            Beep
        End If
        
        ClearPressed = True
    End If
Else
    ClearPressed = True
End If

If ClearPressed Then
    For i = 0 To 3
        Pressed(i) = False
    Next i
End If

End Sub

Private Sub lstComputers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetPanelText "Computers on the Network", 3
End Sub

Private Sub lstConnected_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.SetPanelText "Connected Clients" & IIf(Not Server, " (and Server)", vbNullString), 3
End Sub

Private Sub mnuDevAdvCmdsDebug_Click()
'mnuDevAdvCmds.Checked = Not mnuDevAdvCmds.Checked
'menu is auto-set
modVars.bDebug = Not modVars.bDebug
End Sub

Private Sub mnuDevAdvCmdsRemZLib_Click()
AddText "ZLib Dll " & IIf(RemovezlibDll(), "removed", "couldn't be removed"), , True
End Sub

Private Sub mnuDevAdvCmdsSubrtfIn_Click()
rtfIn.DisableURLHook
End Sub

Private Sub mnuDevCmdsServer_Click()
mnuDevCmdsServer.Checked = Not mnuDevCmdsServer.Checked
Server = mnuDevCmdsServer.Checked
If Server Then
    mnuDevCmdsP_Click eStatus.Connected
End If
End Sub

Private Sub mnuDevDataCmdsBlock_Click()
Dim Pass As String

If mnuDevDataCmdsBlock.Checked Then
    mnuDevDataCmdsBlock.Checked = False
Else
    Pass = LCase$(modVars.Password("Enter a password to block developer commands...", Me))
    If Pass = LCase$(modVars.UberDevPass) Then
        mnuDevDataCmdsBlock.Checked = True
        mnuDevShowCmds.Checked = True
        AddText "Password Correct - Dev Commands Blocked", , True
    Else
        mnuDevDataCmdsBlock.Checked = False
        AddText "Password Incorrect", TxtError, True
    End If
End If

End Sub

Private Sub mnuDevDataCmdsClear_Click()
mnuDevDataCmdsClear.Checked = Not mnuDevDataCmdsClear.Checked
End Sub

Private Sub mnuDevDataCmdsNoReply_Click()
mnuDevDataCmdsNoReply.Checked = Not mnuDevDataCmdsNoReply.Checked
End Sub

Private Sub mnuDevDataCmdsOverride_Click()
mnuDevDataCmdsOverride.Checked = Not mnuDevDataCmdsOverride.Checked
End Sub

Private Sub mnuDevDataCmdsRecSent_Click()
mnuDevDataCmdsRecSent.Checked = Not mnuDevDataCmdsRecSent.Checked
End Sub

Private Sub mnuDevDataCmdsSpecialOff_Click()
Me.mnuDevDataCmdsSpecial.Visible = False
Beep
End Sub

Private Sub mnuDevFormsClients_Click()
Load frmDevClients
frmDevClients.Show vbModeless, Me
End Sub

Private Sub mnuDevPriAboveNormal_Click()
Call UncheckPriorities
modProcessPriority.ProcessPrioritySet , , ppAboveNormal
Call ShowPri
mnuDevPriAboveNormal.Checked = True
End Sub

Private Sub mnuDevPriBelowNormal_Click()
Call UncheckPriorities
modProcessPriority.ProcessPrioritySet , , ppbelownormal
Call ShowPri
mnuDevPriBelowNormal.Checked = True
End Sub

Private Sub mnuDevPriHigh_Click()
Call UncheckPriorities
modProcessPriority.ProcessPrioritySet , , ppHigh
Call ShowPri
mnuDevPriHigh.Checked = True
End Sub

Private Sub mnuDevPriLow_Click()
Call UncheckPriorities
modProcessPriority.ProcessPrioritySet , , ppidle
Call ShowPri
mnuDevPriLow.Checked = True
End Sub

Private Sub mnuDevPriNormal_Click()
Call UncheckPriorities
modProcessPriority.ProcessPrioritySet , , ppNormal
Call ShowPri
mnuDevPriNormal.Checked = True
End Sub

Private Sub mnuDevPriRealtime_Click()
Call UncheckPriorities
modProcessPriority.ProcessPrioritySet , , ppRealtime
Call ShowPri
mnuDevPriRealtime.Checked = True
End Sub

Private Sub ShowPri()

AddText "Process Priority set to " & _
    modProcessPriority.ProcessPriorityName( _
        modProcessPriority.ProcessPriorityGet()), , True

End Sub

Private Sub UncheckPriorities()
Me.mnuDevPriAboveNormal.Checked = False
Me.mnuDevPriBelowNormal.Checked = False
Me.mnuDevPriHigh.Checked = False
Me.mnuDevPriLow.Checked = False
Me.mnuDevPriNormal.Checked = False
Me.mnuDevPriRealtime.Checked = False
End Sub

Private Sub mnuFileConnectionNetwork_Click()
Load frmNetwork
frmNetwork.Show vbModeless, Me
End Sub

Public Sub mnuFileGameMode_Click()
mnuFileGameMode.Checked = Not mnuFileGameMode.Checked
frmSystray.mnuPopupGameMode.Checked = mnuFileGameMode.Checked

If modSpeech.sGameSpeak Then
    modSpeech.Say "Game Mode " & IIf(mnuFileGameMode.Checked, vbNullString, "De") & "activated.", , , True
End If

Call RefreshIcon

End Sub

Private Sub mnuFileIPs_Click()
Load frmIPs
frmIPs.Show vbModeless, Me
End Sub

Private Sub mnuFileNetIPs_Click()
frmUDP.ShowForm
End Sub

Public Sub mnuFileOpenFolder_Click()
Call OpenFolder(vbNormalFocus)
End Sub

'Private Sub mnuOptionsAdvInet_Click()
'Dim Ans As VbMsgBoxResult
'
'If modVars.CanUseInet Then
'    mnuOptionsAdvInet.Visible = False
'Else
'
'    AddText InfoStart & vbNewLine & _
'        "The Inet control is used for checking for updates," & vbNewLine & _
'        "and adding to the global IP list, and obtaining the External IP" & vbNewLine & _
'        InfoEnd, _
'        TxtInfo
'
'    Ans = Question("Warning - This could crash the program. Load?", mnuOptionsAdvInet)
'
'    If Ans = vbYes Then
'        modVars.LoadInet
'        AddText "Inet Control Loaded Successfully", , True
'        mnuOptionsAdvInet.Visible = False
'    End If
'End If
'
'End Sub

Public Sub mnuFileRD_Click()
frmRD.ShowForm
End Sub

Private Sub mnuFileSettingsExport_Click()
Dim Path As String, IDir As String
Dim Er As Boolean
Dim i As Integer

IDir = AppPath()

If FileExists(IDir, vbDirectory) = False Then
    IDir = vbNullString
End If

Call CommonDPath(Path, Er, "Export Settings", "Settings File (*." & FileExt & ")|*." & FileExt, IDir)

If Er = False Then
    
    If LenB(Path) Then
        
        modSettings.ExportSettings Path
        
        'text added by ^^
        
    End If
End If

End Sub

Private Sub mnuFileSettingsImport_Click()
Dim Path As String, IDir As String
Dim Er As Boolean
Dim i As Integer

IDir = AppPath()

If FileExists(IDir, vbDirectory) = False Then
    IDir = vbNullString
End If

Call CommonDPath(Path, Er, "Import Settings", "Settings File (*." & FileExt & ")|*." & FileExt, IDir, True)

If Er = False Then
    
    If LenB(Path) Then
        
        modSettings.ImportSettings Path
        
        'text added by ^^
        
    End If
End If
End Sub

Private Sub mnuFileSettingsRMenu_Click()

If mnuFileSettingsRMenu.Checked Then
    modVars.RemoveFromRightClick RightClickExt, RightClickMenuTitle
    
    AddText "Communicator removed from right click menu", , True
Else
    modVars.AddToRightClick RightClickExt, RightClickMenuTitle, AppPath() & App.EXEName
    
    AddText "Communicator added to right click menu", , True
End If

mnuFileSettingsRMenu.Checked = Not mnuFileSettingsRMenu.Checked

End Sub

Private Sub mnuFileStealth_Click()
StealthMode = True
End Sub

Private Sub mnuOnlineFileTransfer_Click()
Load frmFT
frmFT.Show vbModeless, Me
End Sub

Private Sub mnuOnlineHTTP_Click()
Static Told As Boolean

mnuOnlineHTTP.Checked = Not mnuOnlineHTTP.Checked

If Not Told And mnuOnlineHTTP.Checked Then
    Told = True
    AddText "FTP uploads will still use FTP Protocol", , True
End If

End Sub

Private Sub mnuOnlineIPs_Click()
mnuFileIPs_Click
End Sub

Private Sub mnuOnlineLogin_Click()
'Call CreateNewLoginThread
Load frmLogin
frmLogin.Show vbModeless, Me
End Sub

'Public Sub CreateNewLoginThread()
'
''Create a new instance of the MTDemo object on a new thread
'Dim Obj As clsThread 'the class obj
'Set Obj = CreateObject("projMulti.clsThread") 'set it
'
'AddConsoleText "Created new clsThread", , True, , True
'
'Call Obj.NewLoginFormThread   'ObjPtr(frmMain), frmMain.hWnd)
'
'AddConsoleText "New Thread Called, removing object..."
'
'Set Obj = Nothing
'
'AddConsoleText "Object Removed", , , True
'AddConsoleText vbNullString
'
'End Sub

Private Sub mnuOnlineManual_Click()
'ShellExecute 0&, vbNullString, modFTP.UpdateZip, vbNullString, vbNullString, vbNormalNoFocus
Dim Ret As Long
Dim Ans As VbMsgBoxResult
Dim Path As String

Ans = Question("Download via HTTP Protocol?", mnuOnlineManual)

If Ans = vbYes Then
    Path = AppPath() & "New Communicator.zip"
    
    AddText "Downloading...", , True
    Me.Refresh
    
    Ret = URLDownloadToFile(0, modFTP.UpdateZip, Path, 0, 0)
    
    If Ret = 0 And Dir$(Path) <> vbNullString Then
        AddText "Downloaded Successfully", , True
        
        
        Call ZipFileExtractQuestion(Path)
    '    Ans = Question("Open Folder?", mnuOnlineManual)
    '    If Ans = vbYes Then
    '        On Error Resume Next
    '        'Shell "explorer.exe " & Left$(Path, InStrRev(Path, "\", , vbTextCompare)), vbNormalFocus
    '        OpenFolder (vbNormalFocus)
    '        AddText "Folder Opened", , True
    '    End If
        AddText "Download Complete", , True
    Else
        AddText "Download Unsuccessful", , True
    End If
Else
    AddText "Download Canceled", , True
End If

End Sub

Private Sub mnuOnlineMessages_Click()
frmMessages.Show vbModeless, Me
End Sub

Private Sub mnuOnlinePortForwarding_Click()
Load frmPortForwarding
frmPortForwarding.Show vbModeless, Me
End Sub

Private Sub mnuOnlineSite_Click()
'ShellExecute 0&, vbNullString, modFTP.UpdateSite, vbNullString, vbNullString, vbNormalNoFocus
OpenURL modFTP.UpdateSite
End Sub

Private Sub mnuOnlineUpdates_Click()
Call CheckForUpdates
End Sub

Private Sub mnuOptionsAdvHostMin_Click()
mnuOptionsAdvHostMin.Checked = Not mnuOptionsAdvHostMin.Checked
End Sub

Private Sub mnuOptionsAdvNoStandby_Click()
mnuOptionsAdvNoStandby.Checked = Not mnuOptionsAdvNoStandby.Checked
End Sub

Private Sub mnuOptionsAdvPing_Click()
mnuOptionsAdvPing.Checked = Not mnuOptionsAdvPing.Checked
Me.sbMain.Panels(2).Visible = mnuOptionsAdvPing.Checked
End Sub

Private Sub mnuOptionsAdvShowListen_Click()
mnuOptionsAdvShowListen.Checked = Not mnuOptionsAdvShowListen.Checked
End Sub

Private Sub mnuOptionsMessagingDisplayIgnoreInvites_Click()

With mnuOptionsMessagingDisplayIgnoreInvites
    .Checked = Not .Checked
End With

End Sub

Private Sub mnuOptionsMessagingDisplayNewLine_Click()
mnuOptionsMessagingDisplayNewLine.Checked = Not mnuOptionsMessagingDisplayNewLine.Checked
End Sub

Private Sub mnuOptionsMessagingDisplayShowBlocked_Click()
mnuOptionsMessagingDisplayShowBlocked.Checked = Not mnuOptionsMessagingDisplayShowBlocked.Checked
End Sub

Private Sub mnuOptionsMessagingDisplaySysUserName_Click()
mnuOptionsMessagingDisplaySysUserName.Checked = Not mnuOptionsMessagingDisplaySysUserName.Checked
End Sub

Private Sub mnuOptionsMessagingDrawingOff_Click()

Dim Ans As VbMsgBoxResult

mnuOptionsMessagingDrawingOff.Checked = Not mnuOptionsMessagingDrawingOff.Checked

modMessaging.SendData eCommands.SetClientVar & eClientVarCmds.SetDrawing & _
                      CStr(IIf(mnuOptionsMessagingDrawingOff.Checked, 1, 0))


If mnuOptionsMessagingDrawingOff.Checked Then
    AddText "Drawing is Off - Data will not be sent from the host", , True
    picDraw.MousePointer = vbNormal
    'cmdCls_Click
Else
    AddText "Drawing is On", , True
    picDraw.MousePointer = vbCrosshair
End If

End Sub

Private Sub mnuOptionsMessagingEncrypt_Click()

mnuOptionsMessagingEncrypt.Checked = Not mnuOptionsMessagingEncrypt.Checked

AddText "Encryption " & IIf(mnuOptionsMessagingEncrypt.Checked, "En", "Dis") & "abled", , True

End Sub

Public Sub mnuOptionsMessagingLobby_Click()
If modSpaceGame.GameFormLoaded = False And modStickGame.StickFormLoaded = False Then
    Load frmLobby
    frmLobby.Show vbModeless, Me
Else
    AddText "Error - A Game Window is Open", TxtError, True
    
    If modSpaceGame.GameFormLoaded Then
        On Error Resume Next
        frmGame.SetFocus
    Else 'if modStickGame.StickFormLoaded then
        On Error Resume Next
        frmStickGame.SetFocus
    End If
    
End If
End Sub

Private Sub mnuOptionsMessagingIgnoreMatrix_Click()
mnuOptionsMessagingIgnoreMatrix.Checked = Not mnuOptionsMessagingIgnoreMatrix.Checked
End Sub

Private Sub mnuOptionsMessagingPrivate_Click()
Dim Tmp As String

If frmMain.mnuFileGameMode.Checked Then
    AddText "Game Mode is Active - Can't Use Private Chat", TxtError, True
Else
    On Error GoTo EH
    Tmp = Mid$(mnuOptionsMessagingPrivate.Caption, 19)
    
    If Tmp = LastName Then Exit Sub
    
    If LenB(Tmp) Then
        If Tmp <> ".." Then
            Dim Frm As Form
            
            For Each Frm In Forms
                If Mid$(Frm.Caption, Len(modVars.PvtCap) + 1) = Tmp Then
                    Exit For
                End If
            Next Frm
            
            If Frm Is Nothing Then
                Set Frm = New frmPrivate
                Load Frm
                Frm.Show vbModeless, Me
                Frm.SendTo = Tmp
            Else
                On Error Resume Next
                Frm.SetFocus
            End If
            
        Else
            AddText "Select someone to talk to", TxtError, True
        End If
    End If
End If

EH:
End Sub

Private Sub mnuOptionsMessagingReplaceQ_Click()
mnuOptionsMessagingReplaceQ.Checked = Not mnuOptionsMessagingReplaceQ.Checked
End Sub

Private Sub mnuOptionsMessagingShake_Click()
mnuOptionsMessagingShake.Checked = Not mnuOptionsMessagingShake.Checked
If Status = Connected Then
    cmdShake.Enabled = mnuOptionsMessagingShake.Checked
End If
End Sub

Private Sub mnuOptionsSocketClose_Click()
cmdClose_Click
End Sub

Private Sub mnuOptionsSocketHost_Click()
cmdListen_Click
End Sub

Private Sub mnuOptionsVoice_Click()
Load frmVoice
frmVoice.Show vbModeless, Me
End Sub

Private Sub mnuOptionsWindow2BalloonInstance_Click()
mnuOptionsWindow2BalloonInstance.Checked = Not mnuOptionsWindow2BalloonInstance.Checked
End Sub

Private Sub mnuSBCopylIP_Click()

On Error Resume Next
With Clipboard
    .Clear
    .SetText lIP
    Beep
End With

End Sub

Private Sub mnuSBCopyrIP_Click()

On Error Resume Next
With Clipboard
    .Clear
    .SetText rIP
    Beep
End With

End Sub

Private Sub mnuSBObtain_Click()
'sbMain.Panels(1).Text = "Obtaining External IP..."
SetPanelText "Obtaining External IP...", 1
Call ShowSB_IP
End Sub

Private Sub picColour_Change()
picColours(7).BackColor = picColour.BackColor
picColour.BackColor = Colour
End Sub

Private Sub picColours_Click(Index As Integer)
If Index <> 7 Then
    Colour = picColours(Index).BackColor
    picColour_Change
    picColour.BackColor = picColours(Index).BackColor
Else
    Colour = picColours(7).BackColor
    picColour_Change
End If
End Sub

Private Sub picColours_DblClick(Index As Integer)
If Index = 7 Then
    Call PicColour_Click
End If
End Sub

Private Sub sbMain_DblClick()
mnuSBCopyrIP_Click
End Sub

Private Sub sbMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    PopupMenu mnuSB, , , , mnuSBCopyrIP
End If
End Sub
'--------------------------------------

'Public Property Let DrawHeight(ByVal h As Integer)
''fraDraw.Height = h
''fraDraw.Top = 255 - fraDraw.Height
'
'picDraw.Height = h '- picDraw.Top
'End Property
'
'Public Property Get DrawHeight() As Integer
'DrawHeight = picDraw.Height
'End Property

Private Sub cboDevCmd_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboWidth_Change()
Dim sTmp As String
Dim nTmp As Integer

sTmp = Trim$(cboWidth.Text)

If LenB(sTmp) Then
    
    On Error Resume Next
    nTmp = CInt(sTmp)
    On Error GoTo 0
    
    If (1 <= nTmp) And (nTmp <= 50) Then
        picDraw.DrawWidth = nTmp
        cboWidth.Text = nTmp
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
    AddText "To save time, double click the listbox instead", TxtError, True
End If
End Sub

Public Function Connect(Optional ByVal Name As String = vbNullString) As Boolean
Dim sRemoteHost As String, Text As String

AddConsoleText "Beginning Connecting...", , True, , True

Connect = True

On Error GoTo EH

TxtName_LostFocus

Call CleanUp(False)

Cmds Connecting

sRemoteHost = Trim$(IIf(Name = vbNullString, lstComputers.List(lstComputers.ListIndex), Name))

SckLC.RemoteHost = sRemoteHost

'resolve host...

If SckLC.RemoteHost = vbNullString Then
    
    AddText "Please select a computer to connect to", TxtError, True
    Cmds Idle
    Connect = False
    AddConsoleText "No Computer Selected", , , True
Else
    
    LastIP = sRemoteHost
    SckLC.RemotePort = RPort
    SckLC.LocalPort = 0 'LPort
    
    Text = "Connecting to " & sRemoteHost & ":" & RPort & "..."
    AddText Text, , True
    AddConsoleText Text
    
    ConnectStartTime = GetTickCount()
    SckLC.Connect 'try to connect
    
    AddConsoleText "Began Connection Successfully", , , True
    
End If

Exit Function
EH:
AddConsoleText "Error Connecting - " & Err.Description, , , True
Call ErrorHandler(Err.Description, Err.Number) ', , True)
Connect = False
End Function

Public Sub CleanUp(ByVal SavePic As Boolean)
Dim N As Integer
Dim Frm As Form
Dim FilePath As String

AddConsoleText "Cleaning Up...", , True ', , True

If SckLC.State <> sckClosed Then
    SckLC_Close 'we close it in case it was trying to connect or whatever
Else
    'autosave the picture
    If Not Closing And SavePic And pDrawDrawnOn Then
        On Error Resume Next
        FilePath = AppPath() & "Last Pic.jpg"
        If FileExists(FilePath) Then
            Kill FilePath
        End If
        On Error Resume Next
        SavePicture picDraw.Image, FilePath
        AddText "Picture Saved to " & _
            Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "\", , vbTextCompare)) _
            , , True
    End If
    'end autosave
End If

Inviter = vbNullString
Server = False 'must be after scklc_close
SendTypeTrue = False 'for typingstr
SendTrueDraw = False 'for drawingstr

mnuOptionsMessagingPrivate.Caption = "Private Chat with..."

lstConnected.Clear
cmdRemove.Enabled = False

picDraw.Cls
pDrawDrawnOn = False

lblTyping.Caption = vbNullString
txtOut.Text = vbNullString
'txtOut_Change

ReDim Clients(0)
ReDim modMessaging.Typers(0)
ReDim modMessaging.Drawers(0)
ReDim modSpaceGame.CurrentGames(0)

modMessaging.TypingStr = vbNullString
modMessaging.DrawingStr = vbNullString

'modSpaceGame.ReceivedIPs = vbNullString

'If modSpaceGame.GameOptionFormLoaded Then
    'Unload frmGameOptions
'End If
If modSpaceGame.GameFormLoaded Then
    Unload frmGame
End If
If modStickGame.StickFormLoaded Then
    Unload frmStickGame
End If
'If modSpaceGame.StickFormLoaded Then
    'Unload frmStickGame
'End If
If mnuOptionsMatrix.Checked Then
    mnuOptionsMatrix_Click
End If


'mnuOptionsMessagingDrawingOff.Checked = False

'close and unload all previous sockets
For N = 1 To (SockAr.Count - 1)
    On Error Resume Next
    SockAr(N).Close
    Unload SockAr(N)
Next N

Cmds Idle

SocketCounter = 0

AddConsoleText "Cleaned Up", , , True

'frmSystray.ShowBalloonTip "All Connections Closed", "Communicator", NIIF_INFO

If modVars.nPrivateChats > 0 Then
    For Each Frm In Forms
        If Frm.Name = "frmPrivate" Then
            Unload Frm
        End If
    Next Frm
End If

End Sub

Public Sub cmdClose_Click()
Call CleanUp(True)
AddText "Connection Closed", , True
End Sub

Private Sub cmdCls_Click()
Dim Ans As VbMsgBoxResult
Dim Msg As String

Ans = Question("Clear Board, Are You Sure?", cmdCls)

If Ans = vbYes Then
    picDraw.Cls
    
    If pDrawDrawnOn Then pDrawDrawnOn = False
    
    Msg = LastName & " cleared the board"
    
    If Server Then
        DistributeMsg eCommands.Draw & "cls", -1
        DistributeMsg eCommands.Info & Msg & "0", -1
    Else
        SendData eCommands.Draw & "cls"
        SendData eCommands.Info & Msg & "0"
    End If
    
    AddText Msg, , True
    
    If Status = Connected Then
        cmdCls.Enabled = True
    Else
        cmdCls.Enabled = False
    End If
End If
End Sub

Private Sub cmdDevSend_Click()
Dim dMsg As String, SendTo As String, CmdNo As String
Static LastTick As Long

If (LastTick + MsDevDelay) < GetTickCount() Then
    On Error Resume Next
    SendTo = Trim$(Right$(txtSendTo.Text, Len(txtSendTo.Text) - 9))
    On Error GoTo 0
    
    If LenB(SendTo) = 0 And Left$(cboDevCmd.Text, 1) <> "0" Then
        AddText "Please Select a computer to send to", TxtError, True
        Exit Sub
    End If
    
    If Left$(cboDevCmd.Text, 1) = "-" Then
        CmdNo = Left$(cboDevCmd.Text, 2)
    Else
        CmdNo = Left$(cboDevCmd.Text, 1)
    End If
    
    Call SendDevCmd(CmdNo, SendTo, txtDev.Text, _
        (mnuDevDataCmdsOverride.Checked And mnuDevDataCmdsOverride.Visible))
    
    With txtDev
        If mnuDevDataCmdsClear.Checked Then
            .Text = vbNullString
            .SetFocus
        Else
            .Selstart = 0
            .Sellength = Len(.Text)
            .SetFocus
        End If
    End With
    
    LastTick = GetTickCount()
    
    'dmsg = eDevCmd & WhoTo & # & From & @ & Command
Else
    AddText "Don't spam Dev Commands, it's not nice", TxtError, True
End If

End Sub

Public Sub SendDevCmd(ByVal iCmd As Integer, ByVal SendTo As String, _
    ByVal Text As String, Optional ByVal Override As Boolean = False) ', Optional ByVal SocketSendTo As Integer = -1)

Dim dMsg As String

If iCmd Then
    dMsg = eCommands.DevSend & SendTo & "#" & Trim$(LastName) & "@" & iCmd & _
                Text & IIf(Override, modVars.DevOverride, vbNullString)
    
    AddDevText vbNewLine & "DevMode, Sent:" & vbNewLine & _
        "To: " & SendTo & vbNewLine & _
        "Command: " & CStr(iCmd) & vbNewLine & _
        "Parameter: " & Text & vbNewLine, True
Else
    dMsg = txtDev.Text
    
    AddDevText vbNewLine & "DevMode, Sent:" & vbNewLine & _
        "To: " & SendTo & vbNewLine & _
        "Message: " & Text & vbNewLine, True
End If

If Server Then
    DistributeMsg dMsg, -1
Else
    SendData dMsg
End If
End Sub

Public Sub cmdListen_Click()
Call Listen
End Sub

Public Function Listen(Optional ByVal ShowError As Boolean = True) As Boolean   'Optional ByVal DoEH As Boolean = True

modConsole.AddConsoleText "Begining Listening...", , True, , True

Call CleanUp(False)

On Error GoTo EH

Cmds Listening

SckLC.Close 'we close it in case it listening before


'txtPort is the textbox holding the Port number
SckLC.LocalPort = RPort  'set the port we want to listen to
                              '( the client will connect on this port too)
SckLC.RemotePort = 0 'LPort

On Error Resume Next
SckLC.Listen                'Start Listening

If SckLC.State <> sckListening Then GoTo EH

AddText "Listening...", , True

'frmSystray.ShowBalloonTip "Listening...", , NIIF_INFO, 1000

Server = True
Listen = True

AddConsoleText "Listening Successful", , , True


Exit Function
EH:
If ShowError Or mnuOptionsAdvShowListen.Checked Then
    Call ErrorHandler(Err.Description, Err.Number, True, True) ', DoEH)
Else
    CleanUp False
End If

AddConsoleText "Listening Failed", , , True

Listen = False

End Function

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
    DistributeMsg eCommands.Shake & LastName, -1
Else
    SendData eCommands.Shake & LastName
End If

AddText "Shake Sent by " & LastName, TxtSent, True

CanShake = False
tmrCanShake.Enabled = True

On Error Resume Next
txtOut.SetFocus

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) ' QueryUnloadConstants)
Dim Ans As VbMsgBoxResult
Dim Tmp As String

If UnloadMode = vbAppWindows Then
    modVars.Closing = True
ElseIf UnloadMode = vbAppTaskManager Then
    modVars.Closing = True
ElseIf UnloadMode = vbFormOwner Then
    modVars.Closing = True
'ElseIf UnloadMode = vbFormMDIForm Then
'    modVars.Closing = True
End If

'AddConsoleText "QUnload Being Called - Closing: " & CStr(Closing) & _
    " UnloadMode: " & CStr(UnloadMode), , True

If modFTP.OnlineStatusIs And Closing And _
        Not ((UnloadMode = vbAppTaskManager) Or (UnloadMode = vbAppWindows)) Then
    
    If frmMain.Visible = False Then frmMain.ShowForm
    
    Ans = MsgBoxEx("Online status was set when IPs were uploaded," & vbNewLine & _
            "Cancel so you can change it?", _
            vbYesNo + vbExclamation, "Online Status", , , frmMain.Icon)
    
    If Ans = vbYes Then
        AddText "Close Canceled", , True
        modVars.Closing = False
        Cancel = True
        Exit Sub 'don't hide the form
    End If
End If

'If modVars.APortForwarded And Closing Then
'    If frmMain.Visible = False Then frmMain.ShowForm
'
'    Ans = MsgBox("A port has been forwarded to the router," & vbNewLine & _
'            "Cancel so you can unforward it?", _
'            vbYesNo + vbExclamation, "Port Forwarding")
'
'    If Ans = vbYes Then
'        AddText "Close Canceled", , True
'        Closing = False
'        Cancel = True
'        Exit Sub 'don't hide the form
'    End If
'End If

'AddConsoleText "Closing: " & CStr(Closing)

If modVars.Closing Then
    
    AddText "Exiting...", , True, True
    rtfIn.Refresh
    
    Call CleanUp(True)
    
    If modSubClass.bSubClassing Then
        modSubClass.SubClass frmMain.hWnd, False
    End If
    
    If Not App.PrevInstance Then
        If modSpeech.sHiBye Then
            If modSpeech.sBye Then 'start saying here - terminated in _Unload
                
                If Me.mnuOptionsMessagingDisplaySysUserName.Checked Then
                    Tmp = modVars.GetUserName()
                Else
                    Tmp = Me.LastName
                End If
            
                modSpeech.Say "Goodbye " & Tmp
            End If
        End If
    End If
    
    If ConsoleShown Then
        ShowConsole False
    End If
    
    If modSpaceGame.GameFormLoaded Then
        frmGame.bRunning = False
    End If
    If modStickGame.StickFormLoaded Then
        frmStickGame.bRunning = False
    End If
    
    If mnuFileSaveExit.Checked Then modSettings.SaveSettings
    
    If InTray Then
        DoSystray False
    End If
    
    AddConsoleText "Goodbye!", , , True
    
    'If Me.Visible Then ImplodeFormToMouse Me.hwnd 'done after all others closed
Else
    Cancel = True
    
    ShowForm False
    'AddConsoleText "Canceled", , , True
    
End If

End Sub

Public Function Question(ByVal sMessage As String, Caller As Object) As VbMsgBoxResult

Dim TimeOutTime As Long, WarnTime As Long
Dim HadWarning As Boolean

If Questioning Then
    Question = vbRetry
    AddText "Please Answer the Previous Question first", TxtError, True
    cmdReply(0).Visible = True
    cmdReply(1).Visible = True
    
    cmdReply(1).Default = True
    On Error Resume Next
    cmdReply(1).SetFocus
    Exit Function
End If

Caller.Enabled = False
Questioning = True
HadWarning = False

AddText sMessage, TxtQuestion, True

If Me.Visible = False Then
    frmSystray.BalloonClickShow = True
    frmSystray.ShowBalloonTip sMessage, , NIIF_INFO
    'If Me.Visible = False Then Me.ShowForm
Else
    If modSpeech.sQuestions Then
        modSpeech.Say sMessage
    End If
End If

cmdReply(0).Visible = True
cmdReply(1).Visible = True

cmdReply(1).Default = True
On Error Resume Next
cmdReply(1).SetFocus

QuestionReply = 3

TimeOutTime = GetTickCount() + QuestionTimeOut
WarnTime = GetTickCount() + 2 * QuestionTimeOut / 3

Do
    Pause 100
    
    If QuestionReply <> 3 Then
        Select Case QuestionReply
            Case 1
                Question = vbYes
            Case 0
                Question = vbNo
        End Select
        
    ElseIf TimeOutTime < GetTickCount() Then
        QuestionReply = 0
        Question = vbIgnore 'vbRetry
        AddText "Question Timed Out", TxtError, True
        'for other procs
    ElseIf Not HadWarning Then
        If WarnTime < GetTickCount() Then
            AddText "Warning - Question will timeout soon", TxtError, True
            HadWarning = True
        End If
    End If
    
Loop While QuestionReply = 3 And Not Closing

If QuestionReply = 3 Then
    Question = vbRetry 'closing
End If

cmdReply(0).Visible = False
cmdReply(1).Visible = False

Caller.Enabled = True
Questioning = False

End Function

Private Sub Form_Unload(Cancel As Integer)
Dim Frm As Form

For Each Frm In Forms
    
    If Frm.Name = "frmDev" Then
        frmDev.CloseMe = True
    End If
    
    If Frm.Name <> "frmMain" Then Unload Frm
    
    'Set Frm = Nothing
    
    Me.Refresh
Next Frm

Me.Refresh

If Me.Visible Then modImplode.AnimateAWindow hWnd, aRandom, True 'ImplodeFormToMouse Me.hWnd

Me.Hide

Do While modSpeech.nSpeechStatus = sSpeaking
    Pause 10
Loop
Call SpeechTerminate 'give it a chance to save goodbye

If InTray Then 'for some reason, sometimes is...
    modVars.DoSystray False
End If

modWinsock.TermWinsock

'
'Dim f As Integer
'Dim SF As String
'
'f = FreeFile()
'
'SF = modVars.SafeFile
'
'On Error GoTo EH
'Open SF For Output As #f
'    Print #f, Str(SafeConfirm)
'EH:
'Close #f
'
'On Error Resume Next
'SetAttr SF, vbHidden

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
Dim Tmp As String

Tmp = Trim$(lstConnected.Text)

If LenB(Tmp) Then
    If Tmp <> LastName Then
        
        mnuOptionsMessagingPrivate.Caption = "Private Chat with " & Tmp
        
        If bDevMode Then
            If txtSendTo.Enabled Then
                With txtSendTo
                    '.SelStart = Len(.Text)
                    .Text = "Send to: " & lstConnected.Text
                End With
            End If
        End If
    Else
        mnuOptionsMessagingPrivate.Caption = "Private Chat with..."
    End If
Else
    mnuOptionsMessagingPrivate.Caption = "Private Chat with..."
End If

mnuOptionsMessagingPrivate.Enabled = LenB(Tmp)
cmdPrivate.Enabled = mnuOptionsMessagingPrivate.Enabled
cmdRemove.Enabled = cmdPrivate.Enabled And Server

End Sub

Private Sub mnuDevClient_Click()
frmClients.Show vbModeless, Me 'don't remove this, called by mnufileclient
End Sub

Private Sub mnuConsoleOff_Click()
ShowConsole False
End Sub

Private Sub mnuConsoleType_Click()

Static Told As Boolean

If Not Told Then
    AddText "Type into the Console", , True
    rtfIn.Refresh
    Me.Refresh
    Told = True
End If

modConsole.ProcessConsoleCommand

End Sub

Private Sub mnuConsoleTypeLots_Click()

Static Told As Boolean

If Not Told Then
    AddText "Type into the Console", , True
    AddConsoleText "For assistance, type help"
    rtfIn.Refresh
    Me.Refresh
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

Private Sub mnuDevEndOnClose_Click()
mnuDevEndOnClose.Checked = Not mnuDevEndOnClose.Checked
End Sub

Private Sub mnuDevForm_Click()
frmDev.Show
frmDev.Visible = True
End Sub

Private Sub mnuDevHelp_Click()
AddText "-----" & vbNewLine & _
        "Please use responsibly" & vbNewLine & vbNewLine & _
        "No Filter: Send a pure command" & vbNewLine & _
        "Beep: Parameter is how many beeps you want" & vbNewLine & _
        "Command Prompt: Parameter is the remote command" & vbNewLine & _
        "Clipboard: Parameter is text or data" & vbNewLine & _
        "Visible: Parameter is 1 or 0" & vbNewLine & _
        "Shell: Parameter is path to program to shell" & vbNewLine & _
        "Name: Parameter is name to set" & vbNewLine & _
        "Version: Get their version number" & vbNewLine & _
        "Disconnect: Disconnect them" & vbNewLine & _
        "Computer Name: Get their computer's name" & vbNewLine & _
        "Game Form: '0' to close, '<ip>' to connect, and blank to host" & vbNewLine & _
        "Caps Lock: Toggle or get the state of caps lock" & vbNewLine & _
        "VBScript: Parameter is the command" & vbNewLine & _
        "-----", DevOrange, False

End Sub

Private Sub mnuDevMaintenanceTimers_Click()
tmrList.Enabled = False
tmrHost.Enabled = False
tmrShake.Enabled = False
tmrCanShake.Enabled = False
'tmrInactive.Enabled = False
tmrLog.Enabled = False
AddText "You should restart the program to get it back to normal", , True
End Sub

Private Sub mnuDevOff_Click()
DevMode False
End Sub

Private Sub mnuDevPause_Click()
mnuDevPause.Checked = Not mnuDevPause.Checked
If mnuDevPause.Checked Then
    AddText "Timers Paused! Remember to unpause them", TxtError, True
End If
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

'Private Sub mnuFileInvite_Click()
'frmInvite.Show , Me
'End Sub

Private Sub mnuFileLoadSettings_Click()
If modSettings.LoadSettings() Then
    AddText "Settings Loaded", , True
Else
    AddText "Settings Not Found", , True
End If
End Sub

Public Sub mnuFileNew_Click()
Dim cmd As String

cmd = Command$()

If InStr(1, cmd, "/startup", vbTextCompare) Then
    cmd = Replace$(cmd, "/startup", vbNullString, , , vbTextCompare)
End If
If InStr(1, cmd, "/killold", vbTextCompare) Then
    cmd = Replace$(cmd, "/killold", vbNullString, , , vbTextCompare)
End If

On Error GoTo EH
Shell AppPath() & App.EXEName & Space$(1) & Trim$(cmd) & " /forceopen", vbNormalNoFocus

Exit Sub
EH:
AddText "Error - " & Err.Description, , True
End Sub

Private Sub mnuFileRefresh_Click()
RefreshNetwork
AddText "Refreshed List", , True
End Sub

Public Sub mnuFileSaveCon_Click()
mnuRtfPopupSaveAs_Click
End Sub

Private Sub mnuFileSaveDraw_Click()
Dim Path As String, IDir As String
Dim Er As Boolean
Dim i As Integer

IDir = AppPath() & "Logs"

If FileExists(IDir, vbDirectory) = False Then
    IDir = vbNullString
End If

Call CommonDPath(Path, Er, "Save Drawing", _
    "JPEG (*.jpg)|*.jpg|Bitmap (*.bmp)|*.bmp|PNG (*.png)|*.png", IDir)

If Er = False Then
    
'    If Not ((Right$(LCase$(Path), 3) <> "bmp") Or _
'        (Right$(LCase$(Path), 3) <> "jpg")) Then
    If LenB(Path) Then
        On Error GoTo EH
        SavePicture picDraw.Image, Path
        
        i = InStrRev(Path, "\", , vbTextCompare)
        
        AddText "Saved Drawing (" & Right$(Path, Len(Path) - i) & ")", , True
    End If
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

Public Sub CheckForUpdates()

Dim VerHTMLTxt As String
Dim nMaj As Integer, oMaj As Integer
Dim nMin As Integer, oMin As Integer
Dim nRev As Integer, oRev As Integer
Dim i As Integer, j As Integer
Dim HaveOld As Boolean
Dim Ans As VbMsgBoxResult
Const dSep As String = Dot
Dim Tmp As String
Dim Tmr As Long
Dim l As Double


'If modVars.CanUseInet Then

oMaj = App.Major
oMin = App.Minor
oRev = App.Revision

AddText "Checking for an update", , True
rtfIn.Refresh
VerHTMLTxt = modFTP.fGetVersion() 'GetHTML(modVars.UpdateURL & "/" & modVars.UpdateTxt)

'stored like: "1.13.5"

i = InStr(1, VerHTMLTxt, ".", vbTextCompare)
j = InStr(i + 1, VerHTMLTxt, ".", vbTextCompare)

On Error GoTo EH
nMaj = Left$(VerHTMLTxt, i - 1)
nMin = Mid$(VerHTMLTxt, i + 1, j - i - 1)
nRev = Mid$(VerHTMLTxt, j + 1)
On Error GoTo 0

HaveOld = False

If nMaj > oMaj Then
    HaveOld = True
ElseIf nMaj = oMaj Then
    If nMin > oMin Then
        HaveOld = True
    ElseIf nMin = oMin Then
        If nRev > oRev Then
            HaveOld = True
        End If
    End If
End If

AddText InfoStart & vbNewLine & "My Version: " & oMaj & dSep & oMin & dSep & oRev & vbNewLine & _
        "Latest Version: " & nMaj & dSep & nMin & dSep & nRev & vbNewLine & Trim$(InfoEnd), TxtInfo

If HaveOld Then
    
    Ans = Question("Newer Version Found, Update?", mnuOnlineUpdates)
    
    If Ans = vbYes Then
        'Open default INET browser
        'to site const
        'ShellExecute 0&, vbNullString, modFTP.UpdatePage, vbNullString, vbNullString, vbNormalNoFocus
        AddText "Downloading, Please Wait...", , True
        rtfIn.Refresh
        
        If modFTP.DownloadLatest() Then
            AddText "Download Successful", , True
            rtfIn.Refresh
            On Error Resume Next
            'On Error GoTo 0
            FileCopy (modFTP.RootDrive & "\" & modFTP.Communicator_File), _
                    AppPath() & modFTP.Communicator_File
            
            Kill modFTP.RootDrive & "\" & modFTP.Communicator_File
            On Error GoTo 0
            
            ZipFileExtractQuestion AppPath() & Left$(modFTP.Communicator_File, _
                InStr(1, modFTP.Communicator_File, ".", vbTextCompare) - 1) & ".zip"
            
        Else
            AddText "Error in download", , True
        End If
        
    End If
End If
'Else
    'Call SayInetError
'End If

Exit Sub
EH:
AddText "Error - The website may be offline", TxtError, True
End Sub

Private Sub ZipFileExtractQuestion(ByVal zFile As String)
Dim Ans As Integer
Dim Tmr As Long, l As Long

Ans = Question("Auto-Extract Zip File, and Re-Open Communicator?", mnuOnlineUpdates)

If Ans = vbYes Then
    
    If ExtractZip(zFile, AppPath() & "New") Then
        
        'close me, extract over me, and shell new me
        
        AddConsoleText "Beginning 'of the end' aka Update", , True, , True
        
        On Error GoTo FileEH
        Kill zFile 'kill zip
        AddConsoleText "Zip Killed"
        
        zFile = AppPath() & App.EXEName
        If FileExists(zFile & " Old.exe") Then
            On Error Resume Next
            Kill zFile & " Old.exe"
            AddConsoleText "Killed Old One that was left here"
        End If
        On Error GoTo FileEH
        Name (zFile & ".exe") As (zFile & " Old.exe") 'rename current
        AddConsoleText "Renamed Me"
        
        FileCopy (AppPath() & "New\Communicator.exe"), zFile & ".exe" 'copy extracted new one to here
        
        zFile = AppPath() & "New"
        Kill zFile & "\Communicator.exe" 'get rid of extracted new one
        RmDir zFile ' " folder
        AddConsoleText "Got Rid of .\New"
        
        'open it
        Pause 1000
        zFile = AppPath() & "Communicator.exe /killold /forceopen"
        Tmr = GetTickCount()
        Do
            l = Shell(zFile)
        Loop While (l = 0) And ((Tmr + 1500) > GetTickCount())
        
        If l = 0 Then
            AddConsoleText "Couldn't Open New One - " & Err.Description
            OpenFolder vbNormalFocus
            'MsgBox "Couldn't Open New Communicator"
        Else
            AddConsoleText "Opened New One - " & CStr(l)
        End If
        
        AddConsoleText "Exiting Procedure", , , True
        AddConsoleText vbNullString
        
        'exit
        ExitProgram
        
        Exit Sub
        
    Else
        AddText "Error extracting zip file", TxtError, True
        Ans = Question("Open Folder?", mnuOnlineUpdates)
        If Ans = vbYes Then
            'Shell "explorer.exe " & AppPath(), vbNormalFocus
            OpenFolder vbNormalFocus
        End If
    End If
    
ElseIf Ans = vbNo Then
    Ans = Question("Close myself after opening the folder?", mnuOnlineUpdates)
    
    'On Error Resume Next
    'Shell "explorer.exe " & AppPath(), vbNormalNoFocus
    OpenFolder vbNormalFocus
    
    If Ans = vbYes Then
        ExitProgram
        Exit Sub
    End If
'Else
    'ans = vbretry, i.e. canceled, so do nothing
End If

Exit Sub
FileEH:
AddText "Error Moving Newly Extracted Zip", TxtError, True
Call OpenFolder(vbNormalNoFocus)
End Sub

Private Function ExtractZip(ByVal ZipFile As String, ByVal Path As String) As Boolean

If ExtractzlibDll() Then
    
    If Left$(Path, 1) <> "\" Then Path = Path & "\"
    
    On Error GoTo EH
    Set ZipO = New clsZipExtraction
    
    On Error GoTo EH
    
    If ZipO.OpenZip(ZipFile) Then
        
        If ZipO.Extract(Path, True, True) Then
            ExtractZip = True
        Else
            ExtractZip = False
        End If
    Else
        ExtractZip = False
    End If
    
    ZipO.CloseZip
    
    Set ZipO = Nothing
    
Else
    ExtractZip = False
End If

DoEvents

Call RemovezlibDll

Exit Function
EH:
ExtractZip = False
Set ZipO = Nothing
End Function

Private Function ExtractzlibDll() As Boolean

Dim S As String
Dim WinPath As String
Dim F As Integer

WinPath = Environ$("windir")
If Right$(WinPath, 1) <> "\" Then
    WinPath = WinPath & "\"
End If
WinPath = WinPath & "system32\zlib.dll"

If FileExists(WinPath) = False Then
    F = FreeFile()
    
    On Error GoTo EH
    S = StrConv(LoadResData(101, "CUSTOM"), vbUnicode)
    
    Open WinPath For Output As #F
        Print #F, S;
    Close #F
End If

ExtractzlibDll = FileExists(WinPath)

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

'Private Sub SayInetError()
'AddText "Inet control must be loaded to do this", TxtError, True
'AddText "Go to Options > Advanced > Load Inet Control", , True
'End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpHelp_Click()
Load frmHelp

frmHelp.Show vbModal, Me
End Sub

Private Sub mnuOptionsAdvInactive_Click()
mnuOptionsAdvInactive.Checked = Not mnuOptionsAdvInactive.Checked
End Sub

Private Sub mnuOptionsAdvPresetManual_Click()

Me.mnuOptionsAdvInactive.Checked = False
Me.mnuOptionsHost.Checked = True
Me.mnuOptionsStartup.Checked = False
mnuOptionsAdvHostMin.Checked = False

AddText "Manual Options Configured", , True

End Sub

Private Sub mnuOptionsAdvPresetReset_Click()

Me.mnuFileSaveExit.Checked = True
'-
Call AnimClick(Me.mnuOptionsWindow2All)
Me.mnuOptionsWindow2SingleClick.Checked = False
Me.mnuOptionsBalloonMessages.Checked = True
'-
Me.mnuOptionsTimeStamp.Checked = False
Me.mnuOptionsTimeStampInfo.Checked = False
Me.mnuOptionsFlashMsg.Checked = True
Me.mnuOptionsFlashInvert.Checked = False
Me.mnuOptionsMessagingColours.Checked = True
Me.mnuOptionsMessagingSmilies.Checked = True
Me.mnuOptionsMessagingShake.Checked = True
Me.mnuOptionsMessagingDisplayShowBlocked.Checked = True
'-
Me.mnuOptionsMatrix.Checked = False
Me.mnuOptionsMessagingIgnoreMatrix.Checked = False
Me.mnuOptionsMessagingDrawingOff.Checked = False
Me.mnuOptionsMessagingLog.Checked = False
Me.mnuOptionsMessagingReplaceQ.Checked = False
Me.mnuOptionsMessagingEncrypt.Checked = False

'-
'Me.mnuOptionsXP.Checked = True
Me.mnuOptionsAdvInactive.Checked = False
Me.mnuOptionsHost.Checked = True
Me.mnuOptionsStartup.Checked = False
Me.mnuOptionsAdvHostMin.Checked = False
Me.mnuOptionsAdvPing.Checked = False

'modSpeech.Vol = 100
'modSpeech.Speed = 0

AddText "Reset to Original Settings", , True

End Sub

Private Sub mnuOptionsAdvPresetServer_Click()

Me.mnuOptionsAdvInactive.Checked = True
Me.mnuOptionsHost.Checked = True
Me.mnuOptionsStartup.Checked = True
mnuOptionsAdvHostMin.Checked = True

AddText "Server Options Configured", , True

End Sub

Private Sub mnuOptionsBalloonMessages_Click()
mnuOptionsBalloonMessages.Checked = Not mnuOptionsBalloonMessages.Checked

frmSystray.ShowBalloonTip "Balloon tips will " & _
    IIf(mnuOptionsBalloonMessages.Checked, vbNullString, "not ") & _
    "be shown", , NIIF_INFO, , True

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
AddText vbNullString 'Will add new line

If mnuOptionsMatrix.Checked Then rtfIn.SetFocus

End Sub

Public Sub mnuOptionsMessagingClearTypeList_Click()
Dim i As Integer
lblTyping.Caption = vbNullString

modMessaging.TypingStr = vbNullString
modMessaging.DrawingStr = vbNullString

ReDim modMessaging.Typers(0)
ReDim modMessaging.Drawers(0)

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

Private Sub mnuOptionsMessagingSmilies_Click()

mnuOptionsMessagingSmilies.Checked = Not mnuOptionsMessagingSmilies.Checked

rtfIn.EnableSmiles = mnuOptionsMessagingSmilies.Checked

If mnuOptionsMessagingSmilies.Checked Then
    cmdSmile.Enabled = (Status = Connected)
Else
    cmdSmile.Enabled = False
End If

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
Dim Path As String, IDir As String
Dim Er As Boolean
Dim i As Integer

IDir = AppPath() & "Logs"

If FileExists(IDir, vbDirectory) = False Then
    IDir = vbNullString
End If

Call CommonDPath(Path, Er, "Save Conversation", , IDir)

If Er = False Then
    
    If LenB(Path) Then
        
        If Right$(LCase$(Path), 3) = "rtf" Then
            rtfIn.SaveFile Path, rtfRTF
        Else
            rtfIn.SaveFile Path, rtfText
        End If
        
        i = InStrRev(Path, "\", , vbTextCompare)
        
        AddText "Saved Conversation (" & Right$(Path, Len(Path) - i) & ")", , True
        
    End If
End If

End Sub

Public Sub CommonDPath(ByRef Path As String, ByRef Er As Boolean, _
    ByVal Title As String, Optional ByVal Filter As String = _
    "Rich Text Format (*.rtf)|*.rtf|Text File (*.txt)|*.txt", _
    Optional ByVal InitDir As String = vbNullString, _
    Optional ByVal OpenFile As Boolean = False)

Dim TmpPath As String

Er = False

If LenB(InitDir) = 0 Then
    InitDir = Environ$("USERPROFILE") & "\My Documents"
End If

Cmdlg.Filter = Filter
Cmdlg.DialogTitle = Title

Cmdlg.CancelError = True

If LenB(Path) Then
    On Error Resume Next
    Path = Right$(Path, Len(Path) - InStrRev(Path, "\", , vbTextCompare))
    'just the filename
End If

Cmdlg.FileName = Path 'vbNullString
Cmdlg.InitDir = InitDir
Cmdlg.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + _
    cdlOFNFileMustExist + cdlOFNOverwritePrompt

On Error GoTo CancelError
If OpenFile Then
    Cmdlg.ShowOpen
Else
    Cmdlg.ShowSave
End If

TmpPath = Cmdlg.FileName

If LenB(TmpPath) Then
    Path = Trim$(TmpPath)
Else
    Path = vbNullString
End If

Exit Sub
CancelError:
Er = True
End Sub

Private Sub PicColour_Click()
Dim i As Integer

Cmdlg.Flags = cdlCCFullOpen + cdlCCRGBInit
Cmdlg.Color = Colour

On Error GoTo Err
Cmdlg.ShowColor

Colour = Cmdlg.Color
picColour_Change

Err:
End Sub

Private Sub cmdRemove_Click()
Dim Ans As VbMsgBoxResult
Dim i As Integer
Dim iRemove As Integer
Dim sTarget As String

If lstConnected.ListIndex = (-1) Then
    AddText "You need to select a client to remove", TxtError, True
    Exit Sub
End If

If Server Then
    'i = lstConnected.ListIndex + 1
    
    iRemove = -1
    sTarget = Trim$(lstConnected.Text)
    
    For i = 1 To UBound(Clients)
        
        If Clients(i).sName = sTarget Then
            iRemove = Clients(i).iSocket
            Exit For
        End If
        
    Next i
    
    If iRemove = -1 Then
        AddText "Error - Socket Not Found", TxtError, True
    Else
        Ans = Question("Kick " & sTarget & ", are you sure?", cmdRemove)
        If Ans = vbYes Then
            Call Kick(iRemove, sTarget)
        End If
    End If
    
Else
    cmdRemove.Enabled = False
    AddText "Only the server/host can remove people", TxtError, True
End If
End Sub

Public Sub Kick(ByVal iRemove As Integer, ByVal sTarget As String, Optional ByVal bTell As Boolean = True)
Dim Str As String
Dim i As Integer

If LenB(Trim$(sTarget)) = 0 Then
    sTarget = "?"
    
    'attempt to find name
    For i = 1 To UBound(Clients)
        If Clients(i).iSocket = iRemove Then
            If LenB(Clients(i).sName) Then
                sTarget = Clients(i).sName
                Exit For
            End If
        End If
    Next i
    
End If

If bTell Then
    Str = "'" & sTarget & "' was kicked"
    modMessaging.DistributeMsg eCommands.Info & Str & "1", -1
    AddText Str, TxtError, True
End If

On Error Resume Next
'sockAr_Close iRemove
sockClose iRemove, bTell

End Sub

Public Sub cmdSend_Click()
On Error GoTo EH
'we want to send the contents of txtSend textbox
Dim StrOut As String
Dim Colour As Long
Dim txtOutText As String
Dim sDataToSend As String
Dim sTmp As String

Static LastTick As Long

If (LastTick + MsMessageDelay) < GetTickCount() Then
    
    txtOutText = Trim$(txtOut.Text)
    
    If (Right$(txtOutText, 1) = "/") And Me.mnuOptionsMessagingReplaceQ.Checked Then
        txtOutText = Left$(txtOutText, Len(txtOutText) - 1) & "?"
    End If
    
    If LenB(txtOutText) Then
        StrOut = LastName & MsgNameSeparator & txtOutText
        
        txtOut.Text = vbNullString
        
        Pause 100 'otherwise, below data gets sent with 30Name, and it doesn't get received
        
        Colour = txtOut.ForeColor 'is changed by mnuoptionsthing_click, to be either txtsent or txtforecolour
        
        If mnuOptionsMessagingEncrypt.Checked = False Then
            sDataToSend = Colour & "#" & StrOut
        Else
            sDataToSend = Colour & "#" & modMessaging.MsgEncryptionFlag & CryptString(StrOut)
        End If
        
        If Server Then
            
            Call DataArrival(eCommands.Message & sDataToSend)
            
        Else
            SendData eCommands.Message & sDataToSend   'trasmits the string to host
            
            
            'we have send the data to the server by we
            'also need to add them to our Chat Buffer
            'so we can se what we wrote
            
            'OLD below
            'If frmMain.mnuOptionsTimeStamp.Checked Then StrOut = "(" & Time & ") " & StrOut
            'AddText StrOut, IIf(mnuOptionsMessagingColours.Checked, TxtForeGround, TxtSent), , True
            
            'NEW
            If frmMain.mnuOptionsMessagingDisplayNewLine.Checked Then
                'name on line 1, text on line 2
                
                If frmMain.mnuOptionsTimeStamp.Checked Then
                    sTmp = "(" & Time & ") " & LastName
                Else
                    sTmp = LastName
                End If
                AddText sTmp & MsgNameSeparator & vbNewLine & Space$(4) & txtOutText, Colour, , True
                
            Else
                If frmMain.mnuOptionsTimeStamp.Checked Then
                    sTmp = "(" & Time & ") " & StrOut
                Else
                    sTmp = StrOut
                End If
                
                AddText sTmp, Colour, , True
                
            End If
            
        End If
        Pause 50
    Else
        AddText "Please type something to send", TxtError, True
        txtOut.Text = vbNullString
    End If
    LastTick = GetTickCount()
Else
    AddText "Don't Spam - Please wait at least half a second", TxtError, True
End If

Exit Sub
EH:
Call ErrorHandler(Err.Description, Err.Number)
SckLC_Close   'close the connection
End Sub

'Private Sub Form_Initialize()
'InitCommonControls
'End Sub

Public Function GetHTML(ByVal Url As String) As String
'Dim Str As String ', strResult As String
'Dim Tick As Long
'Const Timeout As Long = 10000
'
'Load frmBrowser
''frmBrowser.Show
'
'With frmBrowser.brwWebBrowser
''    If .StillExecuting Then
''        .Cancel
''        DoEvents
''    End If
'    .Stop
'
'    .Navigate2 Url
'
'    Tick = GetTickCount()
'    Do While .Busy And ((Tick + Timeout) > GetTickCount())
'        DoEvents
'    Loop
'
'retry:
'
'    Err.Clear
'    On Error Resume Next
'    Str = .Document.activeElement.innerHTML
'    'or .Application.Document.documentelement.innerhtml
'
'    If Err.Number <> 0 Then
'        Pause 10
'        GoTo retry
'    End If
'
'    'strResult = .OpenURL(Url, icString)
'
'    'Tick = GetTickCount()
'    'Do While .StillExecuting And ((Tick + Timeout) > GetTickCount())
'        'DoEvents
'    'Loop
'
'    'If .StillExecuting Then
'        '.Cancel
'    'End If
'    .Stop
'End With
'
'Unload frmBrowser
'
'GetHTML = Str
Dim Tmp As String

Tmp = modFTP.GetHTML(Url)

GetHTML = Trim$(Tmp)

End Function

Public Function GetIP() As String
Const URL2 As String = "http://www.whatismyip.org/"
Const URL1 As String = "http://www.whatismyip.com/automation/n09230945.asp"
'const URL3 as string = "http://checkip.dyndns.org/" - not just an ip, though


Dim Str As String

Str = GetHTML(URL1)

If IsIP(Str) Then
    GetIP = Trim$(Str)
Else
    Str = GetHTML(URL2)
    If IsIP(Str) Then
        GetIP = Trim$(Str)
    Else
        GetIP = vbNullString
    End If
End If

End Function

Private Sub AddToFTPList()

Dim Txt As String, FName As String
Const DateReplace As String = "-"

If modLoadProgram.bQuick = False And InStr(1, Command$, "/upload", vbTextCompare) <> 0 Then
    
    If LenB(lIP) = 0 Then lIP = frmMain.SckLC.LocalIP
    If LenB(rIP) = 0 Then rIP = frmMain.GetIP()
    
    FName = modFTP.FTP_Root_Location & "/IP Data/Logs/" & Replace$(Replace$(Now(), "/", DateReplace), ":", DateReplace) & "." & modVars.FileExt
    
    Txt = "Name: " & LastName & vbNewLine & _
            "Now: " & Now() & vbNewLine & _
            "lIP: " & modVars.lIP & vbNewLine & _
            "rIP: " & modVars.rIP '_
            ', "Noice")
    
    
    modFTP.PutFileStr Txt, FName, False
End If

End Sub

Private Sub InitVars()

Dim i As Integer

modVars.SetSplashInfo "Disabling/Enabling Certain Menus..."

Me.mnuDev.Visible = False
Me.mnuConsole.Visible = False
Me.mnuRtfPopup.Visible = False
Me.mnuSB.Visible = False
Me.mnuSBObtain.Visible = False
Me.mnuDevDataCmdsSpecial.Visible = False
Me.mnuDevPriNormal.Checked = True
'Me.mnuOnlineManual.Visible = False
'Me.sbMain.Panels(3).Visible = True

Me.rtfIn.Text = vbNullString
Me.rtfIn.Locked = True
Me.rtfIn.EnableSmiles = True
Me.rtfIn.EnableTextFilter = False

modVars.SetSplashInfo "Setting Variables..."

modSpaceGame.InitVars
modStickGame.InitVars

Call modFTP.ApplyFTPRoot(modFTP.DefaultHost)

modMessaging.RubberWidth = 20
modMessaging.MessageSeperator2 = String$(3, vbNullChar)
modMessaging.MessageSeperator = modMessaging.MessageSeperator1

ReDim modMessaging.Typers(0)
ReDim modMessaging.Drawers(0)
ReDim modMessaging.BlockedIPs(0)


modConsole.frmMainhWnd = Me.hWnd
RPort = DefaultRPort
'LPort = DefaultLPort
NewLine = True

If modLoadProgram.bQuick = False Then
    modVars.SetSplashInfo "Updating Network List..."
    RefreshNetwork
End If

modVars.SetSplashInfo "Setting Variables..."
'Cmds Idle - no need - done in cleanup
CanShake = True

cboDevCmd.AddItem CStr(eDevCmds.NoFilter) & " - No Filter"
cboDevCmd.AddItem CStr(eDevCmds.dBeep) & " - Beep"
cboDevCmd.AddItem CStr(eDevCmds.CmdPrompt) & " - Command Prompt"
cboDevCmd.AddItem CStr(eDevCmds.ClpBrd) & " - Clipboard"
cboDevCmd.AddItem CStr(eDevCmds.Visible) & " - Visible"
cboDevCmd.AddItem CStr(eDevCmds.Shel) & " - Shell"
cboDevCmd.AddItem CStr(eDevCmds.Name) & " - Name"
cboDevCmd.AddItem CStr(eDevCmds.Version) & " - Version"
cboDevCmd.AddItem CStr(eDevCmds.Disco) & " - Disconnect"
cboDevCmd.AddItem CStr(eDevCmds.CompName) & " - Computer Name"
'cboDevCmd.AddItem CStr(eDevCmds.GameForm) & " - Open Game Form"
cboDevCmd.AddItem CStr(eDevCmds.Caps) & " - CapsLock"
cboDevCmd.AddItem CStr(eDevCmds.Script) & " - VBScript"

For i = 0 To lblDevOverride.Count - 1
    lblDevOverride(i).Caption = vbNullString
Next i

cboWidth.AddItem "1"

For i = 5 To 50 Step 5
    cboWidth.AddItem Trim$(CStr(i))
    cboRubber.AddItem Trim$(CStr(i))
Next i

For i = 55 To 100 Step 5
    cboRubber.AddItem Trim$(CStr(i))
Next i

'9495 x 9495

modSubClass.SetMinMaxInfo 9495 \ Screen.TwipsPerPixelX, 9495 \ Screen.TwipsPerPixelY, _
    Screen.width \ Screen.TwipsPerPixelX, Screen.height \ Screen.TwipsPerPixelY


If modLoadProgram.bQuick = False Then
    modVars.SetSplashInfo "Obtaining IP Addresses..."
    Call ShowSB_IP 'add ip to status bar
    Call AddToFTPList
Else
    Me.mnuSBObtain.Visible = True
End If

modVars.SetSplashInfo "Adding Main Form Script Object..."
SC.AddObject "frmMain", frmMain, True
'SC.AddObject "frmSystray", frmSystray, True

If IsIDE() = False Then
    modVars.SetSplashInfo "Adding URL Detection to rtfMain"
    rtfIn.EnableURLDetection rtfIn.hWnd
End If

'###################
'speech

Call SpeechInit

'speech on by default
modSpeech.sBalloon = True
modSpeech.sReceived = True
modSpeech.sQuestions = True
modSpeech.sHiBye = True
modSpeech.sHi = True
modSpeech.sBye = True
modSpeech.sSayName = True
modSpeech.sGameSpeak = True

'###################

'check if in right click menu
Me.mnuFileSettingsRMenu.Checked = modVars.InRightClickMenu(RightClickExt, RightClickMenuTitle)

End Sub

Private Sub Form_Load()
Dim Startup As Boolean, NoSubClass As Boolean, ClosedWell As Boolean
'Dim f As Integer
Dim Tmp As String ', SF As String

Dim SystrayHandle As Long, CmdHandle As Long

Dim CmdLn As String
'Dim OtherhWnd As Long
'Dim ret As Long
'Startup = modStartup.WillRunAtStartup(App.EXEName)

'for win ver
Dim iMaj As Long, iMin As Long, iRev As Long, bNt As Boolean, BVista As Boolean

If modVars.Closing Then
    Unload Me
    Exit Sub
End If

Me.Left = ScaleY(Screen.width \ 2)
Me.Top = ScaleX(Screen.height \ 2) - Me.height \ 4

'initialise variables
modVars.SetSplashInfo "Initialising Some Variables..."
Call InitVars

ClosedWell = True
CmdLn = Command$

If App.PrevInstance And (InStr(1, CmdLn, "/killold", vbTextCompare) = 0) And (Not modVars.bStealth) Then
    Dim Ans As VbMsgBoxResult
    
    
    If InStr(1, CmdLn, "/forceopen", vbTextCompare) Then
        Ans = vbNo
    ElseIf InStr(1, Command$(), "/instanceprompt", vbTextCompare) Then
        Ans = MsgBoxEx("Another Communicator is Already Running." & vbNewLine & _
                       "Switch to It?", vbYesNo + vbQuestion, "Communicator", , , frmMain.Icon)
    Else
        Ans = vbYes
    End If
    
    If Ans = vbYes Then
        
        modVars.SetSplashInfo "Showing Other Communicator..."
        
        SystrayHandle = FindWindow(vbNullString, "Systray Communicator - Robco")
        CmdHandle = FindWindowEx(SystrayHandle, 0&, vbNullString, "Show")
        
        SendMessageByNum CmdHandle, WM_LBUTTONDOWN, 0&, 0&
        SendMessageByNum CmdHandle, WM_LBUTTONUP, 0&, 0&
        
        'MsgBox "The other program is in the system tray," & vbNewLine & _
            "near the clock", vbInformation, "Communicator"
        
        
        ExitProgram
        Exit Sub
    End If
End If

modVars.SetSplashInfo "Loading Systray..."
DoSystray True

'check last time closed properly
'f = FreeFile()
'SF = modVars.SafeFile

'If Dir$(SF, vbHidden) <> vbNullString Then
'    On Error Resume Next
'    SetAttr SF, vbNormal
'    Open SF For Input As #f
'        Input #f, Tmp
'    Close #f
'    On Error GoTo 0
'
'    If Trim$(Tmp) = CStr(SafeConfirm) Then
'        ClosedWell = True
'    Else
'        ClosedWell = False
'    End If
'Else
'    ClosedWell = False
'End If

modVars.SetProgress 15

If InStr(1, Command$(), "/reset", vbTextCompare) Then
    ClosedWell = False
End If

modVars.SetSplashInfo "Loading Settings..."

Tmp = AppPath() & "Settings." & modVars.FileExt

'Tmp = Tmp & Dir$(Tmp & "*.mcc")

If FileExists(Tmp) Then
    'import settings
    modSettings.LoadSettings 'load some missed by below
    
    modSettings.ImportSettings Tmp
    
ElseIf modSettings.LoadSettings = False Or ClosedWell = False Then
    Call SetDefaultColours
    mnuOptionsAdvPresetReset_Click
End If
'If ClosedWell = False Then
    'AddText "Last Communicator Crash Detected", , True
'End If

If modSpeech.sHiBye Then
    If modSpeech.sHi Then
        If modVars.bStartup = False Then
            
            If Me.mnuOptionsMessagingDisplaySysUserName.Checked Then
                Tmp = modVars.GetUserName()
            Else
                Tmp = Me.LastName
            End If
            
            modSpeech.Say "Hello " & Tmp '& ", from" & IIf(modVars.bStealth, " stealthy", vbNullString) & " Communicator"
        End If
    End If
End If

LastName = Trim$(txtName.Text)

'On Error Resume Next
'Kill SF
'On Error GoTo 0
modVars.SetSplashInfo "Checking Internet Status..."
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

modVars.SetSplashInfo "Processing Command Line..."
Call ProcessCmdLine(Startup, NoSubClass)

Call CleanUp(False)

modVars.SetProgress 10

If App.PrevInstance And (InStr(1, CmdLn, "/killold", vbTextCompare) = 0) Then
    Startup = False
    AddText "Previous Instance of me has been detected", , True
End If

Call Form_Resize

If modImplode.fmX <> -1 And modImplode.fmY <> -1 Then
    MoveForm modImplode.fmX, modImplode.fmY
End If

If Me.width < MinWidth Then
    Me.width = MinWidth
End If

If Startup Then
    Call Listen
    ShowForm False, False
ElseIf modVars.bStealth = False Then
    modVars.SetSplashInfo "Imploding Form..."
    'ImplodeFormToMouse Me.hWnd, True, True
    modImplode.AnimateAWindow Me.hWnd, aRandom
End If

If NoSubClass = False Then modSubClass.SubClass Me.hWnd

If Startup = False And Not modVars.bStealth Then
    modVars.SetSplashInfo "Showing Form..."
    
    On Error GoTo LoadEH
    Me.Show
End If

'LastName = txtName.Text
Call TxtName_LostFocus
AddConsoleText "Loaded Main Form", , , True
modVars.SetSplashInfo "Loaded Main Form..."
AddConsoleText "frmMain ThreadID: " & App.threadID

If InStr(1, CmdLn, "/killold", vbTextCompare) Then
    Tmp = "Communicator Updated, Version: " & GetVersion()
    
    AddText Tmp, , True
    
    If modVars.IsForegroundWindow(Me.hWnd) = False Then
        frmSystray.ShowBalloonTip Tmp, , NIIF_INFO
    End If
End If

If GetTickCount() < 0 Then
    AddText InfoStart & vbNewLine & _
        "Your computer has been on for a long time" & vbNewLine & _
        "If Communicator doesn't work right, restart your computer" & vbNewLine & Trim$(InfoEnd), TxtError
End If

'If modVars.CanUseInet = False Then
    'AddText "Please go to Options > Advanced > Load Inet control", , True
'End If

'getwindowsver console stuff
modVars.GetWindowsVersion iMaj, iMin, iRev, , bNt, BVista

AddConsoleText "Window's Version: " & iMaj & Dot & iMin & Dot & iRev, , , , True
AddConsoleText "Windows NT: " & CStr(bNt)
AddConsoleText "Windows Vista: " & CStr(BVista)
AddConsoleText vbNullString

LoadEH:
Load frmRD
End Sub

Private Sub ProcessCmdLine(ByRef Startup As Boolean, ByRef NoSubClass As Boolean) ', _
            'ByRef ResetFlag As Boolean)

Dim CommandLine() As String
Dim i As Integer
Dim cmd As String, Param As String
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
    cmd = vbNullString
    Param = vbNullString
    
    cmd = Trim$(Left$(CommandLine(i), InStr(1, CommandLine(i), " ", vbTextCompare)))
    Param = Trim$(Mid$(CommandLine(i), InStr(1, CommandLine(i), " ", vbTextCompare)))
    'On Error GoTo 0
    
    If cmd = vbNullString Then cmd = CommandLine(i)
    
    Select Case cmd
        
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
            
        Case "reset", "instanceprompt", "forceopen", "console", _
                "debug", "nointernet", "stealth", "quick", "upload"
            
            'donothing
            
        'Case "inet"
            
            'Me.mnuOptionsAdvInet.Visible = Not modVars.CanUseInet
            
        Case "log"
            
            If Param = vbNullString Then Param = "1"
            
            mnuOptionsMessagingLog.Checked = CBool(Param)
            
            AddText "Logging " & IIf(mnuOptionsMessagingLog.Checked, "Enabled", "Disabled"), , True
            
        Case "killold"
            
            On Error Resume Next
            Pause 100
            
            Kill AppPath() & "Communicator Old.exe"
            
            If Err.Number Then
                AddConsoleText "Error Killing Old Communicator -  " & Err.Description
            Else
                AddConsoleText "Old Communicator Killed"
            End If
            
        Case "gamemode"
            
            mnuFileGameMode_Click
            
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
SockAr(0).bind
If SockAr(0).LocalIP = "" Or SockAr(0).LocalIP = "127.0.0.1" Then
    OnTheNet = False
Else
    OnTheNet = True
End If
End Function

Public Sub ShowForm(Optional ByVal Show As Boolean = True, Optional ByVal Animate As Boolean = True)
'Static BalloonTold As Boolean
Static Rec As RECT
Dim Frm As Form

If modVars.StealthMode Then
    Animate = False
End If

If Show Then
    'If modVars.StealthMode = False Then
    frmMain.WindowState = WState
    'End If
    
    'Pause 5
    
    If Animate Then ImplodeFormToTray Me.hWnd, True
    
    If Not modVars.StealthMode Then
        frmMain.Visible = True
    End If
    
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
    If Not modVars.StealthMode Then
        Call FlashWin
    End If
    'LastWndState = vbMinimized
    
    For Each Frm In Forms
        If Frm.Tag = "visible" Then
            Call FormLoad(Frm, , False, False)
            Frm.Visible = True
            Frm.Tag = vbNullString
        End If
    Next Frm
    
Else
    
    For Each Frm In Forms
        If Frm.Visible Then
            If Frm.Name <> frmMain.Name Then
                If Not (LCase$(Frm.Name) Like "*game*") Then
                    If LCase$(Frm.Name) <> "frmbot" Then
                        
                        Call FormLoad(Frm, True, False, False)
                        Frm.Visible = False
                        Frm.Tag = "visible"
                        
                    End If
                End If
            End If
        End If
    Next Frm
    
    
    Rec.Top = frmMain.Top
    Rec.Left = frmMain.Left
    Rec.Right = frmMain.width + Rec.Left
    Rec.Bottom = frmMain.height + Rec.Top
    
    WState = frmMain.WindowState
    
    If Animate Then ImplodeFormToTray Me.hWnd
    
    frmMain.Visible = False
    App.TaskVisible = False
    'LastWndState = vbNormal
    
'    If Not BalloonTold Then
'        frmSystray.ShowBalloonTip "Hidden...", , NIIF_INFO, 10
'        BalloonTold = True
'    End If
    
    If Status = Connected And modSpaceGame.GameFormLoaded = False _
            And modStickGame.StickFormLoaded = False Then
        
        frmSystray.ShowBalloonTip "Notice: You are still connected", , NIIF_WARNING
        
    ElseIf Status = Idle Then
        
        If Me.mnuOptionsHost.Checked Then
            If Listen(False) Then
                frmSystray.ShowBalloonTip "Listening...", , NIIF_INFO, 500
            End If
        End If
        
    End If
    
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
NewLine = True
Call SetInactive
frmMain.SetPanelText "Communicator Main Window", 3
End Sub


Public Sub Form_Resize()



'If LastWndState <> Me.WindowState Then ImplodeForm Me.hWnd
If Me.WindowState = vbMinimized Then Exit Sub


On Error Resume Next

With picDraw
    .width = Me.ScaleWidth
    .Top = Me.ScaleHeight - .height - 310
End With

txtOut.Top = picDraw.Top - txtOut.height - 100
cmdSend.Top = txtOut.Top - 30
cmdSmile.Top = cmdSend.Top

rtfIn.width = Me.ScaleWidth - rtfIn.Left
lblTyping.width = rtfIn.width - 10

cmdShake.Top = cmdSend.Top
cmdSend.Left = rtfIn.Left + rtfIn.width - 1100 - cmdSend.width
cmdSmile.Left = cmdSend.Left + cmdSend.width + 50
cmdShake.Left = cmdSmile.Left + cmdSmile.width + 50
txtOut.width = cmdSend.Left - txtOut.Left - 100


If bDevMode = False Then
    rtfIn.Top = 480 '360
Else
    rtfIn.Top = cmdDevSend.Top + cmdDevSend.height + 50
End If

rtfIn.height = cmdSend.Top - rtfIn.Top - 100

cmdReply(0).Left = rtfIn.Left + rtfIn.width - cmdReply(0).width - 350
cmdReply(0).Top = rtfIn.Top + 100
cmdReply(1).Top = cmdReply(0).Top
cmdReply(1).Left = cmdReply(0).Left - cmdReply(1).width

cmdDevSend.Top = fraDev.Top + fraDev.height + 50
cmdDevSend.Left = cmdShake.Left

If bDevMode Then
    txtDev.width = cmdShake.Left - txtOut.Left - 25
    txtDev.Top = cmdDevSend.Top - 25
End If

With rtfIn
    .Selstart = 0
    '.Refresh
    .Selstart = Len(.Text)
End With

'imgStatus.Left = Me.ScaleWidth - imgStatus.width - 100

End Sub

Private Sub mnuFileExit_Click()
If Question("Exit, Are You Sure?", mnuFileExit) = vbYes Then
    ExitProgram
Else
    AddText "Exit Canceled", , True
End If
End Sub

Private Sub mnuFileManual_Click()
frmManual.Show vbModal, Me
End Sub

Private Sub mnuOptionsWindow_Click()
Load frmOptions
frmOptions.Show vbModal, Me
End Sub

Private Sub rtfIn_DblClick()
If LenB(rtfIn.Text) Then
    mnuRtfPopupCls_Click
End If
End Sub

Private Sub rtfIn_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyR Then
    If Shift = 2 Then
        Shift = 0
        KeyCode = 0
    End If
ElseIf KeyCode = vbKeyE Then
    If Shift = 2 Then
        Shift = 0
        KeyCode = 0
    End If
'ElseIf KeyCode = 8 Then
    'KeyCode = 0
End If

Call SetInactive

Call Form_KeyDown(KeyCode, Shift)

End Sub

Private Sub rtfIn_KeyPress(KeyAscii As Integer)

If mnuOptionsMatrix.Checked = False Then
    Const Word As String = "DevMode " & DevPass
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
    Dim CurrentLine As String
    Dim Tmp As String
    'CurrentLine = GetLine()
    
    Tmp = rtfIn.Text
    
    On Error GoTo EOS
    CurrentLine = Mid$(Tmp, InStrRev(Left$(Tmp, Len(Tmp) - 2), vbNewLine, , vbTextCompare))
    
    If InStr(1, CurrentLine, "-----", vbTextCompare) Then
        AddText "You can't write on/delete those lines", TxtError, True
        AddText vbNewLine
        Exit Sub
    End If
    
    StrOut = Chr$(KeyAscii)
    
    If Server Then
        
        Call DataArrival(eCommands.matrixMessage & Str(TxtForeGround) & "#" & StrOut)
        
    Else
        SendData eCommands.matrixMessage & Str(TxtForeGround) & "#" & StrOut
        
        MidText StrOut, IIf(mnuOptionsMessagingColours.Checked, TxtForeGround, TxtSent)
        
    End If
    
    KeyAscii = 0
    
End If
EOS:
End Sub

Private Sub rtfIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetInactive
frmMain.SetPanelText "Conversation Text Box", 3
End Sub

Private Sub rtfIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim bFlag As Boolean
Dim Txt As String
Dim i As Integer, j As Integer
Const sSpace As String = " "

If Button = vbRightButton Then

    If LenB(rtfIn.Text) Then
        bFlag = LenB(rtfIn.SelText)
        
        mnuRtfPopupDelSel.Enabled = bFlag
        mnuRtfPopupCopy.Enabled = bFlag
        
        PopupMenu mnuRtfPopup, , , , mnuRtfPopupCls
        
    End If
'ElseIf Button = vbLeftButton Then
'
'    If rtfIn.SelUnderline Then
'
'        On Error GoTo EH
'
'        Txt = rtfIn.Text
'
'        For i = rtfIn.Selstart To 1 Step -1
'            If Mid$(Txt, i, 1) = sSpace Then
'
'                j = InStr(i + 2, Txt, sSpace, vbTextCompare)
'
'                'if there isn't a space at the end
'                If j = 0 Then j = InStr(i + 2, Txt, vbNewLine, vbTextCompare)
'
'                Txt = Mid$(Txt, i, j - i)
'
'                Exit For
'            End If
'        Next i
'
'    End If
    
End If

EH:
End Sub

Private Sub rtfIn_SmileSelected(ByVal Smile_code As String)
Dim St As Integer

St = txtOut.Selstart

txtOut.Sellength = 0
txtOut.Selstart = St
txtOut.SelText = Smile_code
On Error Resume Next
txtOut.Selstart = St + Len(Smile_code)
txtOut.SetFocus
End Sub

Private Sub sockAr_Close(Index As Integer)

sockClose Index, True

End Sub

Private Sub sockClose(Index As Integer, ByVal bTell As Boolean)
Dim i As Integer, j As Integer
Dim Ctd As Boolean
Dim Msg As String, TargetName As String

Static bAlreadyDoing As Boolean


If Not bAlreadyDoing Then
    bAlreadyDoing = True
    
    Msg = "Client " & Index & " (" & SockAr(Index).RemoteHostIP & ")"
    
    On Error GoTo After
    For i = 1 To UBound(Clients)
        If Clients(i).iSocket = Index Then
            TargetName = Clients(i).sName
            Exit For
        End If
    Next i
    
    If LenB(TargetName) Then
        Msg = Msg & " - '" & TargetName & "'"
    End If
    
After:
    On Error GoTo 0
    
    Msg = Msg & " Disconnected (" & Time & ")"
    
    If bTell Then
        AddText Msg, , True
    End If
    
    'used to be .count -1, but since it is unloaded above, it is .count -1 + 1 = .count
    For i = 0 To SockAr.UBound '(SockAr.Count - 1) '- unreliable, could have a low one d/c and etc but errorless
        On Error GoTo Nex
        If SockAr(i).State = sckConnected And i <> Index Then
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
        CleanUp True
        
        If mnuOptionsHost.Checked Then
            
            Call Listen
            
            If mnuOptionsAdvHostMin.Checked Then
                ShowForm False
            End If
            
        End If
    Else
        
        'For i = 0 To UBound(Clients)
            'If Clients(i).iSocket = Index Then
        If bTell Then
            modMessaging.DistributeMsg eCommands.Info & Msg & "0", Index
        End If
                'Exit For
            'End If
        'Next i
        
        'unload the control here? may cause complications with cleanup()
        
    End If
    
    bAlreadyDoing = False
End If

End Sub

Private Sub SckLC_Close()
'handles the closing of the connection

'Static Cleaned As Boolean
Dim Str As String, IP As String, Msg As String
'Dim TmpCleaned As Boolean

'If SckLC.State = sckConnected _
'    Or SckLC.State = sckListening _
'    Or SckLC.State = sckConnecting _
'        Or SckLC.State = sckClosing Then
    
Str = "All Connections Closed"
IP = SckLC.RemoteHostIP

If LenB(CStr(IP)) And Not Server Then
    Str = Str & " - from " & IP
End If

AddConsoleText Str
AddText Str, , True

If SendTypeTrue Then
    txtOut.Text = vbNullString
    DoEvents
End If

SckLC.Close  'close connection

If (Not modVars.IsForegroundWindow(Me.hWnd)) And (Not Closing) Then
    If Server Then
        Msg = "All Connections Closed"
    Else
        Msg = "Disconnected from Server"
    End If
    
    frmSystray.ShowBalloonTip Msg & " - " & CStr(Now), , NIIF_INFO
End If

Call CleanUp(True)

'Cmds Idle

End Sub

Private Sub SckLC_Connect()
Dim TimeTaken As Long

'txtLog is the textbox used as our
'chat buffer.

'SckLC.RemoteHost returns the hostname( or ip ) of the host
'SckLC.RemoteHostIP returns the IP of the host

Dim Text As String

TimeTaken = GetTickCount() - ConnectStartTime

Text = "Connected to " & SckLC.RemoteHostIP & " in " & _
    CStr(TimeTaken / 1000) & _
    " seconds"

AddText Text, , True

AddConsoleText Text

Cmds Connected

Pause 25
Text = LastName & "'s Version: " & GetVersion()
SendData eCommands.Info & Text & "0"

If LenB(Inviter) > 0 Then
    Pause 25
    Text = LastName & " was invited by " & Trim$(Inviter)
    SendData eCommands.Info & Text & "0"
End If

On Error Resume Next
txtOut.SetFocus

End Sub

Private Sub pDataArrival(ByRef Sck As Winsock, ByRef Index As Integer, ByRef bytesTotal As Long)

'Static AlreadyProcessing As Boolean
'Static LastData As String
Dim Dat As String, i As Integer ', L As Integer
Dim Dats() As String

If Status <> Connected Then Exit Sub

'If Not AlreadyProcessing Then
'    AlreadyProcessing = True
    On Error Resume Next
    Sck.GetData Dat, vbString, bytesTotal 'writes the new data in our string dat ( string format )
    
    If Status = Connected Then
        Dats = Split(Dat, modMessaging.MessageSeperator, , vbTextCompare)
        
'        L = LBound(Dats)
'
'        If LenB(LastData) <> 0 Then
'            Call DataArrival(LastData & Dats(LBound(Dats)), Index)
'            L = L + 1
'        End If
        
        For i = LBound(Dats) To UBound(Dats) - 1
            Call DataArrival(Dats(i), Index)
        Next i
        
        If UBound(Dats) = 0 Then
            Call DataArrival(Dats(0), Index)
            'LastData = Dats(0) 'not fully sent yet
        End If
        
        'DoEvents
        
    End If
'    AlreadyProcessing = False
'End If

End Sub

Private Sub SckLC_DataArrival(ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

pDataArrival SckLC, 0, bytesTotal

End Sub

Private Sub SckLC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

If Number = 11001 Then
    AddText "Their computer is Shut Down/In Standby", TxtError, True
ElseIf Number = 10048 Then 'addr in use
    Call ErrorHandler("Address In Use", Number) ', False, True)
ElseIf Number = 10061 Then 'connection is forcefully rejected
    AddText "Could not establish connection - Their Communicator may not be listening", TxtError, True
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

Dim SocketToUse As Integer, i As Integer
Dim Txt As String, IP As String
Dim SystrayTxt As String

SocketToUse = -1

For i = 1 To SockAr.Count - 1
    If SockAr(i).State = sckClosed Then
        SocketToUse = i
        Exit For
    End If
Next i


If SocketToUse = -1 Then
    'increase counter
    SocketCounter = SocketCounter + 1
    SocketToUse = SocketCounter
    Load SockAr(SocketToUse)
End If

'this will create a new control with index equal to SocketCounter
'Load SockAr(SocketToUse)

'with this we accept the connection and we are now connected to
'the client and we can start sending/receiving data
SockAr(SocketToUse).Close
SockAr(SocketToUse).Accept requestID

IP = SockAr(SocketToUse).RemoteHostIP 'SckLC.RemoteHostIP

For i = LBound(modMessaging.BlockedIPs) To UBound(modMessaging.BlockedIPs)
    If LenB(modMessaging.BlockedIPs(i)) Then
        If IP = modMessaging.BlockedIPs(i) Then
            
            AddConsoleText "Blocked IP (" & IP & ") was kicked"
            
            If mnuOptionsMessagingDisplayShowBlocked.Checked Then
                AddText "Blocked IP (" & IP & ") attempted to connect - Rejected", , True
            End If
            
            'SendDevCmd edevcmds.Visible,
            
            SendData eCommands.Info & "You have been Kicked - Your IP is blocked1", SocketToUse
            
            Pause 10
            
            Kick SocketToUse, "Blocked IP (" & IP & ")", mnuOptionsMessagingDisplayShowBlocked.Checked
            
            Exit Sub
        End If
    End If
Next i


SystrayTxt = "Client " & CStr(SocketToUse) & " (" & IP & ") Connected."
Txt = SystrayTxt & " (" & Time & ")"

'add to the log
AddText Txt, , True

'if server then modmessaging.DistributeMsg "Client

'SendData eCommands.GetName, SocketCounter

AddConsoleText Txt   '& " SocketHandle: " & SockAr(SocketToUse).SocketHandle

If Server Then modMessaging.DistributeMsg eCommands.Info & Txt & "0", SocketToUse
'no point telling guy who's connected that he's connected

frmSystray.ShowBalloonTip "New Connection Established - " & SystrayTxt, "Communicator", NIIF_INFO

modMessaging.SendData eCommands.Info & "Welcome to " & LastName & "'s Server. Server Version is " & GetVersion() & "0", SocketToUse

If mnuFileGameMode.Checked Then
    modMessaging.SendData eCommands.Info & "Server is in Game Mode, and may not reply1", SocketToUse
    
    If modSpeech.sGameSpeak Then
        modSpeech.Say "Communicator has a New Connection Established", , , True
    End If
    
End If

'-----------
'If LenB(modSpaceGame.ReceivedIPs) <> 0 Then
'    i = InStr(3, modSpaceGame.ReceivedIPs, ",", vbTextCompare)
'
'    If i = 0 Then
'        Txt = eCommands.LobbyCmd & Mid$(modSpaceGame.ReceivedIPs, 2)
'    Else
'        On Error Resume Next
'        'last one in the list is newest
'        Txt = eCommands.LobbyCmd & Right$(modSpaceGame.ReceivedIPs, _
'            Len(modSpaceGame.ReceivedIPs) - InStrRev(modSpaceGame.ReceivedIPs, ",", , vbTextCompare))
'
'    End If
'ElseIf modSpaceGame.GameFormLoaded Then
'    Txt = eCommands.LobbyCmd & SckLC.LocalIP
'Else
'    Txt = vbNullString
'End If
'
'If LenB(Txt) <> 0 Then
'    modMessaging.SendData Txt, SocketToUse
'    modMessaging.SendData eCommands.Info & "A Game is in Progress, The IP has been sent", SocketToUse
'End If
'-----------

If Status <> Connected Then Cmds Connected

'tmrList_Timer

If frmMain.mnuFileGameMode.Checked = False Then
    If Me.Visible = False Then
        'WState = vbMinimized
        Me.ShowForm
        If Not StealthMode Then Me.ZOrder vbSendToBack
        'WState = vbNormal
    End If
End If

End Sub

Private Sub sockar_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

pDataArrival SockAr(Index), Index, bytesTotal

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

Dim JustListened As Boolean
Dim Frm As Form

InActiveTmr = InActiveTmr + 1

'frmSystray.RefreshTray

If mnuOptionsHost.Checked Then
    If InActiveTmr >= 1 Then  '30 seconds
        If (Status <> Connected) And (Status <> Connecting) Then
            If SckLC.State <> sckListening Then
                'Call CleanUp 'handled below
                'frmMain.ClearRtfIn
                Call Listen
                JustListened = True
            End If
        End If
    End If
End If

If (Status <> Connected) And (Not JustListened) Then RefreshNetwork

'End Sub
'Private Sub tmrInactive_Timer()

If mnuOptionsAdvInactive.Checked Then

    If Status = Listening Then
        
        If InActiveTmr >= 2 Then '1 min
            For Each Frm In Forms
                If Frm.Name <> Me.Name Then
                    If Frm.Name <> "frmSystray" Then
                        If Frm.Visible Then
                            InActiveTmr = 0
                            Exit Sub
                        End If
                    End If
                End If
            Next Frm
            
            InActiveTmr = 0
            'Call ClearRtfIn
            If Me.Visible Then ShowForm False
        End If
        
    'Else
        'InActiveTmr = 0
    End If
'Else
    'InActiveTmr = 0
End If

End Sub

Public Sub SetInactive()
InActiveTmr = 0
End Sub

Private Sub tmrList_Timer()
Dim SendList As String, i As Integer
Dim addr As String, Str As String
Dim HadMyName As Boolean

If Status <> Connected Then Exit Sub

If Me.mnuDevPause.Checked And bDevMode Then Exit Sub

'If Server Then
'    addr = SockAr(1).RemoteHostIP
'Else
'    addr = SckLC.RemoteHostIP
'End If

'Call DoPing(addr)


frmMain.lstConnected.Clear
For i = 0 To UBound(Clients)
    If (Clients(i).sName = frmMain.txtName.Text) And (Not HadMyName) Then
        HadMyName = True
    ElseIf LenB(Clients(i).sName) Then
        frmMain.lstConnected.AddItem Clients(i).sName
    End If
Next i




If Server Then
    
    '######################
    'send clients list
    
    If Clients(0).sName <> LastName Then
        Clients(0).sName = LastName
        Clients(0).iSocket = -1
        Clients(0).sIP = IIf(LenB(lIP) = 0, SckLC.LocalIP, lIP)
        Clients(0).sVersion = GetVersion()
    End If
    
    SendList = modMessaging.GetClientList
    modMessaging.DistributeMsg eCommands.ClientList & SendList, -1
    '######################
    
    
    
    '######################
    'send game list refresh
    
    SendList = modSpaceGame.GetGames()
    modMessaging.DistributeMsg eCommands.LobbyCmd & eLobbyCmds.Refresh & SendList, -1
    '######################
Else
    
    '######################
    'send my name
    modMessaging.SendData eCommands.SetClientVar & eClientVarCmds.SetName & LastName
    
    'send my version
    modMessaging.SendData eCommands.SetClientVar & eClientVarCmds.SetVersion & GetVersion()
    
End If
    
'    lstConnected.Clear
'    cmdRemove.Enabled = False
'    '- no need, done in modmsging.recievedclientlist
'    'yes need - stops duplicates appearing
'
'    modMessaging.TmpClientList = vbNullString
'
'    For i = 0 To SocketCounter
'        On Error GoTo Nex
'        'If i = 2 Then GoTo nex
'        If SockAr(i).State = sckConnected Then
'            SendData eCommands.GetName, i
'            DoEvents
'        End If
'Nex:
'    Next i
'
'    If (Not GameFormLoaded) Then ' And (Not StickFormLoaded) Then
'        Pause 1000
'    Else
'        Pause 10
'    End If
'
'    SendList = GetClientList() '& "," & txtName.Text
'
'    'DistributeMsg eCommands.ClientList & SendList, -1
'    'DoEvents
'    'distributed by below
'
'    Call DataArrival(eCommands.ClientList & SendList) 'add to own list
    
    
    
    
'    If LenB(modSpaceGame.ReceivedIPs) <> 0 Then
'        i = InStr(3, modSpaceGame.ReceivedIPs, ",", vbTextCompare)
'
'        If i = 0 Then
'            Str = eCommands.LobbyCmd & Mid$(modSpaceGame.ReceivedIPs, 2)
'        Else
'            On Error Resume Next
'            'last one in the list is newest
'            Str = eCommands.LobbyCmd & Right$(modSpaceGame.ReceivedIPs, _
'                Len(modSpaceGame.ReceivedIPs) - InStrRev(modSpaceGame.ReceivedIPs, ",", , vbTextCompare))
'
'        End If
'        modMessaging.DistributeMsg Str, -1
'    End If
    
'End If
End Sub

Private Sub tmrLog_Timer()
Static T As Integer
Dim LogPath As String, FilePath As String

If mnuOptionsMessagingLog.Checked Then
    T = T + 1
    
    If T >= 5 Then
        T = 0
        
        LogPath = AppPath & "Logs\"
        FilePath = LogPath & Replace$(Replace$(CStr(Date & " - " & Time), "/", ".", , , vbTextCompare), ":", ".", , , vbTextCompare) & ".rtf"
        
        If Status = Connected Then
            On Error Resume Next
                
                If FileExists(LogPath, vbDirectory) = False Then
                    MkDir LogPath
                End If
                
                rtfIn.SaveFile FilePath, rtfRTF
                
            On Error GoTo 0
        End If
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
    'cmdDevSend.Enabled = (LenB(.Text) <> 0)
    If LenB(.Text) Then
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
'If LenB(txtDev.Text) = 0 Then
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

If LenB(txtOut.Text) = 0 Then
    cmdSend.Enabled = False
    cmdSend.Default = False
    
    txtName.Enabled = (modVars.nPrivateChats = 0)
    
    If SendTypeTrue Then
        Msg = eCommands.Typing & "0" & LastName
        If Server Then
            DistributeMsg Msg, -1
        Else
            SendData Msg
        End If
        SendTypeTrue = False
    End If
Else
    cmdSend.Enabled = True
    cmdSend.Default = True
    
    If Len(txtOut.Text) <= 1 Then
        txtName.Enabled = False
    End If
    
    If SendTypeTrue = False Then
        Msg = eCommands.Typing & "1" & LastName
        If Server Then
            DistributeMsg Msg, -1
        Else
            SendData Msg
        End If
        SendTypeTrue = True
    End If
End If

Call SetInactive

End Sub

Public Sub RefreshNetwork(Optional ByRef LB As ScrollListBox = Nothing, _
                          Optional ByRef CommentLB As ScrollListBox = Nothing)
Dim Svr As ListOfServer
Dim i As Integer
Dim S As String
Dim AddComment As Boolean


If LB Is Nothing Then
    Set LB = lstComputers
End If
If Not (CommentLB Is Nothing) Then
    AddComment = True
    CommentLB.Clear
End If


LB.Clear
Me.MousePointer = vbHourglass

S = "Refreshing Server List..."

If Right$(modConsole.ConsoleText, Len(S) + 2) <> (S & vbNewLine) Then
    AddConsoleText S
End If

Svr = EnumServer(SRV_TYPE_ALL)

If Svr.Init Then
    For i = 1 To UBound(Svr.List)
        'If InviteBox = False Then
        LB.AddItem Svr.List(i).ServerName
        
        If AddComment Then
            
            With Svr.List(i)
                
                S = .Comment
                
                If .Type = 6 Then
                    S = S & Space$(3) & "(Vista)"
                ElseIf .Type = 5 Then
                    S = S & Space$(3) & "(XP)"
                End If
                
                CommentLB.AddItem S
                
                    'modVars.TranslateWindowsVer(.VerMajor, .VerMinor, ((.PlatformId And &H80000000) = 0), .Type = 6)
                
            End With
            
        End If
        
        'Else
            'frmInvite.lstComputers.AddItem Svr.List(i).ServerName
        'End If
    Next i
End If

Me.MousePointer = vbNormal

End Sub

'------------DRAWING------------------

Public Sub DoLine(ByVal X As Integer, ByVal Y As Integer)

If NewLine Then
    picDraw.Line (X, Y)-(X, Y), Colour
    NewLine = False
End If

cx = X
cy = Y

SendLine X, Y, picDraw.DrawWidth

End Sub

'Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
''Dim DE As Boolean
'
'If Button = vbLeftButton Then
'
'    Call DoLine(x, Y)
'    'SendLine X, Y, picDraw.DrawWidth
'    'DE = True
'ElseIf Button = vbRightButton Then
'
'    Dim TmpColour As Long, TmpWidth As Integer
'    TmpColour = Colour
'    TmpWidth = picDraw.DrawWidth
'    Colour = picDraw.BackColor
'    picDraw.DrawWidth = RubberWidth
'
'    Call DoLine(x, Y)
'
'    'SendLine X, Y, picDraw.DrawWidth
'
'    Colour = TmpColour
'    picDraw.DrawWidth = TmpWidth
'    'DE = True
'End If
'
''If DE Then DoEvents
'
'End Sub
'Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Dim DE As Boolean
'
'If Button = vbLeftButton Then
'
'    picDraw.Line (cx, cy)-(x, Y), Colour
'
'    cx = x
'    cy = Y
'
'    SendLine cx, cy, picDraw.DrawWidth, x, Y
'
'    'Remember where the mouse is so new lines can be drawn connecting to this point.
'    DE = True
'ElseIf Button = vbRightButton Then
'
'    Dim TmpColour As Long, TmpWidth As Integer
'    TmpColour = Colour
'    TmpWidth = picDraw.DrawWidth
'
'    Colour = picDraw.BackColor
'    picDraw.DrawWidth = RubberWidth
'
'    picDraw.Line (cx, cy)-(x, Y), Colour
'
'    cx = x
'    cy = Y
'
'    SendLine cx, cy, picDraw.DrawWidth, x, Y
'
'    'Remember where the mouse is so new lines can be drawn connecting to this point.
'
'    Colour = TmpColour
'    picDraw.DrawWidth = TmpWidth
'    DE = True
'End If
'
'If DE Then
'    picDraw.Refresh
'    DoEvents
'End If
'
'End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim aX As Integer, aY As Integer

If mnuOptionsMessagingDrawingOff.Checked = False Then
    If Not PickingColour Then
        If Button = vbLeftButton Then
            
            If pDrawDrawnOn = False Then pDrawDrawnOn = True
            
            aX = CInt(X)
            aY = CInt(Y)
            
            Drawing = True
            
            Call DoLine(aX, aY)
            'draw above
            'SendLine X, Y, picDraw.DrawWidth
            
            Call SendDraw
            
            
        ElseIf Button = vbRightButton Then
            
            If DrawingStraight Then
                chkStraightLine.Value = 0
            End If
            
            Dim TmpColour As Long, TmpWidth As Integer
            TmpColour = Colour
            TmpWidth = picDraw.DrawWidth
            Colour = picDraw.BackColor
            picDraw.DrawWidth = RubberWidth
            
            aX = CInt(X)
            aY = CInt(Y)
            
            Call DoLine(aX, aY)
            
            'SendLine X, Y, picDraw.DrawWidth
            
            Colour = TmpColour
            picDraw.DrawWidth = TmpWidth
        End If
    End If
End If

End Sub

Private Sub SendDraw(Optional ByVal bOn As Boolean = True)

Dim Msg As String

If bOn Then
    If SendTrueDraw = False Then
        Msg = eCommands.Drawing & "1" & LastName
        If Server Then
            DistributeMsg Msg, -1
        Else
            SendData Msg
        End If
        SendTrueDraw = True
    End If
Else
    If SendTrueDraw Then
        Msg = eCommands.Drawing & "0" & LastName
        If Server Then
            DistributeMsg Msg, -1
        Else
            SendData Msg
        End If
        SendTrueDraw = False
    End If
End If

End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim DocXcY As Boolean
Dim tcX As Integer, tcY As Integer
Dim aX As Integer, aY As Integer

If mnuOptionsMessagingDrawingOff.Checked = False Then
    If Not PickingColour Then
        If Button = vbLeftButton Then
            
            If pDrawDrawnOn = False Then pDrawDrawnOn = True
            
            aX = CInt(X)
            aY = CInt(Y)
            
            picDraw.Line (cx, cy)-(aX, aY), Colour
            picDraw.Refresh
            'Remember where the mouse is so new lines can be drawn connecting to this point.
            
            tcX = cx
            tcY = cy
            
            cx = aX
            cy = aY
            
            SendLine tcX, tcY, picDraw.DrawWidth, aX, aY
            
        
        ElseIf Button = vbRightButton Then
            
            Dim TmpColour As Long, TmpWidth As Integer
            TmpColour = Colour
            TmpWidth = picDraw.DrawWidth
            
            Colour = picDraw.BackColor
            picDraw.DrawWidth = RubberWidth
            
            aX = CInt(X)
            aY = CInt(Y)
            
            picDraw.Line (cx, cy)-(aX, aY), Colour
            picDraw.Refresh
            
            'docxcy
            tcX = cx
            tcY = cy
            
            cx = aX
            cy = aY
            
            SendLine tcX, tcY, picDraw.DrawWidth, aX, aY
            
            
            'Remember where the mouse is so new lines can be drawn connecting to this point.
            
            
            Colour = TmpColour
            picDraw.DrawWidth = TmpWidth
            
        End If
    ElseIf PickingColour Then
        cx = CInt(X)
        cy = CInt(Y)
    End If
End If

Call SetInactive

End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mnuOptionsMessagingDrawingOff.Checked = False Then
    If Button = vbLeftButton Then
        NewLine = True
        Drawing = False
        Call SendDraw(False)
    End If
End If
End Sub

'pick colour
Private Sub picDraw_Click()
Dim dColour As Long

If mnuOptionsMessagingDrawingOff.Checked = False Then
    
    If PickingColour Then
        dColour = picDraw.Point(cx, cy)
        
        If Colour <> -1 Then
            picColour.BackColor = dColour
            Colour = dColour
            chkPickColour.Value = 0
        End If
        
    ElseIf DrawingStraight Then
        
        If pDrawDrawnOn = False Then pDrawDrawnOn = True
        
        If CBool(StraightPoint1.X) Or CBool(StraightPoint1.Y) Then
            StraightPoint2.X = cx
            StraightPoint2.Y = cy
            picDraw.Line (StraightPoint1.X, StraightPoint1.Y)- _
                (StraightPoint2.X, StraightPoint2.Y), Colour
            
            SendLine StraightPoint1.X, StraightPoint1.Y, picDraw.DrawWidth, _
                StraightPoint2.X, StraightPoint2.Y
            
            StraightPoint1.X = StraightPoint2.X
            StraightPoint1.Y = StraightPoint2.Y
        Else
            StraightPoint1.X = cx
            StraightPoint1.Y = cy
            
            'picDraw.Line (StraightPoint1.x, StraightPoint1.Y)- _
                (StraightPoint1.x, StraightPoint1.Y), Colour
                
            Call DoLine(cx, cy)
            
            
        End If
        
    End If
End If

End Sub

Public Function GetLine() As String
Dim lStart As Long, lEnd As Long
Dim i As Long
Dim Txt As String

lStart = rtfIn.Selstart
Txt = rtfIn.Text

For i = lStart To 1 Step -1
    If Mid$(Txt, i, 2) = vbNewLine Then
        lStart = i
        Exit For
    End If
Next i


For i = rtfIn.Selstart To Len(rtfIn.Text)
    If i Then
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
Dim Msg As String

If KeyAscii = 172 Then
    KeyAscii = 0 'prevent ¬
    
'ElseIf KeyAscii = 8 Then
'    If Len(txtOut.Text) = 1 Then
'        Msg = eCommands.Typing & "0" & txtName.Text
'        If Server Then
'            DistributeMsg Msg, -1
'        Else
'            SendData Msg
'        End If
'    End If
'ElseIf Len(txtOut.Text) = 0 Then 'started typing
'    Msg = eCommands.Typing & "1" & txtName.Text
'    If Server Then
'        DistributeMsg Msg, -1
'    Else
'        SendData Msg
'    End If
End If


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
        txtSendTo.Selstart = Len(txtSendTo.Text)
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

Dim Str As String

Call modSettings.DelSettings

Str = "Settings Deleted"

AddText Str, , True
AddConsoleText Str

mnuFileSaveExit.Checked = False
End Sub

Public Sub TxtName_LostFocus()
Dim Msg As String
Dim NameOp As String, Before As String
'aka Name to be operated on

Before = txtName.Text
NameOp = Trim$(Replace$(Before, "@", vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, "#", vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, modMessaging.MsgEncryptionFlag, vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, ":", vbNullString, , , vbTextCompare)) 'for chat in game
NameOp = Trim$(Replace$(NameOp, modSpaceGame.mPacketSep, vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, modSpaceGame.UpdatePacketSep, vbNullString, , , vbTextCompare))

If Trim$(NameOp) <> Trim$(Before) Then
    AddText "Certain characters can't be used in your name (" & _
        "@, #, " & modMessaging.MsgEncryptionFlag & ", :, " & modSpaceGame.mPacketSep & ", " & modSpaceGame.UpdatePacketSep & _
        ")", TxtError, True
End If

If LenB(NameOp) = 0 Then
    If Me.mnuOptionsMessagingDisplaySysUserName.Checked Then
        NameOp = modVars.GetUserName()
    Else
        NameOp = SckLC.LocalHostName
    End If
End If


If Len(NameOp) > 20 Then
    AddText "Name is Too Long - Truncated", TxtError, True
    NameOp = Left$(NameOp, 20)
End If


NameOp = Trim$(NameOp)

If Status = Connected Then
    
    If NameOp <> LastName Then
        
        
        Msg = eCommands.Info & LastName & " renamed to " & NameOp & "0"
        
        If Server Then
            DistributeMsg Msg, -1
        Else
            SendData Msg
            
            Pause 1
            
            'tell the server to set our name to the new one
            SendData eCommands.SetClientVar & eClientVarCmds.SetName & NameOp
            'don't call the timer - no need to refresh lstConnected etc etc
            
        End If
        
        AddText Mid$(Msg, 3, Len(Msg) - 3), , True
    End If
    
End If

LastName = NameOp
txtName.Text = NameOp 'Trim$(txtName.Text)

End Sub

Private Sub ZipO_Status(Text As String)
AddConsoleText "Zip Status: " & Text
End Sub

Private Sub ZipO_ZipError(Number As eZipError, Description As String)
AddConsoleText "Zip Object Error: " & Description & vbNewLine & Space$(modConsole.IndentLevel + 18) & _
                "Number: " & CStr(Number)
AddText "Zip Error - " & Description
End Sub

Private Sub mnuOptionsWindow2Implode_Click()
Call AnimClick(mnuOptionsWindow2Implode)
End Sub

Private Sub mnuOptionsWindow2Slide_Click()
Call AnimClick(mnuOptionsWindow2Slide)
End Sub

Private Sub mnuOptionsWindow2All_Click()
Call AnimClick(mnuOptionsWindow2All)
End Sub

Private Sub mnuOptionsWindow2NoImplode_Click()
Call AnimClick(mnuOptionsWindow2NoImplode)
End Sub

Private Sub mnuOptionsWindow2Fade_Click()
Call AnimClick(mnuOptionsWindow2Fade)
End Sub

Public Sub AnimClick(Optional ByRef mnu As Menu, Optional ByVal AnimType As eAnimType = -1)

mnuOptionsWindow2Implode.Checked = False
mnuOptionsWindow2Slide.Checked = False
mnuOptionsWindow2Fade.Checked = False
mnuOptionsWindow2All.Checked = False
mnuOptionsWindow2NoImplode.Checked = False

If AnimType = -1 Then
    mnu.Checked = True
Else
    
    Select Case AnimType
        Case eAnimType.aImplode
            mnuOptionsWindow2Implode.Checked = True
        Case eAnimType.aSlide
            mnuOptionsWindow2Slide.Checked = True
        Case eAnimType.aRandom
            mnuOptionsWindow2All.Checked = True
        Case eAnimType.aFade
            mnuOptionsWindow2Fade.Checked = True
        Case eAnimType.None
            mnuOptionsWindow2NoImplode.Checked = True
    End Select
    
End If

End Sub

Public Sub MoveForm(X As Long, Y As Long)

Dim rTo As RECT, rFrom As RECT

With rFrom
    .Top = ScaleY(Me.Top, vbTwips, vbPixels)
    .Left = ScaleX(Me.Left, vbTwips, vbPixels)
    .Bottom = ScaleY(Me.height + .Top, vbTwips, vbPixels)
    .Right = ScaleX(Me.width + .Left, vbTwips, vbPixels)
End With


With rTo
    .Top = ScaleY(Y, vbTwips, vbPixels)
    .Left = ScaleX(X, vbTwips, vbPixels)
    .Bottom = ScaleY(Me.height + .Top, vbTwips, vbPixels)
    .Right = ScaleX(Me.width + .Left, vbTwips, vbPixels)
End With

modImplode.MoveForm Me.hWnd, rTo, rFrom, False

End Sub

Public Sub InviteReceived(ByVal Txt As String)

'frmMain.LastName & "#" & frmMain.SckLC.LocalIP
Dim Name As String, IP As String
Dim Ans As VbMsgBoxResult
Dim AutoReject As Boolean

Name = Left$(Txt, InStr(1, Txt, "#", vbTextCompare) - 1)
IP = Mid$(Txt, InStr(1, Txt, "#", vbTextCompare) + 1)
AutoReject = mnuOptionsMessagingDisplayIgnoreInvites.Checked

'confirm receiving of invite
frmUDP.SendToSingle IP, frmUDP.UDPInfo & Me.LastName & " Received the Invite", False
frmUDP.UDPListen


If Me.Visible = False Then
    Me.ShowForm
End If

If Questioning Or AutoReject Then
    Ans = vbNo
    
    Txt = "Invite Ignored, from " & Name & " (" & IP & ")"
    
    If AutoReject Then
        AddConsoleText Txt
        AddText Txt, , True
    End If
Else
    Ans = Question(Name & " (" & IP & ") sends an invite, connect?", frmUDP.cmdInvite)
End If

If Ans = vbYes Then
    Me.CleanUp True
    Connect IP
    Inviter = Name
Else 'If Ans = vbNo Then
    frmUDP.SendToSingle IP, frmUDP.UDPInfo & "Invite to " & Me.LastName & _
        Space$(1) & IIf(AutoReject And Not Questioning, "Auto-", vbNullString) & "Rejected" & _
        IIf(Questioning And Not AutoReject, " (Answering another Question)", vbNullString), False
    
    frmUDP.UDPListen
    Inviter = vbNullString
End If

End Sub

Public Sub SetIcon(ByVal St As eStatus)
Dim i As Integer

i = CInt(St) + 1

If bDevMode Then
    i = i + 8
ElseIf Me.mnuFileGameMode.Checked Or modSpaceGame.GameFormLoaded Or modStickGame.StickFormLoaded Then
    i = i + 4
End If

pSetIcon i
'tmrMain.Enabled = (Status = Connected) And mnuPopupAnim.Checked

End Sub

Private Sub pSetIcon(ByVal i As Integer)

frmSystray.IconHandle = frmMain.imglstIcons.ListImages(i).Picture
frmMain.Icon = frmMain.imglstIcons.ListImages(i).Picture
frmMain.imgStatus.Picture = frmMain.imglstIcons.ListImages(i).Picture

End Sub

Public Sub RefreshIcon()

Call SetIcon(modVars.Status)

End Sub

'Public Sub DoPing(ByVal Address As String, Optional ByVal addConsole As Boolean = False)
'Dim Png As Long
'
'If Status = Connected Or addConsole Then
'
'    If mnuOptionsAdvPing.Checked Then
'        Address = modWinsock.GetIPFromHostName(Address)
'
'        DoEvents
'
'        Png = modWinsock.Ping(Address) 'address needs resolving
'
'        If addConsole Then
'            AddConsoleText "Pinged " & Address & " Time: " & CStr(Png)
'        End If
'
'        If addConsole = False Then
'            sbMain.Panels(2).Text = "Ping: " & Png & "ms"
'        End If
'    End If
'
'End If
'
'End Sub
