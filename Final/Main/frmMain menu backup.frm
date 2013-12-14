VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Communicator"
   ClientHeight    =   9285
   ClientLeft      =   75
   ClientTop       =   765
   ClientWidth     =   9945
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   9945
   Begin MSWinsockLib.Winsock SckLC 
      Left            =   4680
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
   Begin VB.TextBox txtOut 
      Height          =   765
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   41
      Top             =   3000
      Width           =   4575
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "[Status]"
      Top             =   360
      Width           =   1935
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   3120
      ScaleHeight     =   1905
      ScaleWidth      =   3585
      TabIndex        =   50
      Top             =   5280
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Image imgClientDP 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1485
         Left            =   60
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.CommandButton cmdSlash 
      Caption         =   "/"
      Height          =   375
      Left            =   9000
      TabIndex        =   45
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox picBig 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   5280
      ScaleHeight     =   2985
      ScaleWidth      =   3105
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Image imgBig 
         Height          =   285
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   285
      End
   End
   Begin VB.Timer tmrHost 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   600
      Top             =   720
   End
   Begin VB.Timer tmrInfoHide 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6000
      Top             =   720
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3480
      ScaleHeight     =   225
      ScaleWidth      =   5025
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   5055
   End
   Begin projMulti.ucFileTransfer ucFileTransfer 
      Left            =   8160
      Top             =   4200
      _extentx        =   661
      _extenty        =   661
   End
   Begin VB.Timer tmrMain 
      Interval        =   10000
      Left            =   2880
      Top             =   1440
   End
   Begin VB.ListBox lstComputers 
      Height          =   840
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ListBox lstConnected 
      Height          =   840
      Left            =   1680
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame fraTyping 
      Height          =   615
      Left            =   120
      TabIndex        =   48
      Top             =   4920
      Width           =   3195
      Begin VB.Label lblTyping 
         Alignment       =   2  'Center
         Caption         =   "Drawing Label"
         Height          =   435
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrivate 
      Caption         =   "Private Chat"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   12
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
      TabIndex        =   15
      Top             =   3000
      Width           =   3195
      Begin VB.PictureBox picColours 
         BackColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00004080&
         Height          =   255
         Index           =   9
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H0000C0C0&
         Height          =   255
         Index           =   10
         Left            =   600
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00008000&
         Height          =   255
         Index           =   11
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00FF0000&
         Height          =   255
         Index           =   12
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00808080&
         Height          =   255
         Index           =   13
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   1560
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   1920
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkPickColour 
         Caption         =   "Pick Colour Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   650
         Width           =   1695
      End
      Begin VB.CheckBox chkStraightLine 
         Caption         =   "Straight Line Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   890
         Width           =   1695
      End
      Begin VB.PictureBox picClearBoard 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   1455
         TabIndex        =   20
         Top             =   600
         Width           =   1455
         Begin VB.CommandButton cmdCls 
            Caption         =   "Clear Board"
            Height          =   375
            Left            =   360
            TabIndex        =   21
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
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.ComboBox cboWidth 
         Height          =   315
         Left            =   600
         TabIndex        =   16
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cboRubber 
         Height          =   315
         Left            =   2160
         TabIndex        =   18
         Text            =   "5"
         Top             =   240
         Width           =   615
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   600
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H0000FF00&
         Height          =   255
         Index           =   3
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Index           =   4
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00C000C0&
         Height          =   255
         Index           =   5
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   1560
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblColour 
         Caption         =   "Colour   Last Colour"
         Height          =   405
         Left            =   2280
         TabIndex        =   32
         Top             =   1240
         Width           =   855
      End
      Begin VB.Label lblDraw 
         Caption         =   "Draw:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblRubber 
         Caption         =   "Rubber:"
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSmile 
      Height          =   375
      Left            =   9000
      Picture         =   "frmMain.frx":636A
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3000
      Width           =   285
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   52
      Top             =   9030
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "IP Panel"
            TextSave        =   "IP Panel"
            Object.ToolTipText     =   "IP Information (Right Click Here)"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   11880
            Text            =   "General Info"
            TextSave        =   "General Info"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Version etc"
            TextSave        =   "Version etc"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      Height          =   2895
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   2835
      ScaleWidth      =   7155
      TabIndex        =   51
      Top             =   5640
      Width           =   7215
   End
   Begin VB.Timer tmrLog 
      Interval        =   10000
      Left            =   5880
      Top             =   1920
   End
   Begin VB.Timer tmrShake 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8280
      Top             =   3720
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "&No"
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   46
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "&Yes"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   47
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdShake 
      Caption         =   "Shake"
      Height          =   375
      Left            =   8160
      TabIndex        =   44
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   600
      MaxLength       =   20
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Host"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   8160
      TabIndex        =   42
      Top             =   3000
      Width           =   735
   End
   Begin projMulti.smRtfFBox rtfIn 
      Height          =   2400
      Left            =   3360
      TabIndex        =   14
      Top             =   600
      Width           =   6000
      _extentx        =   10583
      _extenty        =   4233
      font            =   "frmMain.frx":671C
      hwnd            =   295113
      mouseicon       =   "frmMain.frx":6748
      text            =   "rtfIn"
      enabletextfilter=   -1
      selrtf          =   $"frmMain.frx":6766
   End
   Begin VB.Label lblBorder 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgDP 
      Appearance      =   0  'Flat
      Height          =   600
      Index           =   0
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
   Begin VB.Image imgStatus 
      Height          =   690
      Left            =   2640
      Top             =   45
      Width           =   795
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   50
      TabIndex        =   1
      Top             =   120
      Width           =   495
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
      Begin VB.Menu mnuFileMini 
         Caption         =   "Show Mini Window"
         Shortcut        =   ^M
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
         Begin VB.Menu mnuFileSettingsUserProfileSave 
            Caption         =   "Save Settings"
         End
         Begin VB.Menu mnuFileSettingsUserProfileLoad 
            Caption         =   "Load Settings"
         End
         Begin VB.Menu mnuFileSettingsExport 
            Caption         =   "Export Settings..."
         End
         Begin VB.Menu mnuFileSettingsImport 
            Caption         =   "Import Settings..."
         End
         Begin VB.Menu mnuFileSettingsUserProfile 
            Caption         =   "User Profile Settings"
            Begin VB.Menu mnuFileSettingsUserProfileExportOnExit 
               Caption         =   "Save Settings On Exit"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuFileSettingsUserProfileDelete 
               Caption         =   "Delete UserProfile Settings"
            End
            Begin VB.Menu mnuFileSettingsUserProfileView 
               Caption         =   "View UserProfile Settings..."
            End
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
            Caption         =   "Detailed Network List..."
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
         Caption         =   "Voice Options..."
      End
      Begin VB.Menu mnuOptionsSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsDP 
         Caption         =   "Display Pictures"
         Begin VB.Menu mnuOptionsDPSet 
            Caption         =   "Set Display Picture..."
         End
         Begin VB.Menu mnuOptionsDPReset 
            Caption         =   "Reset All Pictures"
         End
         Begin VB.Menu mnuOptionsDPRefresh 
            Caption         =   "Refresh Pictures"
         End
         Begin VB.Menu mnuOptionsDPSaveAll 
            Caption         =   "Save All Pictures (Received Files)"
         End
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
               Visible         =   0   'False
            End
            Begin VB.Menu mnuOptionsWindow2All 
               Caption         =   "All Methods"
            End
            Begin VB.Menu mnuOptionsWindow2AnimationSep 
               Caption         =   "-"
            End
            Begin VB.Menu mnuOptionsWindow2NoImplode 
               Caption         =   "Don't Animate"
               Checked         =   -1  'True
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
               Caption         =   "Lobby Window..."
               Shortcut        =   ^G
            End
            Begin VB.Menu mnuOptionsMessagingWindowsWebcam 
               Caption         =   "Webcam Window..."
               Shortcut        =   ^P
               Visible         =   0   'False
            End
            Begin VB.Menu mnuOptionsMessagingWindowsFT 
               Caption         =   "Manual File Transfer..."
               Shortcut        =   ^H
            End
         End
         Begin VB.Menu mnuOptionsMessagingDisplay 
            Caption         =   "Display"
            Begin VB.Menu mnuOptionsMessagingDisplayCompact 
               Caption         =   "Compact Typing Box"
            End
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
            Begin VB.Menu mnuOptionsMessagingDisplaySep1 
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
            Begin VB.Menu mnuOptionsMessagingDisplaySmilies 
               Caption         =   "Smilies"
               Begin VB.Menu mnuOptionsMessagingDisplaySmiliesEnable 
                  Caption         =   "Enable Smilies"
                  Checked         =   -1  'True
               End
               Begin VB.Menu mnuOptionsMessagingDisplaySmiliesOld 
                  Caption         =   "Use MSN Smilies"
               End
            End
         End
         Begin VB.Menu mnuOptionsMessagingLogging 
            Caption         =   "Logging"
            Begin VB.Menu mnuOptionsMessagingLoggingConv 
               Caption         =   "Log Conversations"
            End
            Begin VB.Menu mnuOptionsMessagingLoggingAutoSave 
               Caption         =   "Auto Save Current Conversation"
            End
            Begin VB.Menu mnuOptionsMessagingLoggingPrivate 
               Caption         =   "Log Private Chats"
            End
            Begin VB.Menu mnuOptionsMessagingLoggingDrawing 
               Caption         =   "Auto-Save Drawing"
            End
         End
         Begin VB.Menu mnuOptionsMessagingCharMap 
            Caption         =   "Character Map..."
         End
         Begin VB.Menu mnuOptionsMessagingSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsMessagingServerMsg 
            Caption         =   "Set Server Message..."
         End
         Begin VB.Menu mnuOptionsMatrix 
            Caption         =   "Matrix Chat Mode"
         End
         Begin VB.Menu mnuOptionsMessagingEncrypt 
            Caption         =   "Encrypt Sent Messages"
         End
         Begin VB.Menu mnuOptionsMessagingClearTypeList 
            Caption         =   "Clear Typing/Drawing List"
         End
         Begin VB.Menu mnuOptionsMessagingSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsMessagingReplaceQ 
            Caption         =   "Replace / with ?"
         End
         Begin VB.Menu mnuOptionsMessagingIgnoreMatrix 
            Caption         =   "Ignore Matrix Messages"
         End
         Begin VB.Menu mnuOptionsMessagingDrawingOff 
            Caption         =   "Turn Off Drawing"
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
         Begin VB.Menu mnuOptionsAdvDisplay 
            Caption         =   "Display"
            Begin VB.Menu mnuOptionsAdvDisplayStyles 
               Caption         =   "Enable Visual Styles"
            End
            Begin VB.Menu mnuOptionsAdvDisplayGlassBG 
               Caption         =   "Enable Glass Border"
            End
            Begin VB.Menu mnuOptionsAdvDisplayVistaControls 
               Caption         =   "Enable Vista Controls"
            End
         End
         Begin VB.Menu mnuOptionsAdvSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsAdvNoStandby 
            Caption         =   "Prevent Standby/Hibernation"
         End
         Begin VB.Menu mnuOptionsAdvNoStandbyConnected 
            Caption         =   "Prevent Standby/Hibernation When Connected"
            Checked         =   -1  'True
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
         Begin VB.Menu mnuOptionsAdvAutoUpdate 
            Caption         =   "Remind me to check for updates every 5 days"
            Checked         =   -1  'True
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
      Begin VB.Menu mnuOptionsSocket 
         Caption         =   "Socket"
         Begin VB.Menu mnuOptionsSocketHost 
            Caption         =   "Host"
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
         Caption         =   "Communicator Website..."
      End
      Begin VB.Menu mnuOnlineLogin 
         Caption         =   "Login/Stats..."
      End
      Begin VB.Menu mnuonlinesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnlineManual 
         Caption         =   "HTTP Download"
      End
      Begin VB.Menu mnuOnlineIPs 
         Caption         =   "View IPs..."
      End
      Begin VB.Menu mnuOnlinePortForwarding 
         Caption         =   "Port Forwarding..."
      End
      Begin VB.Menu mnuonlinesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnlineFileTransfer 
         Caption         =   "File Transfer..."
      End
      Begin VB.Menu mnuOnlineMessages 
         Caption         =   "Messages..."
      End
   End
   Begin VB.Menu mnuDev 
      Caption         =   "DevMode"
      Begin VB.Menu mnuDevForms 
         Caption         =   "Windows"
         Begin VB.Menu mnuDevFormsCmds 
            Caption         =   "Dev Command Window..."
            Shortcut        =   ^Y
         End
         Begin VB.Menu mnuDevForm 
            Caption         =   "Dev Window (Received Data)..."
         End
         Begin VB.Menu mnuDevDataForm 
            Caption         =   "Simulation/Advanced Window..."
         End
         Begin VB.Menu mnuDevFormsClients 
            Caption         =   "Client List..."
         End
      End
      Begin VB.Menu mnuDevDataCmds 
         Caption         =   "Data Commands"
         Begin VB.Menu mnuDevDataCmdsTypeShow 
            Caption         =   "Show When I Type"
         End
         Begin VB.Menu mnuDevDataCmdsSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDevDataCmdsRecSent 
            Caption         =   "Show Received/Sent Data"
         End
         Begin VB.Menu mnuDevShowCmds 
            Caption         =   "Show Recieved Dev Commands"
         End
         Begin VB.Menu mnuDevDataCmdsSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDevDataCmdsBlock 
            Caption         =   "Block Remote Commands..."
         End
         Begin VB.Menu mnuDevDataCmdsSetBlockMessage 
            Caption         =   "Set Block Message..."
         End
         Begin VB.Menu mnuDevDataCmdsSpecial 
            Caption         =   "Special"
            Begin VB.Menu mnuDevDataCmdsOverride 
               Caption         =   "Override Block"
            End
            Begin VB.Menu mnuDevDataCmdsSpecialOff 
               Caption         =   "Uber Dev Mode"
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
            Caption         =   "Debug Mode (Add Console to Text File)"
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
         Begin VB.Menu mnuDevAdvCmdsRemZLib 
            Caption         =   "Remove ZLib Dll"
         End
      End
      Begin VB.Menu mnuDevMaintenance 
         Caption         =   "Maintenance"
         Begin VB.Menu mnuDevMaintenanceTimers 
            Caption         =   "Disable Timers"
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
         Caption         =   "Help..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About..."
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuHelpBug 
         Caption         =   "Bug Report..."
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
         Caption         =   "Single Command..."
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuConsoleTypeLots 
         Caption         =   "Mutiple Commands..."
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
   Begin VB.Menu mnuFont 
      Caption         =   "Font"
      Begin VB.Menu mnuFontColour 
         Caption         =   "Colour..."
      End
      Begin VB.Menu mnuFontDialog 
         Caption         =   "Font..."
      End
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Set Status"
      Begin VB.Menu mnuStatusAway 
         Caption         =   "AFK"
      End
      Begin VB.Menu mnuStatusResetName 
         Caption         =   "Reset to User Name"
      End
   End
   Begin VB.Menu mnuCommands 
      Caption         =   "Slash Commands"
      Begin VB.Menu mnuCommandsMe 
         Caption         =   "Insert /Me"
      End
      Begin VB.Menu mnuCommandsDescribe 
         Caption         =   "Insert /Describe"
      End
      Begin VB.Menu mnuCommandsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommandsSpeechEmph 
         Caption         =   "Speech Emphasis Tags"
      End
      Begin VB.Menu mnuCommandsJeffery 
         Caption         =   "Blocked Nose Tags"
      End
      Begin VB.Menu mnuCommandsGregory 
         Caption         =   "Lancashire Tags"
      End
   End
   Begin VB.Menu mnuDP 
      Caption         =   "Display Pictures"
      Begin VB.Menu mnuDPView 
         Caption         =   "View Picture..."
      End
      Begin VB.Menu mnuDPOpen 
         Caption         =   "Open Picture's Folder..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DefaultPasswordMaxLen = 50
Private ServerMsg As String

Private Const TextBoxHeight = 285, TextBoxHeightInc = 3
Private Const NewLineLimit = 8

'private chat
Private iSelectedClientSock As Integer

'item in listbox
Private Const LB_GETITEMHEIGHT = &H1A1
Private Const LB_GETITEMRECT = &H198
Private Const LB_GETTOPINDEX = &H18E
Private Const LB_FINDSTRING = &H18F

Private Const LB_SETTOPINDEX = &H197
'Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


''menu popup
'Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Const ksAway = " (AFK)"
Private Const ListeningStr = "Awaiting Connection..."

'font stuff
Private prtfFontName As String
Private prtfFontSize As Single
Private prtfBold As Boolean, prtfItalic As Boolean

'setting main icon (taskbar + alt tab)
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4
Private Const WM_SETICON = &H80
Private Const ICON_BIG = 1
Private Const ICON_SMALL = 0


'logging
Private CurrentLogFile As String
Private Const err_INVALIDORNOACCESS = 75

'DP stuff
Public DP_Path As String
Private iCurrentMouseOver As Integer
Private Const DP_Error_Socket = "Hang on! Data is needed from the server. Try again in thirty seconds"
Private Const DP_Error_Clients = "Hold your horses! Can't Set Display Picture. Data needed from the server"
                                '"Whoa - Can't Set Image Yet - Clients Not Initialised. Slow down"
Private Const DP_Error_Server = "Hold your horses! Wait until the client list is set up"

''icon stuff
'Private Declare Function StretchBlt Lib "gdi32" ( _
    ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
    ByVal dwRop As Long) As Long


'menu height
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYCAPTION = 4
Private Const SM_CYFIXEDFRAME = 8

Private Const IconRefreshDelay As Long = 60000 * 3 '3mins

'pick colour (faster)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Const picDrawBackColour = &H8000000F

'dragging
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

'vista colour stuff
Private Const Default_Backcolour = &H8000000F

'right click menu
Private Const RightClickExt = "*", RightClickMenuTitle = "Open Communicator"

Private Inviter As String 'who invited us? - add to convo on connect

'for the IP status bar
Private Const MinWidth As Integer = 10500


Private Const QuestionTimeOut As Long = 60000

'for zipping
Private WithEvents ZipO As clsZipExtraction
Attribute ZipO.VB_VarHelpID = -1

Public pDrawDrawnOn As Boolean 'has picDraw been drawn on?

Private pHasFocus As Boolean

Private QuestionReply As Byte
Private Const Shake_Delay = 5000
Private LastShake As Long
'Private CanShake As Boolean
'Private LastWndState As FormWindowStateConstants
Private Questioning As Boolean

Private InActiveTmr As Integer

Private Const MsMessageDelay As Long = 500 '1000

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


'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'Private Const WM_LBUTTONDBLCLK = &H203

Public LastName As String, LastStatus As String

Private Const LIPHeading As String = "Internal IP: "
Private Const RIPHeading As String = "External IP: "

Private Type PointInt
    X As Integer
    Y As Integer
End Type

Private StraightPoint1 As PointInt
Private StraightPoint2 As PointInt

Private iInfoTimer As Integer
Private sInfoText As String
Private bInfoCanMove As Boolean
'################################################################################################
Public Sub SetInfo(ByVal sTxt As String)

On Error Resume Next
'lblInfo.Caption = sTxt
picInfo.Visible = True
sInfoText = sTxt

iInfoTimer = 0

'picInfo.Left = 3480 - don't need to, plus it messes up with vista border
picInfo.height = 255
picInfo.width = Me.ScaleWidth - picInfo.Left - 100

tmrInfoHide.Enabled = True
End Sub

Private Sub cmdSlash_Click()
PopupMenu mnuCommands
End Sub

Private Sub cmdSmile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If rtfIn.SmileyBoxVisible Then
    'cancel, allow it to lose focus
    ReleaseCapture
End If
End Sub

Private Sub mnuCommandsDescribe_Click()
Const kAction = "Run for the hills!"
Dim i As Integer

txtOut.SelText = "/Describe " & kAction

i = InStr(1, txtOut.Text, kAction)

txtOut.Selstart = i - 1
txtOut.Sellength = Len(kAction)

On Error Resume Next
txtOut.SetFocus
End Sub

Private Sub mnuCommandsGregory_Click()
AddTags UCase$(modMessaging.GregTag) '"LANCS"
End Sub

Private Sub mnuCommandsJeffery_Click()
AddTags UCase$(modMessaging.JeffTag) '"BLOCKEDNOSE"
End Sub

Private Sub mnuCommandsMe_Click()
Const kAction = "likes cake"
Dim i As Integer

txtOut.SelText = "/Me " & kAction

i = InStr(1, txtOut.Text, kAction)

txtOut.Selstart = i - 1
txtOut.Sellength = Len(kAction)

On Error Resume Next
txtOut.SetFocus
End Sub

Private Sub mnuCommandsSpeechEmph_Click()
AddTags "EMPH"
End Sub

Private Sub AddTags(ByVal sTag As String)

Dim iSelEnd As Integer

sTag = LCase$(sTag)

With txtOut
    iSelEnd = .Selstart + .Sellength
    .Sellength = 0
    .SelText = "<" & sTag & ">"
    
    .Selstart = iSelEnd + Len("<" & sTag & ">")
    .SelText = "</" & sTag & ">"
    
    .Selstart = iSelEnd + 2 + Len(sTag)
    
    
    
    On Error Resume Next
    .SetFocus
End With

End Sub

Private Sub mnuHelpBug_Click()
Unload frmBug
Load frmBug
frmBug.Show vbModeless, Me
End Sub

Private Sub mnuFileMini_Click()

mnuFileMini.Checked = Not mnuFileMini.Checked

Unload frmMini
If mnuFileMini.Checked Then
    Load frmMini
End If

End Sub

Private Sub mnuOptionsDPSaveAll_Click()
mnuOptionsDPSaveAll.Checked = Not mnuOptionsDPSaveAll.Checked
End Sub

'Private Sub mnuDevAdvCmdsIgnoreSD_Click()
'mnuDevAdvCmdsIgnoreSD.Checked = Not mnuDevAdvCmdsIgnoreSD.Checked
'End Sub

Private Sub mnuOptionsMessagingCharMap_Click()
Unload frmCharMap
Load frmCharMap
frmCharMap.Show vbModeless, Me
End Sub

Public Sub mnuOptionsMessagingDisplayCompact_Click()
mnuOptionsMessagingDisplayCompact.Checked = Not mnuOptionsMessagingDisplayCompact.Checked

txtOut.height = IIf(mnuOptionsMessagingDisplayCompact.Checked, TextBoxHeight, TextBoxHeightInc * TextBoxHeight)

Form_Resize
End Sub

Private Sub mnuOptionsMessagingLoggingDrawing_Click()
mnuOptionsMessagingLoggingDrawing.Checked = Not mnuOptionsMessagingLoggingDrawing.Checked
End Sub

Private Sub mnuOptionsMessagingLoggingPrivate_Click()
mnuOptionsMessagingLoggingPrivate.Checked = Not mnuOptionsMessagingLoggingPrivate.Checked
End Sub

Private Sub mnuOptionsMessagingServerMsg_Click()
Dim Msg As String

Msg = modVars.Password("Enter a server message", Me, "Server Message", "Hello there", False, DefaultPasswordMaxLen)

If LenB(Msg) Then
    ServerMsg = Msg
End If

End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bInfoCanMove = False
End Sub

Private Sub tmrInfoHide_Timer()
Const kWidth = 40

picInfo.Cls
picInfo.CurrentY = 0
picInfo.CurrentX = picInfo.width / 2 - picInfo.TextWidth(sInfoText) / 2
picInfo.Print sInfoText

'ValBelow = PauseDelay/Interval
If iInfoTimer > 60 Then
    
    'On Error Resume Next
    'picInfo.Left = picInfo.Left + kWidth
    
    If bInfoCanMove Then
        If picInfo.width > kWidth Then
            picInfo.width = picInfo.width - kWidth
            
            'reduce height
            If picInfo.width <= 1160 Then
                On Error Resume Next 'no error is raised...?
                picInfo.height = picInfo.height - 10
            End If
            
        Else
            picInfo.Visible = False
            tmrInfoHide.Enabled = False
        End If
    End If
Else
    iInfoTimer = iInfoTimer + 1
End If
    
End Sub
'################################################################################################

Private Sub mnuDevDataCmdsTypeShow_Click()
mnuDevDataCmdsTypeShow.Checked = Not mnuDevDataCmdsTypeShow.Checked
End Sub

Public Property Get FT_Path() As String
Dim P As String

P = AppPath() & "Received Files"

If FileExists(P, vbDirectory) = False Then
    On Error Resume Next
    MkDir P
End If

FT_Path = P

End Property

Public Property Get rtfFontSize() As Single
rtfFontSize = prtfFontSize
End Property

Public Property Let rtfFontSize(sF As Single)
prtfFontSize = sF
'rtfIn.Font.Size = sF
End Property

Public Property Get rtfFontName() As String
rtfFontName = prtfFontName
End Property

Public Property Let rtfFontName(sF As String)
prtfFontName = sF
txtOut.Font = sF
End Property

Public Property Get rtfItalic() As Boolean
rtfItalic = prtfItalic
End Property

Public Property Let rtfItalic(bI As Boolean)
prtfItalic = bI
txtOut.Font.Italic = bI
'rtfIn.Font.Italic = bI
End Property

Public Property Get rtfBold() As Boolean
rtfBold = prtfBold
End Property

Public Property Let rtfBold(bB As Boolean)
prtfBold = bB
txtOut.Font.Bold = bB
'rtfIn.Font.Bold = bB
End Property

Private Sub cmdAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    mnuFileManual_Click
End If
End Sub

Private Sub cmdShake_KeyPress(KeyAscii As Integer)
On Error Resume Next
txtOut.SetFocus
End Sub

Private Sub imgDP_DblClick(Index As Integer)
mnuDPView_Click
End Sub

Private Sub imgDP_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'show the bigger picture

If imgDP(Index).Picture.Handle <> 0 Then
    iCurrentMouseOver = Index
    
    'bmainNavigate "about:<html><body scroll='no'><img src='" & _
        Path & _
        "'></img></body></html>"
    
    'picBig.width = imgBig.width
    'picBig.height = imgBig.height
    
    ShowBigDP Index
End If

End Sub

Private Sub ShowBigDP(Index As Integer)
Const MinLeft = 1000

If imgBig.Tag <> CStr(Index) Then
    
    imgBig.Stretch = False
    Set imgBig.Picture = imgDP(Index).Picture
    
    
    'assume can do
    picBig.Left = imgDP(Index).Left + imgDP(Index).width / 2 - imgBig.width / 2
    picBig.width = imgBig.width
    picBig.height = imgBig.height
    
    
    If picBig.Left < MinLeft Then picBig.Left = MinLeft
    
    
    If (picBig.Left + picBig.width) > (Me.width - 20) Then
        'no can do
        imgBig.Stretch = True
        
        ResetpicBigXY
        
        picBig.Left = imgDP(Index).Left + imgDP(Index).width - picBig.width / 2
    End If
    
    
    picBig.Visible = True
    
    imgBig.Tag = CStr(Index)
End If

End Sub

Private Sub ResetpicBigXY()

picBig.width = 6300 '4500
picBig.height = 5460 '3900
imgBig.width = picBig.width
imgBig.height = picBig.height

End Sub

Private Sub imgBig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDP_MouseDown iCurrentMouseOver, Button, Shift, X, Y
End Sub

Private Sub imgBig_DblClick()
imgDP_DblClick iCurrentMouseOver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
NewLine = True
Call SetInactive
SetInfoPanel "Communicator Main Window"

HideExtras
bInfoCanMove = True
End Sub

Private Sub HidePicBig()
picBig.Visible = False
imgBig.Tag = vbNullString
End Sub

Private Sub HideExtras(Optional ByVal bStatusAndName As Boolean = True)
HidePicClient
HidePicBig

If bStatusAndName Then
    txtStatus.Visible = False
    txtName_LostFocus
End If

End Sub

Private Sub mnuDevDataCmdsSetBlockMessage_Click()
Dim sTxt As String

sTxt = modVars.Password("Enter a Custom Block Message...", _
    Me, "Dev Block Message", modMessaging.DevBlockedMessage, False, DefaultPasswordMaxLen)


If LenB(sTxt) Then
    modMessaging.DevBlockedMessage = sTxt
End If

End Sub

Private Sub mnuDevFormsCmds_Click()
Unload frmDevCmd
Load frmDevCmd
frmDevCmd.Show vbModeless, Me
End Sub

Private Sub mnuDPOpen_Click()

modVars.OpenFolder vbNormalFocus, modDP.DP_Dir_Path

End Sub

Private Sub mnuDPView_Click()
Dim Path As String

Path = modDP.DP_Dir_Path
If Right$(Path, 1) <> "\" Then Path = Path & "\"

If iCurrentMouseOver = modMessaging.MySocket Then
    Path = Path & "local.jpg"
    
ElseIf iCurrentMouseOver = 0 Then
    Path = Path & "-1.jpg"
    
Else
    Path = Path & CStr(iCurrentMouseOver) & ".jpg"
    
End If

modVars.OpenImage Path

End Sub

Private Sub mnuFileSettingsUserProfileDelete_Click()
Dim sFile As String

mnuFileSettingsUserProfileExportOnExit.Checked = False

'ClearFolder modSettings.GetUserSettingsPath()

sFile = modSettings.GetSettingsFile()

If FileExists(sFile) Then
    On Error GoTo EH
    Kill sFile
End If

AddText "Deleted User Profile Settings", , True

Exit Sub
EH:
AddText "Error - " & Err.Description, TxtError, True
End Sub

Private Sub mnuFileSettingsUserProfileExportOnExit_Click()
mnuFileSettingsUserProfileExportOnExit.Checked = Not mnuFileSettingsUserProfileExportOnExit.Checked
End Sub

Private Sub mnuFileSettingsUserProfileSave_Click()
SaveUserProfileSettings
'AddText "Settings Saved", , True
SetInfo "Settings Saved"
End Sub

Private Sub mnuFileSettingsUserProfileLoad_Click()
LoadUserProfileSettings

SetInfo "Loaded Settings"
End Sub

Private Sub SaveUserProfileSettings()
modSettings.ExportSettings modSettings.GetSettingsFile(), False
End Sub

Private Sub LoadUserProfileSettings()
modSettings.ImportSettings modSettings.GetSettingsFile(), False
End Sub

'Private Sub mnuFileSaveExit_Click()
'mnuFileSaveExit.Checked = Not mnuFileSaveExit.Checked
'End Sub
'
'Private Sub mnuFileSaveSettings_Click()
'modSettings.SaveSettings
'AddText "Settings Saved", , True
'End Sub

Private Sub mnuFileSettingsUserProfileView_Click()
modVars.OpenFolder vbNormalFocus, modSettings.GetUserSettingsPath()
End Sub

Private Sub mnuFontColour_Click()
txtOut_DblClick
End Sub

Private Sub mnuFontDialog_Click()

With Me.Cmdlg
    .FontName = rtfFontName
    .FontSize = rtfFontSize
    .FontItalic = rtfItalic
    .FontBold = rtfBold
    
    
    .Flags = cdlCFForceFontExist Or cdlCFLimitSize Or cdlCFBoth 'Or cdlCFScalableOnly
    .Max = MaxFont
    .Min = MinFont
    
    On Error GoTo EH
    .ShowFont
    
    rtfFontName = .FontName
    rtfFontSize = .FontSize
    rtfItalic = .FontItalic
    rtfBold = .FontBold
End With

EH:
End Sub

Private Sub mnuOptionsAdvAutoUpdate_Click()
mnuOptionsAdvAutoUpdate.Checked = Not mnuOptionsAdvAutoUpdate.Checked
End Sub

'Private Sub mnuOptionsDPDefault_Click()
'Dim Path As String
'Dim i As Integer, iClient As Integer
'
'On Error GoTo EH
'If Not Server Then
'    If modMessaging.MySocket = 0 Then
'        AddText DP_Error_Socket, TxtError, True
'        'AddText "Wait about 12 seconds, then try again", TxtError, True
'        Exit Sub
'    End If
'End If
'
'
'Path = modDP.DP_Dir_Path
'If FileExists(Path, vbDirectory) = False Then
'    On Error Resume Next
'    MkDir Path
'End If
'Path = Path & "\Local.jpg"
'If FileExists(Path) Then
'    '################################################################
'    'set it
'    iClient = -1
'    For i = 0 To UBound(Clients)
'        If Clients(i).iSocket = modMessaging.MySocket Then
'            iClient = i
'            Exit For
'        End If
'    Next i
'
'    If iClient > -1 Then
'        Set Clients(iClient).IPicture = LoadPicture(Path)
'        ShowDP iClient
'        modDP.My_DP_Path = Path
'
'        If Server Then
'            For i = 0 To UBound(Clients)
'                Clients(i).bSentHostDP = False
'            Next i
'        Else
'            modDP.bSentMyPicture = False
'        End If
'
'        i = InStrRev(Path, "\", , vbTextCompare)
'        AddText "Loaded Picture (" & Right$(Path, Len(Path) - i) & ")", , True
'    Else
'        AddText DP_Error_Clients, TxtError, True
'    End If
'    '#######################################################################
'Else
'    AddText "The file doesn't exist - You'll have to go the long way", TxtError, True
'End If
'
'
'EH:
'End Sub

'Private Sub mnuOptionsDPEnable_Click()
'mnuOptionsDPEnable.Checked = Not mnuOptionsDPEnable.Checked
'mnuOptionsDPSet.Enabled = mnuOptionsDPEnable.Checked
'End Sub

Private Sub mnuOptionsDPRefresh_Click()
Dim i As Integer


For i = 0 To UBound(Clients)
    With Clients(i)
        .bDPSet = False
        .bSentHostDP = False
        .sHasiDPs = vbNullString
        Set .IPicture = Nothing
    End With
Next i



If Not Server Then
    'yo server, we don't have a DP
    SendData eCommands.SetClientVar & eClientVarCmds.SetDPSet & "0"
End If


ResetImgDP 0
For i = 1 To imgDP.UBound
    Unload imgDP(i)
Next i

AddText "Pictures will be updated soon", , True

End Sub

Private Sub mnuOptionsDPReset_Click()
Dim i As Integer


For i = 0 To UBound(Clients)
    With Clients(i)
        .bDPSet = False
        .bSentHostDP = False
        .sHasiDPs = vbNullString
        Set .IPicture = Nothing
    End With
Next i


'imgDP(0).BorderStyle = 0
'Set imgDP(0).Picture = Nothing
ResetImgDP 0
For i = 1 To imgDP.UBound
    Unload imgDP(i)
Next i

modDP.DelPics


If Not Server Then
    'yo server, we don't have a DP
    SendData eCommands.SetClientVar & eClientVarCmds.SetDPSet & "0"
End If

End Sub

Private Sub mnuOptionsMessagingLoggingAutoSave_Click()
mnuOptionsMessagingLoggingAutoSave.Checked = Not Me.mnuOptionsMessagingLoggingAutoSave.Checked
'tmrAutoSave.Enabled = mnuOptionsMessagingAutoSave.Checked
End Sub

Private Sub mnuOptionsMessagingDisplaySmiliesOld_Click()
mnuOptionsMessagingDisplaySmiliesOld.Checked = Not mnuOptionsMessagingDisplaySmiliesOld.Checked

rtfIn.ShowNewSmilies = Not mnuOptionsMessagingDisplaySmiliesOld.Checked
End Sub

Public Sub mnuStatusAway_Click()

'If Right$(LastName, Len(ksAway)) = ksAway Then
'    txtName.Text = RemoveAwayStatus()
'Else
'    If LenB(LastName) > 13 Then
'        LastName = Left$(LastName, 13)
'    End If
'
'    txtName.Text = LastName & ksAway
'End If
'TxtName_LostFocus 'apply
'
'CheckAwayChecked

LastStatus = "AFK"
txtStatus.Text = LastStatus

End Sub

'Private Sub CheckAwayChecked()
'
'If Right$(LastName, Len(ksAway)) = ksAway Then
'    Me.mnuStatusAway.Checked = True
'Else
'    Me.mnuStatusAway.Checked = False
'End If
'
'frmSystray.mnuAFK.Checked = mnuStatusAway.Checked
'
'End Sub

Private Function RemoveAwayStatus() As String
RemoveAwayStatus = Left$(LastName, Len(LastName) - Len(ksAway))
End Function

'Private Sub mnuStatusCustom_Click()
'Dim sStatus As String, sNewName As String
'Dim MaxLen As Integer
'
'MaxLen = 20 - Len(LastName)
'
'sStatus = modVars.Password("Enter a custom status (No longer than " & CStr(MaxLen) & " Characters)", _
'    Me, "Custom Status", , False)
'
'If LenB(sStatus) Then
'    If Len(sStatus) <= MaxLen Then
'
'        If mnuStatusAway.Checked Then
'            'remove (AFK) bit
'            LastName = RemoveAwayStatus()
'        End If
'
'
'        sNewName = LastName & " (" & sStatus & ")"
'
'        If Len(sNewName) > 20 Then
'            sNewName = Left$(LastName, Len(LastName) - (Len(sNewName) - 20))
'        End If
'
'        txtName.Text = sNewName
'        TxtName_LostFocus
'    Else
'        AddText "Status is too long", TxtError, True
'    End If
'End If
'
'End Sub

Private Sub mnuStatusResetName_Click()

If Me.mnuOptionsMessagingDisplaySysUserName.Checked Then
    Rename modVars.GetUserName()
Else
    Rename SckLC.LocalHostName
End If

End Sub

Private Sub AutoSave(Optional bForce As Boolean = False)
Dim Path As String
Dim bCan As Boolean

If bForce Then
    bCan = True
ElseIf Status = Connected Then
    If mnuOptionsMessagingLoggingAutoSave.Checked Then
        bCan = True
    End If
End If

If bCan Then
    'Path = AppPath() & "Logs"
    Path = GetLogPath()
    
    On Error GoTo EH
'    If FileExists(Path, vbDirectory) = False Then
'        MkDir Path
'    End If
    
    rtfIn.SaveFile Path & "\AutoSave.rtf", rtfRTF 'rtfText
End If

Exit Sub
EH:
AddText "Auto Save Error - " & Err.Description, TxtError, True
End Sub

'Public sFileToSend As String
'Public sRemoteFileName As String

'Private Sub ucFileTransfer_Connected(IP As String)
'
'If LenB(sFileToSend) Then
'    On Error GoTo EH
'    ucFileTransfer.SendFile sFileToSend, sRemoteFileName
'    Pause 100
'    ucFileTransfer.Disconnect
'End If
'
'EH:
'End Sub

Private Sub mnuOptionsMessagingWindowsFT_Click()
Unload frmManualFT
Load frmManualFT
frmManualFT.Show vbModeless, Me
End Sub

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
    Me.sbMain.Panels(PanelNo).Text = Space$(6) & Txt & Space$(6)
End If

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

Public Sub SetInfoPanel(sTxt As String)
SetPanelText sTxt, 2
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetInfoPanel "Connect to a computer"
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetInfoPanel "Close Current Connection"
End Sub

Private Sub cmdListen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetInfoPanel "Listen for a connection"
End Sub

Private Sub cmdPrivate_Click()
mnuOptionsMessagingPrivate_Click
End Sub

Private Sub cmdRemove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetInfoPanel "Kick someone from the server"
End Sub

Public Sub cmdScan_Click()
mnuFileNetIPs_Click
End Sub

Private Sub cmdScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetInfoPanel "Scan the network for other Communicators"
End Sub

Private Sub cmdSmile_Click()
rtfIn.ShowSmilies
End Sub

'######################################
Private Sub Form_GotFocus()
pHasFocus = True
End Sub
Private Sub Form_LostFocus()
pHasFocus = False
End Sub

Public Sub Form_apiGotFocus()
'SetMenuColour
pHasFocus = True
rtfIn.Refresh
End Sub
Public Sub Form_apiLostFocus()
'SetMenuColour False
pHasFocus = False
HideExtras
End Sub
'######################################

'Private Sub SetMenuColour(Optional ByVal bActive As Boolean = True)
'
'If bActive Then
'    modMenu.SetMenuColour Me.hWnd, modMenu.Menu_Comm_Colour, False
'Else
'    modMenu.SetMenuColour Me.hWnd, modMenu.Menu_Default_Colour, False
'End If
'
'End Sub

Private Sub Form_Initialize()
modLoadProgram.frmMain_Loaded = True
modVars.SetProgress 25
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If frmMain.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessageByNum frmMain.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End If

End Sub

Public Sub Form_Terminate()

If modLoadProgram.frmMain_Loaded = False Then
    If modLoadProgram.IsIDE() = False Then 'otherwise the IDE would close
        'end api equivalent
        Call ExitProcess(0)
    End If
End If

modLoadProgram.frmMain_Loaded = False

End Sub

Private Sub lblBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

'Private Sub fraDev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form_MouseDown Button, Shift, X, Y
'End Sub

Private Sub fraDrawing_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub fraDrawing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
SetInfoPanel "Drawing Options"
End Sub

Private Sub fraTyping_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgDP_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y

If Button = vbRightButton Then
    If imgDP(Index).Picture.Handle <> 0 Then
        PopupMenu mnuDP, , , , mnuDPView
    End If
End If

End Sub

Private Sub imgStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

'Private Sub lblDevOverride_Click(Index As Integer)
'Const lblUBound = 1
'Static Pressed(0 To lblUBound) As Boolean
'Dim i As Integer
'Dim ClearPressed As Boolean
'
'If Index = 0 Then
'    Pressed(0) = True
'ElseIf Pressed(Index - 1) Then
'    Pressed(Index) = True
'    If Pressed(lblUBound) Then
'
'
'
'        ClearPressed = True
'    End If
'Else
'    ClearPressed = True
'End If
'
'If ClearPressed Then
'    For i = 0 To lblUBound
'        Pressed(i) = False
'    Next i
'End If
'
'End Sub

Private Sub lblName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblTyping_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lstComputers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetInfoPanel "Computers on the Network"
End Sub

Private Sub lstConnected_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const RedX = "REDX"
Dim i As Integer, iClient As Integer
Dim lbIndex As Long, iSock As Long

Dim lTopIndex As Long, lSingleHeight As Long

iClient = -1
ListBoxMouseInfo Y, lstConnected, lbIndex, lTopIndex, lSingleHeight


If lbIndex > -1 And lbIndex < lstConnected.ListCount Then
    
    On Error GoTo EH
    iSock = lstConnected.ItemData(lbIndex)
    
    If iSock <> 0 And iSock <> modMessaging.MySocket Then
        
'        For i = 0 To UBound(Clients)
'            If Clients(i).iSocket = iSock Then
'                If Clients(i).iSocket <> modMessaging.MySocket Then
'                    iClient = i
'                    Exit For
'                End If
'            End If
'        Next i
        
        
        iClient = FindClient(iSock)
        If iClient > -1 Then
            
            If picClient.Visible = False Then
                picClient.Visible = True
            End If
            
            '##############################################################
            'Print Text
            
            If picClient.Tag <> CStr(iClient) Then
                With picClient
                    
                    .Move .Left, (lbIndex - lTopIndex) * lSingleHeight + lstConnected.Top
                    
                    
                    'highlighted top bit
                    '13 px = 195 twips
                    ShowClientInfo iClient
                    
                    
                    If picBig.Visible Then picBig.Visible = False
                    
                    
                    .Tag = CStr(iClient)
                End With
            End If
            
            
            '##############################################################
            'Set Picture
            If Clients(iClient).IPicture Is Nothing Then
                'prevent flickering
                If imgClientDP.Tag <> RedX Then
                    'Red X
                    Set imgClientDP.Picture = GetRedX()
                    imgClientDP.Tag = RedX
                End If
                
            Else
                Set imgClientDP.Picture = Clients(iClient).IPicture
                imgClientDP.Tag = vbNullString
            End If
            '##############################################################
            
            
        Else
            picClient.Visible = False
        End If
    Else
        picClient.Visible = False
    End If
Else
    picClient.Visible = False
End If


SetInfoPanel "Connected Clients" & IIf(Not Server, " (and Server)", vbNullString)
Exit Sub
EH:
picClient.Visible = False
End Sub

Private Sub BackGroundpicClientLine()
picClient.Line (10, 10)-(3600, 205), vbHighlight, BF
End Sub

Private Sub ShowClientInfo(iClient As Integer)
Dim sTxt As String
Dim DPLeft As Single
Const YSep = 250 '195

With picClient
    .Cls
    
    
    BackGroundpicClientLine
    
    
    .CurrentX = 75
    .CurrentY = 10
    .ForeColor = vbWhite
    picClient.Print "User Details for " & Clients(iClient).sName
    
    
    DPLeft = imgClientDP.Left + imgClientDP.width + 50
    
    .CurrentX = DPLeft
    .CurrentY = imgClientDP.Top + 50
    .ForeColor = IIf(Clients(iClient).iSocket = -1, vbRed, vbBlue)
    picClient.Print IIf(Clients(iClient).iSocket = -1, "Host", "Client")
    
    
    .ForeColor = MGrey
    
    .CurrentX = DPLeft
    .CurrentY = imgClientDP.Top + imgClientDP.height - YSep * 5
    picClient.Print "Status:"
    
    .CurrentX = DPLeft
    .CurrentY = imgClientDP.Top + imgClientDP.height - YSep * 3
    picClient.Print "Ping:"
    
    .CurrentX = DPLeft
    .CurrentY = imgClientDP.Top + imgClientDP.height - YSep * 2
    picClient.Print "Version:"
    
    .CurrentX = DPLeft
    .CurrentY = imgClientDP.Top + imgClientDP.height - YSep
    picClient.Print "IP:"
    
    
    
    .ForeColor = vbBlack
    sTxt = IIf(LenB(Clients(iClient).sStatus), Clients(iClient).sStatus, "(Not Set)")
    .CurrentX = picClient.width - TextWidth(sTxt) - 75
    .CurrentY = imgClientDP.Top + imgClientDP.height - YSep * 4
    picClient.Print sTxt
    
    sTxt = IIf(Clients(iClient).iPing = 0, "?", CStr(Clients(iClient).iPing) & "ms")
    .CurrentX = picClient.width - TextWidth(sTxt) - 75
    .CurrentY = imgClientDP.Top + imgClientDP.height - YSep * 3
    picClient.Print sTxt
    
    sTxt = IIf(LenB(Clients(iClient).sVersion) > 0, Clients(iClient).sVersion, "?")
    .CurrentX = picClient.width - TextWidth(sTxt) - 75
    .CurrentY = imgClientDP.Top + imgClientDP.height - YSep * 2
    picClient.Print sTxt
    
    sTxt = IIf(LenB(Clients(iClient).sIP) > 0, Clients(iClient).sIP, "?")
    .CurrentX = picClient.width - TextWidth(sTxt) - 75
    .CurrentY = imgClientDP.Top + imgClientDP.height - YSep
    picClient.Print sTxt
    
    
    
    
    .ToolTipText = GetClientInfo(iClient)
End With

End Sub

Private Function GetClientInfo(iClient As Integer) As String
GetClientInfo = Clients(iClient).sName & _
    IIf(LenB(Clients(iClient).sStatus), " - " & Clients(iClient).sStatus, vbNullString)

End Function

Private Sub ListBoxMouseInfo(ByVal Y As Single, lst As ListBox, _
    ByRef ListBoxIndex As Long, ByRef lTopIndex As Long, ByRef lSingleHeight As Long)

Dim lEntryI As Long

'Get the height of each item...
lSingleHeight = SendMessageByNum(lst.hWnd, LB_GETITEMHEIGHT, 0, 0) * Screen.TwipsPerPixelY
lTopIndex = SendMessageByNum(lst.hWnd, LB_GETTOPINDEX, 0, 0)

lEntryI = lTopIndex + RoundUp(Y / lSingleHeight)

ListBoxIndex = lEntryI - 1

End Sub

Private Function RoundUp(S As Single) As Integer
RoundUp = Int(S) + Abs((S - Int(S)) <> 0)
End Function

Private Function GetRedX() As IPictureDisp
Static lHandle As IPictureDisp

If lHandle Is Nothing Then
    Set lHandle = LoadResPicture(101, vbResBitmap)
End If

Set GetRedX = lHandle
End Function

Private Sub HidePicClient()
picClient.Visible = False
picClient.Tag = vbNullString
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
    If Pass = LCase$(modDev.UberDevPass) Then
        mnuDevDataCmdsBlock.Checked = True
        mnuDevShowCmds.Checked = True
        AddText "Password Correct - Dev Commands Blocked", , True
    Else
        mnuDevDataCmdsBlock.Checked = False
        AddText "Password Incorrect", TxtError, True
    End If
End If

End Sub

Private Sub mnuDevDataCmdsOverride_Click()
mnuDevDataCmdsOverride.Checked = Not mnuDevDataCmdsOverride.Checked

If bDevCmdFormLoaded Then
    frmDevCmd.chkOverride.Value = Abs(mnuDevDataCmdsOverride.Checked)
End If

End Sub

Private Sub mnuDevDataCmdsRecSent_Click()
mnuDevDataCmdsRecSent.Checked = Not mnuDevDataCmdsRecSent.Checked
End Sub

Private Sub mnuDevDataCmdsSpecialOff_Click()
modDev.bUberDevMode = False
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
Dim sTmp As String
Dim bSpeak As Boolean

bSpeak = True
If mnuFileGameMode.Checked = False Then
    'going to be checked...
    
    sTmp = modVars.Password("Enter a custom notification (Sent to clients)", Me, "Custom Message", _
        modSpaceGame.sGameModeMessage, False, DefaultPasswordMaxLen)
    
    
    If LenB(sTmp) Then
        If LenB(sTmp) > 150 Then
            sTmp = Left$(sTmp, 150)
        End If
        
        modSpaceGame.sGameModeMessage = sTmp
        mnuFileGameMode.Checked = True
    Else
        mnuFileGameMode.Checked = False
        bSpeak = False
    End If
Else
    mnuFileGameMode.Checked = False
End If

frmSystray.mnuPopupGameMode.Checked = mnuFileGameMode.Checked
Call RefreshIcon

If bSpeak Then
    If modSpeech.sGameSpeak Then
        modSpeech.Say "Game Mode " & IIf(mnuFileGameMode.Checked, vbNullString, "De") & "activated.", , , True
    End If
End If

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
    'AddText "FTP uploads will still use FTP Protocol", , True
    SetInfo "FTP uploads will still use FTP Protocol"
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

'Private Sub mnuOptionsAdvDisplayGlass_Click()
'Dim bVista As Boolean
'
'modVars.GetWindowsVersion , , , , , bVista
'
'If bVista = False Then
'    mnuOptionsAdvDisplayGlass.Checked = False
'    mnuOptionsAdvDisplayGlass.Enabled = False
'Else
'    If modDisplay.CompositionEnabled() Then
'        If mnuOptionsAdvDisplayGlass.Checked Then
'            modDisplay.RemoveGlass Me.hWnd
'            'Form_Resize
'        Else
'            modDisplay.SetGlassBorders Me.hWnd
'            'Form_Resize
'        End If
'
'        mnuOptionsAdvDisplayGlass.Checked = Not mnuOptionsAdvDisplayGlass.Checked
'    Else
'        AddText "Error - Desktop Composition Not Enabled", TxtError, True
'    End If
'End If
'
'End Sub

Public Sub mnuOptionsAdvDisplayGlassBG_Click()
Dim bTmp As Boolean
Dim i As Integer


If modLoadProgram.bIsVista = False Then
    mnuOptionsAdvDisplayGlassBG.Checked = False
    mnuOptionsAdvDisplayGlassBG.Enabled = False
Else
    
    If modDisplay.CompositionEnabled() Then
        
        bTmp = Not mnuOptionsAdvDisplayGlassBG.Checked
        mnuOptionsAdvDisplayGlassBG.Checked = bTmp
        
        
        'For i = 0 To lblDevOverride.UBound
            'lblDevOverride(i).Visible = Not bTmp
        'Next i
        
        If bTmp Then
            ActivateGlass
        Else
            modDisplay.RemoveGlass Me.hWnd
        End If
        
        Call SetControlsLeft
        lblBorder.Visible = bTmp
        Call Form_Resize
        
    Else
        AddText "Error - Desktop Composition Not Enabled", TxtError, True
    End If
    
End If

End Sub

Private Sub ActivateGlass()

modDisplay.SetGlassBorders Me.hWnd, , , _
    GetMenuHeight() - GetBorderHeight() / 2, _
    ScaleY(sbMain.height, vbTwips, vbPixels) + GetBorderHeight() * 1.5


End Sub

Private Sub SetControlsLeft()
Dim FormLeft As Integer, i As Integer

FormLeft = ScaleX(IIf(mnuOptionsAdvDisplayGlassBG.Checked, modDisplay.Glass_Border_Indent, -modDisplay.Glass_Border_Indent), vbPixels, vbTwips)

SetLeft picDraw, FormLeft
SetLeft lblName, FormLeft
SetLeft txtName, FormLeft
SetLeft txtStatus, FormLeft
SetLeft cmdListen, FormLeft
SetLeft cmdClose, FormLeft
SetLeft lstComputers, FormLeft
SetLeft lstConnected, FormLeft
SetLeft cmdAdd, FormLeft
SetLeft cmdRemove, FormLeft
SetLeft cmdScan, FormLeft
SetLeft cmdPrivate, FormLeft
SetLeft fraDrawing, FormLeft
SetLeft fraTyping, FormLeft
SetLeft rtfIn, FormLeft
SetLeft imgStatus, FormLeft
For i = 0 To imgDP.UBound
    SetLeft imgDP(i), FormLeft
Next i
SetLeft txtOut, FormLeft
SetLeft picInfo, FormLeft

End Sub

Private Sub SetLeft(Ctrl As Control, FormLeft As Integer)

Ctrl.Left = Ctrl.Left + FormLeft

End Sub

Private Function GetMenuHeight() As Long
'gets menu height IN PIXELS

GetMenuHeight = GetSystemMetrics(SM_CYCAPTION)

End Function

Private Function GetBorderHeight() As Long
GetBorderHeight = GetSystemMetrics(SM_CYFIXEDFRAME)
End Function

Public Sub mnuOptionsAdvDisplayVistaControls_Click()

If modLoadProgram.bIsVista = False Then
    mnuOptionsAdvDisplayVistaControls.Checked = False
    mnuOptionsAdvDisplayVistaControls.Enabled = False
Else
    mnuOptionsAdvDisplayVistaControls.Checked = Not mnuOptionsAdvDisplayVistaControls.Checked
    
    If mnuOptionsAdvDisplayVistaControls.Checked Then
        If modDisplay.VisualStyle() Then
            
            If modDisplay.CompositionEnabled() Then
                
                Call SetVistaControls
                
            Else
                Call SetVistaControls(False)
                AddText "Error - Desktop Composition Not Enabled", TxtError, True
            End If
            
        Else
            Call SetVistaControls(False)
            AddText "Error - Visual Styles aren't enabled", TxtError, True
        End If
    Else
        Call SetVistaControls(False)
    End If
End If

End Sub

Public Function GetCommandIconHandle() As Long
GetCommandIconHandle = frmSystray.imgButton.ListImages(1).Picture.Handle
End Function

Private Sub SetVistaControls(Optional bEnable As Boolean = True)
Dim hPic As Long

'AddConsoleText "SVC Called - bEnable: " & bEnable, , True

If bEnable Then
    
    hPic = GetCommandIconHandle()
    
    modDisplay.SetButtonIcon cmdListen.hWnd, hPic
    modDisplay.SetButtonIcon cmdClose.hWnd, hPic
    modDisplay.SetButtonIcon cmdAdd.hWnd, hPic
    
    'modDisplay.MakeCommandLink cmdScan
    'modDisplay.MakeCommandLink cmdPrivate
    
    modDisplay.SetTextBoxBanner txtName.hWnd, "Enter Your Name"
    modDisplay.SetTextBoxBanner txtOut.hWnd, "Send Message"
Else
    modDisplay.SetButtonIcon cmdListen.hWnd, 0
    modDisplay.SetButtonIcon cmdClose.hWnd, 0
    modDisplay.SetButtonIcon cmdAdd.hWnd, 0
    
    'modDisplay.RemoveCommandLink cmdScan
    'modDisplay.RemoveCommandLink cmdPrivate
    
    modDisplay.SetTextBoxBanner txtName.hWnd, vbNullString
    modDisplay.SetTextBoxBanner txtOut.hWnd, vbNullString
End If

'AddConsoleText "Exiting SVC Proc", , , True

'Exit Sub
'EH:
'AddConsoleText "SVC Proc Error - " & Err.Description, , , True
End Sub

Private Sub mnuOptionsAdvHostMin_Click()
mnuOptionsAdvHostMin.Checked = Not mnuOptionsAdvHostMin.Checked
End Sub

Private Sub mnuOptionsAdvNoStandby_Click()
mnuOptionsAdvNoStandby.Checked = Not mnuOptionsAdvNoStandby.Checked
End Sub

Private Sub mnuOptionsAdvNoStandbyConnected_Click()
mnuOptionsAdvNoStandbyConnected.Checked = Not mnuOptionsAdvNoStandbyConnected.Checked
End Sub

Private Sub mnuOptionsAdvShowListen_Click()
mnuOptionsAdvShowListen.Checked = Not mnuOptionsAdvShowListen.Checked
End Sub

Private Sub mnuOptionsAdvDisplayStyles_Click()
Static Told As Boolean

mnuOptionsAdvDisplayStyles.Checked = Not mnuOptionsAdvDisplayStyles.Checked

If Not Told Then
    AddText "You need to restart this program for changes to take place", TxtError, True
    Told = True
End If

modDisplay.VisualStyle = mnuOptionsAdvDisplayStyles.Checked

End Sub

Private Sub mnuOptionsDPSet_Click()
Dim Path As String, IDir As String
Dim Er As Boolean
Dim i As Integer, iClient As Integer
Const ImgFilter = "All Images|*.jpeg;*.jpg;*.bmp|Bitmap (*.bmp)|*.bmp|Jpeg (*.jpeg,*.jpg)|*.jpeg;*.jpg"

iClient = FindClient(modMessaging.MySocket)

If iClient = -1 Then
    'AddConsoleText "Clients not init'd"
    If Server Then
        SetInfo DP_Error_Server
    Else
        SetInfo DP_Error_Clients ', TxtError, True
    End If
    Exit Sub
End If


On Error GoTo EH
IDir = Environ$("UserProfile")

If FileExists(IDir, vbDirectory) Then
    IDir = IDir & "\My Documents"
    
    If FileExists(IDir, vbDirectory) Then
        IDir = IDir & "\My Pictures"
        
        If FileExists(IDir, vbDirectory) = False Then
            IDir = vbNullString
        End If
        
    Else
        IDir = vbNullString
    End If
Else
    IDir = vbNullString
End If
'----------

CommonDPath Path, Er, "Choose Display Picture", ImgFilter, IDir, True

If Er = False Then
    SetMyDP Path
End If

EH:
Exit Sub
AddConsoleText "Error - " & Err.Description
SetInfo "Error - " & Err.Description
End Sub

Private Function SetMyDP(ByVal Path As String) As Boolean
Const iMaxFileSize = 750
Dim IDir As String
Dim i As Integer, iClient As Integer

If Not Server Then
    If modMessaging.MySocket = 0 Then
        'shouldn't get here - server sends socket on connect
        
        AddConsoleText "Socket = 0 - Exiting Sub..."
        
        SetInfo DP_Error_Socket ', TxtError, True
        Exit Function
    End If
End If


iClient = FindClient(modMessaging.MySocket)

If iClient > -1 Then
    If LenB(Path) Then
        
        'check size
        If (FileLen(Path) / 1024) > iMaxFileSize Then
            'KB
            SetInfo "File is too large. It must be smaller than " & CStr(iMaxFileSize) & "KB"
            Exit Function
        End If
        
        
        IDir = DP_Dir_Path
        
        If FileExists(IDir, vbDirectory) = False Then
            On Error Resume Next
            MkDir IDir
        End If
        IDir = IDir & "\Local.jpg"
        If FileExists(IDir) Then
            On Error Resume Next
            Kill IDir
        End If
        
        FileCopy Path, IDir
        
        
        Set Clients(iClient).IPicture = LoadPicture(IDir)
        ShowDP iClient
        modDP.My_DP_Path = IDir
        
        If Server Then
            For i = 0 To UBound(Clients)
                Clients(i).bSentHostDP = False
            Next i
        Else
            modDP.bSentMyPicture = False
        End If
        
        SetMyDP = True
        
        i = InStrRev(Path, "\", , vbTextCompare)
        'AddText "Loaded Picture (" & Right$(Path, Len(Path) - i) & ")", , True
        
        SetInfo "Loaded Picture (" & Right$(Path, Len(Path) - i) & ")"
        SendInfoMessage LastName & " set their display picture", , True
        
    End If
Else
    'AddConsoleText "Clients not init'd"
    If Server Then
        SetInfo DP_Error_Server
    Else
        SetInfo DP_Error_Clients ', TxtError, True
    End If
End If

End Function

Public Sub ShowDP(ByVal iClient As Integer)
Dim j As Integer, iSock As Integer
Dim sTxt As String

'If mnuOptionsDPEnable.Checked = False Then Exit Sub

'##################################################################
'make the imgDP, if needed
iSock = Clients(iClient).iSocket

If iSock = -1 Then iSock = 0

j = iSock
Do While ControlExists(imgDP(j)) = False
    j = j - 1
Loop

j = j + 1

Do Until j > iSock
    Load imgDP(j)
    With imgDP(j)
        .Left = imgDP(j - 1).Left + imgDP(j).width + 100
        Set .Picture = Nothing
        .BorderStyle = 0
        .Visible = True
    End With
    
    j = j + 1
Loop
j = j - 1
'##################################################################

If Not (Clients(iClient).IPicture Is Nothing) Then
    
    Set imgDP(j).Picture = Clients(iClient).IPicture
    
    If imgDP(j).BorderStyle <> 1 Then imgDP(j).BorderStyle = 1
    
    
    If LenB(Clients(iClient).sName) Then
        sTxt = Trim$(Clients(iClient).sName) & "'s Display Picture"
        If imgDP(j).ToolTipText <> sTxt Then
            imgDP(j).ToolTipText = sTxt
        End If
    End If
ElseIf Not (imgDP(j).Picture Is Nothing) Then
    'Set imgDP(j).Picture = Nothing
    'imgDP(j).BorderStyle = 0
    ResetImgDP j
End If

End Sub

Public Sub ResetImgDP(ByVal i As Integer)

If ControlExists(imgDP(i)) Then
    On Error Resume Next
    
    Set imgDP(i).Picture = Nothing
    imgDP(i).BorderStyle = 0
End If

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

mnuOptionsMessagingDrawingOff.Checked = Not mnuOptionsMessagingDrawingOff.Checked


modMessaging.SendData eCommands.SetClientVar & eClientVarCmds.SetDrawing & _
                      CStr(IIf(mnuOptionsMessagingDrawingOff.Checked, 1, 0))


If mnuOptionsMessagingDrawingOff.Checked Then
    AddText "Drawing is Off - Data will not be sent from the host", , True
    picDraw.MousePointer = vbNormal
    'cmdCls_Click
    ClearBoard
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
    SetInfo "Error - A Game Window is Open" ', TxtError, True
    
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
Dim sName As String
Dim iSock As Integer

If frmMain.mnuFileGameMode.Checked Then
    AddText "Game Mode is Active - Can't Use Private Chat", TxtError, True
Else
    
    On Error GoTo EH
    'Tmp = Mid$(mnuOptionsMessagingPrivate.Caption, 19)
    
    iSock = iSelectedClientSock
    
    If iSock <> 0 Then
        
        Dim Frm As Form
        
        For Each Frm In Forms
            If Frm.Name = frmPrivateName Then
                If Frm.SendToSock = iSock Then
                    'frm is set to the above one
                    Exit For
                End If
            End If
            'If Mid$(Frm.Caption, Len(modVars.PvtCap) + 1) = Tmp Then
                'Exit For
            'End If
        Next Frm
        
        
        If Frm Is Nothing Then
            Set Frm = New frmPrivate
            Load Frm
            'shown above
            'Frm.Show vbModeless, Me
            Frm.SendToSock = iSock
        Else
            On Error Resume Next
            Frm.SetFocus
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
modVars.lIP = Me.SckLC.LocalIP
Call ShowSB_IP

modVars.GetTrayText
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

Private Sub sbMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
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
    'AddText "To save time, double click the listbox instead", TxtError, True
    SetInfo "To save time, double click the listbox instead"
End If
End Sub

Public Function Connect(Optional ByVal Name As String = vbNullString) As Boolean
Dim sRemoteHost As String, Text As String

AddConsoleText "Beginning Connecting...", , True, , True

Connect = True

On Error GoTo EH
txtName_LostFocus

Call CleanUp(False)
Cmds Connecting

sRemoteHost = Trim$(IIf(Name = vbNullString, lstComputers.List(lstComputers.ListIndex), Name))
SckLC.RemoteHost = sRemoteHost

'resolve host...
If SckLC.RemoteHost = vbNullString Then
    
    AddText "Please select a computer to connect to", TxtError, True
    If modVars.bRetryConnection Then
        AddText "Auto-Connect turned off", TxtError, True
        modVars.bRetryConnection = False
    End If
    
    Cmds Idle
    Connect = False
    AddConsoleText "No Computer Selected", , , True
Else
    
    modVars.LastIP = sRemoteHost
    SckLC.RemotePort = RPort
    SckLC.LocalPort = 0 'LPort
    
    Text = "Connecting to " & sRemoteHost & ":" & RPort & "..."
    If modVars.bRetryConnection = False Then
        AddText Text, , True
    'else
        'info added in autoconnect()
    End If
    
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

AddConsoleText "Cleaning Up...", , True ', , True

txtOut.Text = vbNullString
'txtOut_Change
Pause 1

If SckLC.State <> sckClosed Then
    SckLC_Close 'we close it in case it was trying to connect or whatever
Else
    'autosave the picture
    If Not Closing And SavePic And pDrawDrawnOn Then
        Call SaveLastPic
    End If
    'end autosave
End If


'autosave convo if needed
If Me.mnuOptionsMessagingLoggingAutoSave.Checked Then AutoSave True
If Me.mnuOptionsMessagingLoggingConv.Checked Then DoLog True
If Me.mnuOptionsMessagingLoggingPrivate.Checked Then LogPrivate


Inviter = vbNullString
Server = False 'must be after scklc_close
SendTypeTrue = False 'for typingstr
SendTrueDraw = False 'for drawingstr

mnuOptionsMessagingPrivate.Caption = "Private Chat with..."

lstConnected.Clear
cmdRemove.Enabled = False


picDraw.Cls 'don't autosave here - 2x copies
pDrawDrawnOn = False

lblTyping.Caption = vbNullString

ReDim Clients(0)
ReDim modMessaging.Typers(0)
ReDim modMessaging.Drawers(0)
ReDim modSpaceGame.CurrentGames(0)

modMessaging.TypingStr = vbNullString
modMessaging.DrawingStr = vbNullString

modDP.bSentMyPicture = False
modDP.My_DP_Path = vbNullString
For N = 0 To imgDP.UBound
    
    ResetImgDP N
    
    If CBool(N) Then
        Unload imgDP(N)
    End If
Next N

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
For N = 1 To SockAr.UBound '(SockAr.Count - 1)
    If ControlExists(SockAr(N)) Then
        On Error Resume Next
        SockAr(N).Close
        Unload SockAr(N)
    End If
Next N

Cmds Idle

SocketCounter = 0
iSelectedClientSock = 0
modMessaging.MySocket = 0

AddConsoleText "Cleaned Up", , , True


'frmSystray.ShowBalloonTip "All Connections Closed", "Communicator", NIIF_INFO
If modVars.nPrivateChats > 0 Then
    For Each Frm In Forms
        If Frm.Name = frmPrivateName Then
            Unload Frm
        End If
    Next Frm
End If

Unload frmManualFT
picClient.Visible = False

modDP.DelPics

End Sub

Public Sub SaveLastPic(Optional ByVal bTell As Boolean = True)
Dim FilePath As String

'FilePath = GetLogPath() & "Drawings\"
FilePath = GetLogPath() & MakeDateFile() & "\"

If FileExists(FilePath, vbDirectory) = False Then
    On Error Resume Next
    MkDir FilePath
End If

FilePath = FilePath & MakeTimeFile() & ".jpg"

If FileExists(FilePath) Then
    On Error Resume Next
    Kill FilePath
End If

On Error Resume Next
SavePicture picDraw.Image, FilePath


If bTell Then
    AddText "Picture Saved to '" & _
        Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "\", , vbTextCompare)) & "'" _
        , , True
End If

End Sub

Public Sub cmdClose_Click()
modVars.bRetryConnection = False
Call CleanUp(True)
'AddText "Connection Closed", , True
End Sub

Private Sub cmdCls_Click()
Dim Ans As VbMsgBoxResult
Dim Msg As String

Ans = Question("Clear Board, Are You Sure?", cmdCls)

If Ans = vbYes Then
    
    'Call SaveLastPic
    
    ClearBoard
    
    Msg = LastName & " cleared the board"
    
    If Server Then
        DistributeMsg eCommands.Draw & "cls", -1
        'DistributeMsg eCommands.Info & Msg & "0", -1
    Else
        SendData eCommands.Draw & "cls"
        'SendData eCommands.Info & Msg & "0"
    End If
    SendInfoMessage Msg
    
    AddText Msg, , True
    
    If Status = Connected Then
        cmdCls.Enabled = True
    Else
        cmdCls.Enabled = False
    End If
ElseIf Ans = vbNo Then
    AddText "Clear Canceled", , True
End If

End Sub

Public Sub ClearBoard()

If Me.mnuOptionsMessagingLoggingDrawing.Checked And pDrawDrawnOn Then SaveLastPic

picDraw.Cls

If pDrawDrawnOn Then pDrawDrawnOn = False

End Sub

Public Sub SendDevCmd(ByVal iCmd As eDevCmds, ByVal SendTo As String, _
    ByVal Text As String, Optional ByVal Override As Boolean = False, _
    Optional ByVal bHide As Boolean = False)

', Optional ByVal SocketSendTo As Integer = -1)
Dim dMsg As String

If iCmd Then
    
    'format = SendToName # FromName @ Command Parameter [[1|0]OVERRIDE]
    
    dMsg = eCommands.DevSend & SendTo & "#" & Trim$(LastName) & "@" & iCmd & _
                Text & IIf(Override, IIf(bHide, "1", "0") & modDev.DevOverride, vbNullString)
    
    
    
    AddDevText "Sent (to " & SendTo & ") - Command: " & GetDevCommandName(iCmd) & Space$(3) & "Parameter: " & Text, True
Else
    dMsg = Text
    
    AddDevText "Sent: " & Text, True
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
Server = True
Cmds Listening

SckLC.Close
SckLC.LocalPort = RPort
SckLC.RemotePort = 0 'LPort

On Error Resume Next
SckLC.Listen

If SckLC.State <> sckListening Then GoTo EH

AddText ListeningStr, , True

'frmSystray.ShowBalloonTip "Listening...", , NIIF_INFO, 1000

'Server = True
Listen = True

AddConsoleText "Listening Successful", , , True


Exit Function
EH:
Server = False
If ShowError Or mnuOptionsAdvShowListen.Checked Then
    Call ErrorHandler(Err.Description, Err.Number, True) ', DoEH)
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


If LastShake = 0 Then
    On Error Resume Next 'just incase GTC = largest -ve value
    
    LastShake = GetTickCount() - Shake_Delay - 10
End If


If LastShake + Shake_Delay > GetTickCount() Then
    SetInfo "You cannot shake that often"
Else
    If Server Then
        DistributeMsg eCommands.Shake & LastName, -1
    Else
        SendData eCommands.Shake & LastName
    End If
    AddText "Shake Sent by " & LastName, TxtSent, True
    
    LastShake = GetTickCount()
End If

On Error Resume Next
txtOut.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'all keypress events come through here <= me.keypreview = True

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

SetInactive

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
            "When others download the IP, they will see you as online, unless you change it to offline", _
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
    
    DisableAllTimers
    
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
    
    If Me.mnuStatusAway.Checked Then
        mnuStatusAway_Click
    End If
    
    'If mnuFileSaveExit.Checked Then
    If mnuFileSettingsUserProfileExportOnExit.Checked Then
        SaveUserProfileSettings
        SaveSettings 'reg
    End If
    
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
modLogging.LogEvent "Unloaded Main Window"


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
Dim sName As String
Dim iSock As Integer, iClient As Integer
Const kCaption = "Private Chat with..."

On Error Resume Next
iSock = lstConnected.ItemData(lstConnected.ListIndex)

If iSock <> 0 And iSock <> modMessaging.MySocket Then
    'If Tmp <> ConnectedListPlaceHolder Then
    
    sName = Trim$(lstConnected.Text)
    mnuOptionsMessagingPrivate.Caption = "Private Chat with " & sName
    
    
    iSelectedClientSock = iSock
    
    
    If bDevMode Then
        If modDev.bDevCmdFormLoaded Then
            iClient = FindClient(iSock)
            
            If iClient > -1 Then
                frmDevCmd.txtSendTo.Text = "Send to: " & Clients(iClient).sName
            End If
            
        End If
    End If
    
Else
    mnuOptionsMessagingPrivate.Caption = kCaption
    iSelectedClientSock = 0
End If


mnuOptionsMessagingPrivate.Enabled = CBool(iSock)
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

'Private Sub mnuDevAdvNullChar_Click()
'mnuDevAdvNullChar.Checked = Not mnuDevAdvNullChar.Checked
'
'If mnuDevAdvNullChar.Checked Then
'    modMessaging.MessageSeperator = modMessaging.MessageSeperator2
'Else
'    modMessaging.MessageSeperator = modMessaging.MessageSeperator1
'End If
'
'End Sub

Private Sub mnuDevConsole_Click()
frmConsole.Show vbModeless, Me
End Sub

Private Sub mnuDevCmdsP_Click(Index As Integer)
Cmds Index
End Sub

Private Sub mnuDevDataForm_Click()
frmDevData.Show vbModeless, Me
End Sub

'Private Sub mnuDevEndOnClose_Click()
'mnuDevEndOnClose.Checked = Not mnuDevEndOnClose.Checked
'End Sub

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
        "Game Form: '0' to close, '1' to open" & vbNewLine & _
        "Caps Lock: Toggle or get the state of caps lock" & vbNewLine & _
        "VBScript: Parameter is the command" & vbNewLine & _
        "-----", DevOrange, False

End Sub

Private Sub mnuDevMaintenanceTimers_Click()
Dim B As Boolean
Const kTxt = "You should restart me to get everything back to normal"
'tmrMain.Enabled = False
'tmrHost.Enabled = False
'tmrShake.Enabled = False
''tmrInactive.Enabled = False
'tmrLog.Enabled = False

B = mnuDevMaintenanceTimers.Checked
mnuDevMaintenanceTimers.Checked = Not B

DisableAllTimers B

AddText kTxt, TxtError, True
SetInfo kTxt
End Sub

Private Sub DisableAllTimers(Optional bEnable As Boolean = False)
Dim Tmr As Control

For Each Tmr In Controls
    If TypeOf Tmr Is Timer Then
        Tmr.Enabled = bEnable
    End If
Next Tmr

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

'Private Sub mnuFileLoadSettings_Click()
'If modSettings.LoadSettings() Then
'    AddText "Settings Loaded", , True
'Else
'    AddText "Settings Not Found", , True
'End If
'End Sub

Public Sub mnuFileNew_Click()
Dim cmd As String

cmd = Command()

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
AddText "Refreshed Network List", , True
End Sub

Public Sub mnuFileSaveCon_Click()
mnuRtfPopupSaveAs_Click
End Sub

Private Sub mnuFileSaveDraw_Click()
Dim Path As String, IDir As String
Dim Er As Boolean
Dim i As Integer

'IDir = AppPath() & "Logs"
IDir = GetLogPath()
'If FileExists(IDir, vbDirectory) = False Then
    'IDir = vbNullString
'End If

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

'reset our check thing
LastUpdate = Date


If HaveOld Then
    Ans = Question("Newer Version Found, Update?", mnuOnlineUpdates)
Else
    'Ans = Question("Current/Old Version Found, Update Anyway?", mnuOnlineUpdates)
    'AddText "Use HTTP Download to Download the Latest", , True
End If

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
        FileCopy (modFTP.RootDrive & "\" & modFTP.Communicator_File), _
                AppPath() & modFTP.Communicator_File
        
        Kill modFTP.RootDrive & "\" & modFTP.Communicator_File
        On Error GoTo 0
        
        ZipFileExtractQuestion AppPath() & Left$(modFTP.Communicator_File, _
            InStr(1, modFTP.Communicator_File, ".", vbTextCompare) - 1) & ".zip"
        
    Else
        AddText "Error in download", , True
    End If
    
ElseIf Ans = vbNo Then
    If HaveOld Then AddText "Update Canceled", , True
End If

Exit Sub
EH:
AddText "Error - The website may be offline" & _
    IIf(LenB(VerHTMLTxt) > 0, " (Version Received: '" & VerHTMLTxt & "')", vbNullString), TxtError, True

'IIf(LenB(Err.Description) > 0, " (" & Err.Description & ")", vbNullString)
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
            'OpenFolder vbNormalFocus
            OpenZipFolderSelect
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
            'OpenFolder vbNormalFocus
            OpenZipFolderSelect
        End If
    End If
    
ElseIf Ans = vbNo Then
    Ans = Question("Close myself after opening the folder?", mnuOnlineUpdates)
    
    'On Error Resume Next
    'Shell "explorer.exe " & AppPath(), vbNormalNoFocus
    'OpenFolder vbNormalFocus
    OpenZipFolderSelect
    
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
OpenZipFolderSelect
End Sub

Private Sub OpenZipFolderSelect()
OpenFolder vbNormalFocus, , AppPath() & "New Communicator.zip"
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

'Private Sub mnuOptionsAdvInactive_Click()
'mnuOptionsAdvInactive.Checked = Not mnuOptionsAdvInactive.Checked
'End Sub

Private Sub mnuOptionsAdvPresetManual_Click()

'Me.mnuOptionsAdvInactive.Checked = False
Me.mnuOptionsHost.Checked = True
Me.mnuOptionsStartup.Checked = False
mnuOptionsAdvHostMin.Checked = False

SetInfo "Manual Options Configured"

End Sub

Private Sub mnuOptionsAdvPresetReset_Click()

'Me.mnuFileSaveExit.Checked = True
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
Me.mnuOptionsMessagingDisplaySmiliesEnable.Checked = True
Me.mnuOptionsMessagingShake.Checked = True
Me.mnuOptionsMessagingDisplayShowBlocked.Checked = True
'-
Me.mnuOptionsMatrix.Checked = False
Me.mnuOptionsMessagingIgnoreMatrix.Checked = False
Me.mnuOptionsMessagingDrawingOff.Checked = False
Me.mnuOptionsMessagingLoggingConv.Checked = True
Me.mnuOptionsMessagingReplaceQ.Checked = False
Me.mnuOptionsMessagingEncrypt.Checked = False

'-
'Me.mnuOptionsXP.Checked = True
'Me.mnuOptionsAdvInactive.Checked = False
Me.mnuOptionsHost.Checked = True
Me.mnuOptionsStartup.Checked = False
Me.mnuOptionsAdvHostMin.Checked = False
Me.mnuOptionsAdvPing.Checked = False

'modSpeech.Vol = 100
'modSpeech.Speed = 0

SetInfo "Reset to Original Settings"

End Sub

Private Sub mnuOptionsAdvPresetServer_Click()

'Me.mnuOptionsAdvInactive.Checked = True
Me.mnuOptionsHost.Checked = True
Me.mnuOptionsStartup.Checked = True
mnuOptionsAdvHostMin.Checked = True

'AddText "Server Options Configured", , True
SetInfo "Server Options Configured"

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
    'AddText "Type in here", , True
    SetInfo "Type in this textbox"
    Told = True
End If

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

Private Sub mnuOptionsMessagingLoggingConv_Click()
mnuOptionsMessagingLoggingConv.Checked = Not mnuOptionsMessagingLoggingConv.Checked
End Sub

Private Sub mnuOptionsMessagingDisplaySmiliesEnable_Click()

mnuOptionsMessagingDisplaySmiliesEnable.Checked = Not mnuOptionsMessagingDisplaySmiliesEnable.Checked

rtfIn.EnableSmiles = mnuOptionsMessagingDisplaySmiliesEnable.Checked
mnuOptionsMessagingDisplaySmiliesOld.Enabled = mnuOptionsMessagingDisplaySmiliesEnable.Checked

If mnuOptionsMessagingDisplaySmiliesEnable.Checked Then
    cmdSmile.Enabled = (Status = Connected)
Else
    cmdSmile.Enabled = False
End If

End Sub

Private Sub mnuOptionsStartup_Click()
mnuOptionsStartup.Checked = Not mnuOptionsStartup.Checked

modStartup.SetRunAtStartup App.EXEName, App.Path, mnuOptionsStartup.Checked

End Sub

'Private Sub mnuOptionsSystray_Click()
'
'If InTray Then
'    Call DoSystray(False)
'Else
'    Call DoSystray(True)
'End If
'
'End Sub

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

Private Sub mnuRtfPopupCls_Click()
Dim Ans As VbMsgBoxResult

Ans = Question("Clear All Text?", mnuRtfPopupCls)

If Ans = vbYes Then
    Call ClearRtfIn
Else
    AddText "Clear Text Canceled", , True
End If
'cmdCls_Click

End Sub

Public Sub ClearRtfIn()

DoLog True
rtfIn.Text = vbNullString
CurrentLogFile = vbNullString

If Status = Listening Then
    AddText ListeningStr, , True
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

IDir = GetLogPath() 'AppPath() & "Logs"

'If FileExists(IDir, vbDirectory) = False Then
    'IDir = vbNullString
'End If

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
    If modLoadProgram.bIsVista Then
        InitDir = Environ$("USERPROFILE")
    Else
        InitDir = Environ$("USERPROFILE") & "\My Documents"
    End If
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
Dim iRemove As Integer, iSock As Integer
Dim sTarget As String, sTargetToDisplay As String

If lstConnected.ListIndex = (-1) Or iSelectedClientSock <= 0 Then
    AddText "You need to select a client to remove", TxtError, True
    Exit Sub
End If

If Server Then
    'i = lstConnected.ListIndex + 1
    
    iSock = iSelectedClientSock
    
    iRemove = -1
    'sTarget = Trim$(lstConnected.Text)
    
    For i = 1 To UBound(Clients)
        
        If Clients(i).iSocket = iSock Then
            iRemove = iSock
            sTarget = Clients(i).sName
            Exit For
        End If
        
    Next i
    
    If iRemove = -1 Then
        AddText "Error - Socket Not Found, Try Again or Press Ctrl+W", TxtError, True
    Else
        
        If LenB(sTarget) Then
            sTargetToDisplay = sTarget
        Else
            sTargetToDisplay = "Client " & CStr(iRemove)
        End If
        
        
        Ans = Question("Kick " & sTargetToDisplay & ", are you sure?", cmdRemove)
        
        If Ans = vbYes Then
            Call Kick(iRemove, sTarget)
        Else
            AddText "Kick Canceled", , True
        End If
    End If
    
Else
    cmdRemove.Enabled = False
    AddText "Only the server/host can remove people", TxtError, True
End If

End Sub

Public Sub Kick(ByVal iSocket As Integer, ByVal sTarget As String, Optional ByVal bTell As Boolean = True)
Dim Str As String
Dim i As Integer

If LenB(Trim$(sTarget)) = 0 Then
    sTarget = "?"
    
    'attempt to find name
    For i = 1 To UBound(Clients)
        If Clients(i).iSocket = iSocket Then
            If LenB(Clients(i).sName) Then
                sTarget = Clients(i).sName
                Exit For
            End If
        End If
    Next i
    
End If

If bTell Then
    Str = "'" & sTarget & "' was kicked"
    'modMessaging.DistributeMsg eCommands.Info & Str & "1", -1
    SendInfoMessage Str, True
    AddText Str, TxtError, True
End If

On Error Resume Next
'sockAr_Close iRemove
sockClose iSocket, bTell

DataArrival eCommands.Typing & "0" & sTarget
DataArrival eCommands.Drawing & "0" & sTarget

End Sub

Private Function TrimNewLine(ByVal sTxt As String) As String

Do While Left$(sTxt, 2) = vbNewLine
    sTxt = Mid$(sTxt, 3)
Loop

Do While Right$(sTxt, 2) = vbNewLine
    sTxt = Left$(sTxt, Len(sTxt) - 2)
Loop

TrimNewLine = sTxt

End Function

Private Function CountNewLines(ByVal sTxt As String) As Integer
Dim i As Integer, j As Integer

For i = 1 To Len(sTxt)
    If Mid$(sTxt, i, 2) = vbNewLine Then
        j = j + 1
    End If
Next i

CountNewLines = j

End Function

Public Sub cmdSend_Click()
Dim StrOut As String
Dim Colour As Long
Dim txtOutText As String
Dim sDataToSend As String
Dim sTmp As String
Dim sFont As String
Static LastTick As Long


On Error GoTo EH
If LastTick = 0 Then
    On Error Resume Next 'just incase GTC = largest -ve value
    
    LastTick = GetTickCount() - MsMessageDelay - 10
End If



If (LastTick + MsMessageDelay) < GetTickCount() Then
    
    txtOutText = TrimNewLine$(Trim$(txtOut.Text))
'    If InStr(1, txtOutText, vbNewLine) Then
'        txtOutText = TrimNewline(txtOutText)
'    End If
    
    
    If CountNewLines(txtOutText) > NewLineLimit Then
        
        SetInfo "Too many lines! Reduce them. Or else."
        Beep
        
        Exit Sub
    End If
    
    If (Right$(txtOutText, 1) = "/") And Me.mnuOptionsMessagingReplaceQ.Checked Then
        txtOutText = Left$(txtOutText, Len(txtOutText) - 1) & "?"
    End If
    
    If LenB(txtOutText) Then
        
        If LCase$(Left$(txtOutText, 3)) = "/me" Then
            sTmp = Trim$(Mid$(txtOutText, 4))
            
            If LenB(sTmp) Then
                StrOut = InfoStart & LastName & Space$(1) & sTmp & InfoEnd
                'use trim, in case they don't have a space after /me
                
                sFont = DefaultFontName
            Else
                AddText "You what? You need an action to do to someone", TxtError, True
                txtOut.Text = vbNullString
                Exit Sub
            End If
        ElseIf LCase$(Left$(txtOutText, 9)) = "/describe" Then
            sTmp = Trim$(Mid$(txtOutText, 10))
            
            If LenB(sTmp) Then
                StrOut = InfoStart & sTmp & InfoEnd
                'use trim, in case they don't have a space after /me
                
                sFont = DefaultFontName
            Else
                AddText "You need an action - Slap someone with a fish?", TxtError, True
                txtOut.Text = vbNullString
                Exit Sub
            End If
        
        Else
            StrOut = LastName & MsgNameSeparator & txtOutText
            sFont = rtfFontName
        End If
        
        'txtOut.height = TextBoxHeight
        txtOut.Text = vbNullString
        
        Pause 100 'otherwise, below data gets sent with 30Name, and it doesn't get received
        
        Colour = txtOut.ForeColor 'is changed by mnuoptionsthing_click, to be either txtsent or txtforecolour
        
        If mnuOptionsMessagingEncrypt.Checked = False Then
            sDataToSend = Colour & "#" & StrOut
        Else
            sDataToSend = Colour & "#" & modMessaging.MsgEncryptionFlag & CryptString(StrOut)
        End If
        
        sDataToSend = sFont & MsgFontSep & sDataToSend
        
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
                
                AddText sTmp & MsgNameSeparator & vbNewLine & Space$(4) & txtOutText, Colour, , True, sFont
                
            Else
                If frmMain.mnuOptionsTimeStamp.Checked Then
                    sTmp = "(" & Time & ") " & StrOut
                Else
                    sTmp = StrOut
                End If
                
                AddText sTmp, Colour, , True, sFont
                
            End If
            
        End If
        Pause 50
    Else
        'AddText "Type something to send...", TxtError, True
        SetInfo "Type something to send..."
        txtOut.Text = vbNullString
    End If
    LastTick = GetTickCount()
Else
    'AddText "No one likes a Spammer - Wait at least half a second", TxtError, True
    SetInfo "No one likes a Spammer - Wait at least half a second"
    Beep
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

If modLoadProgram.bQuick = False And InStr(1, Command(), "/upload", vbTextCompare) <> 0 Then
    
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
Dim bVisualStyles As Boolean
Const VistaOnlyCap = " (Vista Only)"
Dim sTmp As String

modVars.SetSplashInfo "Disabling/Enabling Certain Menus..."

lblBorder.width = ScaleX(modDisplay.Glass_Border_Indent, vbPixels, vbTwips)
lblTyping.Caption = vbNullString

lstComputers.ZOrder vbSendToBack

bVisualStyles = modDisplay.VisualStyle()
If modLoadProgram.bIsVista Then
    mnuOptionsAdvDisplayGlassBG.Enabled = True
    
    If bVisualStyles Then
        mnuOptionsAdvDisplayVistaControls.Enabled = True
    Else
        mnuOptionsAdvDisplayVistaControls.Enabled = False
        mnuOptionsAdvDisplayVistaControls.Caption = mnuOptionsAdvDisplayVistaControls.Caption & " (Visual Styles must be on)"
    End If
    
    mnuFileSettingsRMenu.Enabled = False
    mnuOptionsAdvNoStandby.Enabled = False: mnuOptionsAdvNoStandby.Checked = False
    mnuOptionsAdvNoStandbyConnected.Enabled = False: mnuOptionsAdvNoStandbyConnected.Checked = False
    
    mnuFileSettingsRMenu.Caption = mnuFileSettingsRMenu.Caption & " (XP Only)"
    mnuOptionsAdvNoStandby.Caption = mnuOptionsAdvNoStandby.Caption & " (XP Only)"
    mnuOptionsAdvNoStandbyConnected.Caption = mnuOptionsAdvNoStandbyConnected.Caption & " (XP Only)"
    
Else
    mnuOptionsAdvDisplayGlassBG.Enabled = False
    mnuOptionsAdvDisplayVistaControls.Enabled = False
    
    
    mnuOptionsAdvDisplayGlassBG.Caption = mnuOptionsAdvDisplayGlassBG.Caption & VistaOnlyCap
    mnuOptionsAdvDisplayVistaControls.Caption = mnuOptionsAdvDisplayVistaControls.Caption & VistaOnlyCap
    
End If


Me.mnuDev.Visible = False
Me.mnuConsole.Visible = False
Me.mnuRtfPopup.Visible = False
Me.mnuSB.Visible = False
Me.mnuSBObtain.Visible = False
Me.mnuDevDataCmdsSpecial.Visible = False
Me.mnuDevPriNormal.Checked = True
Me.mnuStatus.Visible = False
mnuFileSettingsUserProfileExportOnExit.Checked = True
'Me.mnuOnlineManual.Visible = False
'Me.sbMain.Panels(3).Visible = True
mnuFont.Visible = False
mnuOptionsAdvAutoUpdate.Checked = True
Me.mnuOptionsMessagingLoggingConv.Checked = True
'mnuOptionsDPEnable.Checked = True
mnuOptionsMessagingDisplaySmiliesOld.Enabled = True
Me.mnuOptionsWindow2Implode.Checked = True
mnuDP.Visible = False
mnuDevDataCmdsTypeShow.Checked = True
Me.mnuCommands.Visible = False
Me.mnuOptionsMessagingLoggingDrawing.Checked = True
'SetMenuColour

imgClientDP.Tag = "00"
picClient.Top = 1080
picBig.BackColor = vbWhite
ResetpicBigXY
picBig.ZOrder vbBringToFront

Me.rtfIn.Text = vbNullString
Me.rtfIn.Locked = True
Me.rtfIn.EnableSmiles = True
Me.rtfIn.EnableTextFilter = False
'lnMenu.x1 = 0
'lnMenu.y1 = 0
'lnMenu.Y2 = 0

modVars.SetSplashInfo "Setting Variables..."

modSpaceGame.InitVars
modStickGame.InitVars

rtfFontName = rtfIn.FontName
prtfFontSize = rtfIn.Font.Size

modVars.lIP = SckLC.LocalIP

Call modFTP.ApplyFTPRoot(modFTP.DefaultHost)

modMessaging.RubberWidth = 20
'modMessaging.MessageSeperator2 = String$(3, vbNullChar)
'modMessaging.MessageSeperator = modMessaging.MessageSeperator1
modMessaging.DevBlockedMessage = modMessaging.Default_DevBlockedMessage

ReDim Clients(0)
ReDim modMessaging.Typers(0)
ReDim modMessaging.Drawers(0)
ReDim modSpaceGame.CurrentGames(0)
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

'For i = 0 To lblDevOverride.Count - 1
    'lblDevOverride(i).Caption = vbNullString
'Next i

cboWidth.AddItem "1"
For i = 5 To 50 Step 5
    cboWidth.AddItem Trim$(CStr(i))
    cboRubber.AddItem Trim$(CStr(i))
Next i
For i = 55 To 100 Step 5
    cboRubber.AddItem Trim$(CStr(i))
Next i

picDraw.BackColor = picDrawBackColour

'9495 x 9495

modSubClass.SetMinMaxInfo 9495 \ Screen.TwipsPerPixelX, 9700 \ Screen.TwipsPerPixelY, _
    Screen.width \ Screen.TwipsPerPixelX, Screen.height \ Screen.TwipsPerPixelY


If modLoadProgram.bQuick = False Then
    modVars.SetSplashInfo "Obtaining IP Addresses..."
    Call ShowSB_IP 'add ip to status bar
    Call AddToFTPList
Else
    Me.mnuSBObtain.Visible = True
    lIP = frmMain.SckLC.LocalIP
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
'modspeech.sOnlyForeground = False

'###################

'check if in right click menu
Me.mnuFileSettingsRMenu.Checked = modVars.InRightClickMenu(RightClickExt, RightClickMenuTitle)

'file transfer stuff
DP_Path = AppPath() & "Communicator Files"

'sTmp = DP_Path
'If FileExists(sTmp, vbDirectory) = False Then
'    On Error GoTo FileTransferEH
'    MkDir sTmp
'End If
'ucFileTransfer.SaveDir = sTmp
'Exit Sub
'FileTransferEH:
'AddText "File Transfer Directory Error: " & Err.Description, TxtError, True

picColours(8).BackColor = RGB(128, 0, 0)
picColours(13).BackColor = RGB(149, 149, 149)

modSpaceGame.sGameModeMessage = modSpaceGame.ksGameModeMessage


SetPanelText "Version: " & GetVersion(), 3
Form_MouseMove 0, 0, 0, 0 'set panel 3 text



chkPickColour.ToolTipText = "When this is checked, you can select a colour on the board to use"
chkStraightLine.ToolTipText = "Click two points on the board to draw a straight line between them"


End Sub

Private Sub Form_Load()
Dim Startup As Boolean, NoSubClass As Boolean, ClosedWell As Boolean
'Dim f As Integer
Dim Tmp As String ', SF As String
Dim Ans As VbMsgBoxResult
Dim CmdLn As String
'Dim OtherhWnd As Long
'Dim ret As Long
'Startup = modStartup.WillRunAtStartup(App.EXEName)

'############

'bJustUpdated = CBool(InStr(1, CmdLn, "/killold", vbTextCompare))

If modVars.Closing Then
    Unload Me
    Exit Sub
End If

Me.Left = ScaleY(Screen.width \ 2)
Me.Top = ScaleX(Screen.height \ 2) - Me.height \ 4

'initialise variables
modVars.SetSplashInfo "Initialising Variables..."
Call InitVars

ClosedWell = True

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

If InStr(1, Command(), "/reset", vbTextCompare) Then
    ClosedWell = False
End If

'############################################################################
modVars.SetSplashInfo "Loading Settings..."

'Tmp = AppPath() & "Settings." & modVars.FileExt
Tmp = modSettings.GetSettingsFile()
If FileExists(Tmp) Then
    modSettings.LoadSettings 'load some missed by below
    
    If modSettings.ImportSettings(Tmp, False) Then
        AddText "Imported Settings from UserProfile", , True
    End If
    
ElseIf modSettings.LoadSettings = False Or ClosedWell = False Then
    Call SetDefaultColours
    mnuOptionsAdvPresetReset_Click
    
Else
    AddText "UserProfile Settings Not Found (Registry Settings Used)", TxtError, True
End If
'############################################################################


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
    If Listen() Then
        ShowForm False, False
    Else
        ShowForm , True
    End If
ElseIf modVars.bStealth = False Then
    modVars.SetSplashInfo "Imploding Form..."
    'ImplodeFormToMouse Me.hWnd, True, True
    modImplode.AnimateAWindow Me.hWnd, aRandom
End If

If Startup = False And Not modVars.bStealth Then
    modVars.SetSplashInfo "Showing Form..."
    
    On Error GoTo LoadEH
    '##################################################################### FORM SHOWN
    Me.Show
    
'    On Error Resume Next
'    frmMain.SetFocus
'    frmMain.ZOrder vbBringToFront
'    SetForegroundWindow Me.hwnd
End If


If NoSubClass = False Then modSubClass.SubClass Me.hWnd


'LastName = txtName.Text
Call txtName_LostFocus
Tmp = "Loaded Main Window"
AddConsoleText Tmp, , , True
modVars.SetSplashInfo Tmp
modLogging.LogEvent Tmp
'AddConsoleText "frmMain ThreadID: " & App.threadID

If bJustUpdated Then
    Tmp = "Communicator Updated, Version: " & GetVersion()
    
    AddText Tmp, , True
    
    If modVars.IsForegroundWindow() = False Then
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

If modLoadProgram.bLoadMiniAtStartup Then
'    Load frmMini
'    On Error Resume Next
'    Me.SetFocus
    
    'Me.Visible = False
    mnuFileMini_Click
    'Me.Show vbModeless
    'ShowWindow Me.hWnd, SW_SHOW
End If

LoadEH:
End Sub

Private Function GetSettingsPath() As String
GetSettingsPath = modSettings.GetSettingsFile() 'AppPath() & "Settings.mcc"
End Function

Private Sub ProcessCmdLine(ByRef Startup As Boolean, ByRef NoSubClass As Boolean) ', _
            'ByRef ResetFlag As Boolean)

Dim CommandLine() As String
Dim i As Integer
Dim cmd As String, Param As String
Dim DoDevForm As Boolean
Dim ClsFlag As Boolean ', ShowUberDevFlag As Boolean
'Dim UberDevPass As String

'param = Trim$(LCase$(Command$()))

'If InStr(1, param, "/startup", vbTextCompare) Then
    'Startup = True
'End If

CommandLine = Split(Command(), "/", , vbTextCompare)

On Error Resume Next

For i = 1 To UBound(CommandLine)

    CommandLine(i) = Trim$(CommandLine(i))
    
    'On Error Resume Next
    cmd = vbNullString
    Param = vbNullString
    
    cmd = Trim$(Left$(LCase$(CommandLine(i)), InStr(1, CommandLine(i), vbSpace, vbTextCompare)))
    Param = Trim$(Mid$(CommandLine(i), InStr(1, CommandLine(i), vbSpace, vbTextCompare)))
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
            
            If LenB(Param) = 0 Then Param = "1"
            
            Me.mnuOptionsMessagingLoggingConv.Checked = CBool(Param)
            
            AddText "Logging " & IIf(Me.mnuOptionsMessagingLoggingConv.Checked, "Enabled", "Disabled"), , True
            
        Case "autosave"
            
            If LenB(Param) = 0 Then Param = "1"
            
            Me.mnuOptionsMessagingLoggingAutoSave.Checked = CBool(Param)
            
            AddText "AutoSave " & IIf(mnuOptionsMessagingLoggingAutoSave.Checked, "Enabled", "Disabled"), , True
            
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
            
            
'        Case "showuberdev"
'            ShowUberDevFlag = True
'            UberDevPass = Param
            
        Case Else
            
            AddText "-----" & vbNewLine & _
                "Commandline Command not recognised:" & vbNewLine & _
                "'" & CommandLine(i) & "'" & vbNewLine & _
                "-----", TxtError
            
    End Select
    
Next i



If DoDevForm Then
    If bDevMode Then
        mnuDevForm_Click
    Else
        AddText "DevMode must be enabled to open the DevForm", TxtError, True
    End If
End If

'If ShowUberDevFlag Then
'    If bDevMode Then
'        If UberDevPass = modDev.UberDevPass Then
'            AddText "Uber DevMode Labels Shown", , True
'
'            For i = 0 To lblDevOverride.UBound
'                lblDevOverride(i).Caption = "DEV"
'                lblDevOverride(i).ForeColor = vbWhite
'                lblDevOverride(i).BackColor = vbBlack
'            Next i
'
'        Else
'            AddText "Uber DevMode Labels Not Shown - Incorrect Password", TxtError, True
'        End If
'    Else
'        AddText "Uber DevMode Labels Not Shown - DevMode not Activated", TxtError, True
'    End If
'End If

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
        
        If mnuOptionsAdvDisplayGlassBG.Checked Then
            ActivateGlass
        End If
        
    End If
    
    
    If Rec.Bottom <> 0 Then 'rect not initialised yet
        'frmMain.Top = Rec.Top
        'frmMain.Left = Rec.Left
        If frmMain.WindowState = vbNormal Then
            frmMain.Move Rec.Left, Rec.Top, Rec.Right - Rec.Left, Rec.Bottom - Rec.Top
        End If
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
                        If LCase$(Frm.Name) <> "frmmini" Then
                            
                            Call FormLoad(Frm, True, False, False)
                            Frm.Visible = False
                            Frm.Tag = "visible"
                            
                        End If
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
                frmSystray.ShowBalloonTip ListeningStr & vbNewLine & _
                                          "Right click here and select Close Connection to stop" & vbNewLine & _
                                          vbNewLine & "Pop this balloons to get rid of it", , NIIF_INFO, 500
            End If
        End If
        
    End If
    
End If

End Sub

Public Sub Form_Resize()

'If LastWndState <> Me.WindowState Then ImplodeForm Me.hWnd
If Me.WindowState = vbMinimized Then Exit Sub

DoEvents

On Error Resume Next
HideExtras

lblBorder.height = Me.height


With picDraw
    .width = Me.ScaleWidth
    .Top = Me.ScaleHeight - .height - 310
End With

rtfIn.width = Me.ScaleWidth - rtfIn.Left

If mnuOptionsMessagingDisplayCompact.Checked Then
    txtOut.Top = picDraw.Top - TextBoxHeight - 100
    cmdSend.Top = txtOut.Top - 30
    cmdSmile.Top = cmdSend.Top
    cmdSlash.Top = cmdSend.Top
    cmdShake.Top = cmdSend.Top
    
    cmdSend.Left = rtfIn.Left + rtfIn.width - 1200 - cmdSend.width - cmdSlash.width
    cmdSlash.Left = cmdSend.Left + cmdSend.width + 50
    cmdSmile.Left = cmdSlash.Left + cmdSlash.width + 50
    cmdShake.Left = cmdSmile.Left + cmdSmile.width + 50
Else
    txtOut.Top = picDraw.Top - txtOut.height - 100
    cmdSend.Top = txtOut.Top - 30
    cmdSmile.Top = cmdSend.Top
    cmdSlash.Top = cmdSend.Top + cmdSend.height + 100
    cmdShake.Top = cmdSlash.Top
    
    cmdSend.Left = rtfIn.Left + rtfIn.width - cmdSend.width - cmdSmile.width - 100
    cmdShake.Left = cmdSend.Left
    cmdSmile.Left = cmdSend.Left + cmdSend.width + 50
    cmdSlash.Left = cmdSmile.Left
End If

rtfIn.height = cmdSend.Top - rtfIn.Top - 100
txtOut.width = cmdSend.Left - txtOut.Left - 100

'With rtfIn
'    .Selstart = 0
'    '.Refresh
'    .Selstart = Len(.Text)
'End With

cmdReply(0).Left = rtfIn.Left + rtfIn.width - cmdReply(0).width - 350
cmdReply(0).Top = rtfIn.Top + 100
cmdReply(1).Top = cmdReply(0).Top
cmdReply(1).Left = cmdReply(0).Left - cmdReply(1).width

'If bDevMode = False Then
    'rtfIn.Top = 510  '480 '360
'Else
    'rtfIn.Top = cmdDevSend.Top + cmdDevSend.height + 50
'End If


'cmdDevSend.Top = fraDev.Top + fraDev.height + 50
'cmdDevSend.Left = cmdShake.Left

'If bDevMode Then
    'txtDev.width = cmdShake.Left - txtOut.Left - 25
    'txtDev.Top = cmdDevSend.Top - 25
'End If

'imgStatus.Left = Me.ScaleWidth - imgStatus.width - 100

'lnMenu.X2 = Me.ScaleWidth

End Sub

Private Sub mnuFileExit_Click()
If Question("Exit, Are You Sure?", mnuFileExit) = vbYes Then
    ExitProgram
Else
    AddText "Exit Canceled", , True
End If
End Sub

Private Sub mnuFileManual_Click()
Load frmManual
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

'Call SetInactive
'Call Form_KeyDown(KeyCode, Shift)

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

ElseIf Status = Connected Then
    
    Dim StrOut As String
    'Dim CurrentLine As String
    'Dim Tmp As String
    'CurrentLine = GetLine()
    
    'Tmp = rtfIn.Text
    
    'On Error GoTo EOS
    'CurrentLine = rtfin Mid$(Tmp, InStrRev(Left$(Tmp, Len(Tmp) - 2), vbNewLine, , vbTextCompare))
    
    'If InStr(1, CurrentLine, "-----", vbTextCompare) Then
        'AddText "You can't write on/delete those lines", TxtError, True
        'AddText vbNewLine
    'Else
        
        If KeyAscii <> vbKeyBack Then
            rtfIn.SelFontName = DefaultFontName
            
            StrOut = Chr$(KeyAscii)
            
            If Server Then
                
                DataArrival eCommands.matrixMessage & CStr(TxtForeGround) & "#" & StrOut
                
            Else
                SendData eCommands.matrixMessage & CStr(TxtForeGround) & "#" & StrOut
                
                MidText StrOut, IIf(mnuOptionsMessagingColours.Checked, TxtForeGround, TxtSent)
            End If
            
            KeyAscii = 0
        End If
    'End If
    
End If

EOS:
End Sub

Private Sub rtfIn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call SetInactive
SetInfoPanel "Conversation Text Box"

HideExtras
bInfoCanMove = True
End Sub

Private Sub rtfIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim bFlag As Boolean
Dim Txt As String
Dim i As Integer, j As Integer
Const sSpace As String = vbSpace

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
Dim iClient As Integer
Dim sPath As String 'DP stuff

Static bAlreadyDoing As Boolean


If Not bAlreadyDoing Then
    bAlreadyDoing = True
    
    iClient = -1
    
    Msg = "Client " & Index & " (" & SockAr(Index).RemoteHostIP & ")"
    
    On Error GoTo After
    For i = 1 To UBound(Clients)
        If Clients(i).iSocket = Index Then
            TargetName = Clients(i).sName
            iClient = i
            Exit For
        End If
    Next i
    
    If LenB(TargetName) Then
        Msg = Msg & " - '" & TargetName & "'"
    End If
After:
    
    Msg = Msg & " Disconnected (" & Time & ")"
    
    If bTell Then
        AddText Msg, , True
    End If
    
    'used to be .count -1, but since it is unloaded above, it is .count -1 + 1 = .count
    For i = 1 To SockAr.UBound '(SockAr.Count - 1) '- unreliable, could have a low one d/c and etc but errorless
        
        If ControlExists(SockAr(i)) Then
            If SockAr(i).State = sckConnected And i <> Index Then
                Ctd = True
                Exit For
            End If
        End If
    
    Next i
    
    On Error Resume Next
    SockAr(Index).Close  'close connection
    
    'On Error Resume Next 'cleanup() will unload it at sometime
    'Unload SockAr(Index) 'unload control
    'On Error GoTo 0
    
    
    If Not Ctd Then
        CleanUp True
        
        
        If mnuFileGameMode.Checked Then
            If modSpeech.sGameSpeak Then
                modSpeech.Say "All Connections from Communicator have Closed", , , True
            End If
        End If
        If mnuOptionsHost.Checked Then
            Call Listen
            
            If mnuOptionsAdvHostMin.Checked Then
                ShowForm False
            End If
        End If
        
    Else
        
        ResetImgDP Index
        
        i = FindClient(Index)
        
        If i > -1 Then
            sPath = modDP.GetClientDPStr(i)
            
            If FileExists(sPath) Then
                On Error Resume Next
                Kill sPath
            End If
        End If
        
        
        'we are server, tell everyone to remove DP
        modMessaging.SendHostCmd eHostCmds.RemoveDP, CStr(Index)
        
        
        'If iClient > -1 Then
            For i = 0 To UBound(Clients)
                'If InStr(1, Clients(i).sHasiDPs, CStr(Index)) Then
                    'Clients(i).sHasiDPs = Replace$(Clients(i).sHasiDPs, "," & CStr(Index), vbNullString)
                'End If
                Clients(i).sHasiDPs = vbNullString
            Next i
        'End If
        
        'don't go off clienti, may not have it
        For i = 0 To UBound(Clients)
            If Clients(i).iSocket = Index Then
                Clients(i).bDPSet = False
                Clients(i).BlockDrawing = False
                Clients(i).bSentHostDP = False
                Set Clients(i).IPicture = Nothing
                Clients(i).iPing = 0
                Clients(i).sHasiDPs = vbNullString
                Clients(i).sIP = vbNullString
                Clients(i).sName = vbNullString
                Clients(i).sVersion = vbNullString
                Exit For
            End If
            
            
        Next i
        
        'For i = 0 To UBound(Clients)
            'If Clients(i).iSocket = Index Then
        If bTell Then
            'modMessaging.DistributeMsg eCommands.Info & Msg & "0", Index
            SendInfoMessage Msg
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

Str = "Connection Closed"
IP = SckLC.RemoteHostIP

If LenB(CStr(IP)) And Not Server Then
    Str = Str & " - from " & IP
End If

Str = Str & " (" & CStr(Time) & ")"

AddConsoleText Str
If modVars.bRetryConnection = False Then
    AddText Str, , True, True
    
    If (Not modVars.IsForegroundWindow()) And (Not Closing) Then
        If Server Then
            Msg = "All Connections Closed - " & CStr(Time)
        Else
            Msg = "Disconnected from Server - " & CStr(Time)
        End If
        
        frmSystray.ShowBalloonTip Msg, , NIIF_INFO
    End If
    
End If

If SendTypeTrue Then
    txtOut.Text = vbNullString
    DoEvents
End If

SckLC.Close  'close connection


If modVars.bRetryConnection = False Then
    Call CleanUp(True)
End If

'Cmds Idle

End Sub

Private Sub SckLC_Connect()
Dim TimeTaken As Long
Dim Text As String

'txtLog is the textbox used as our
'chat buffer.

'SckLC.RemoteHost returns the hostname( or ip ) of the host
'SckLC.RemoteHostIP returns the IP of the host

TimeTaken = GetTickCount() - ConnectStartTime

Text = "Connected to " & SckLC.RemoteHostIP & " in " & _
    CStr(TimeTaken / 1000) & _
    " seconds (" & FormatDateTime$(Time$, vbLongTime) & ")"

AddText Text, , True, True
AddConsoleText Text
Cmds Connected

Pause 25
Text = FormatApostrophe(LastName) & " Version: " & GetVersion()

'SendData eCommands.Info & Text & "0"
SendInfoMessage Text

If LenB(Inviter) > 0 Then
    Pause 25
    Text = LastName & " was invited by " & Trim$(Inviter)
    'SendData eCommands.Info & Text
    SendInfoMessage Text
End If

modVars.bRetryConnection = False

If frmMain.Visible = False Then
    frmSystray.ShowBalloonTip "Communicator has connected - " & SckLC.RemoteHostIP & ":" & CStr(RPort), , NIIF_INFO, , True
    ShowForm
End If

On Error Resume Next
txtOut.SetFocus
End Sub

Public Sub pDataArrival(ByRef Sck As Winsock, ByRef Index As Integer, ByRef bytesTotal As Long)

'Static AlreadyProcessing As Boolean
'Static LastData As String
Dim Dat As String, i As Integer ', L As Integer
Dim Dats() As String
Static LastDat As String

'"Message1" & modMessaging.MessageSeperator & "Messa"

If Status = Connected Then
    
    On Error Resume Next
    Sck.GetData Dat, vbString, bytesTotal 'writes the new data in our string dat (string format)
    
    Dats = Split(Dat, modMessaging.MessageSeperator)
    
    
    If LenB(LastDat) Then
        If Left$(Dats(0), 1) <> modMessaging.MessageStart Then
            Dats(0) = LastDat & Dats(0)
            LastDat = vbNullString
        End If
    End If
    
    
    'go through all except last (incase it's truncated)
    For i = LBound(Dats) To UBound(Dats) - 1
        
        If LenB(Dats(i)) Then
            Call DataArrival(Mid$(Dats(i), 2), Index)
        End If
        
    Next i
    
    
    'If UBound(Dats) = 0 Then
        'Call DataArrival(Mid$(Dats(0), 2), Index)
        'never gets here
        
    If LenB(Dats(UBound(Dats))) Then
        
        If Left$(Dats(UBound(Dats)), 1) = modMessaging.MessageStart Then
            'otherwise, remember the last one (truncated)
            LastDat = Dats(UBound(Dats))
            
            'Debug.Print "Lastdat used: " & LastDat
            
        End If
        
    ElseIf LenB(LastDat) Then
        LastDat = vbNullString
    End If
    
End If


End Sub

Private Sub SckLC_DataArrival(ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

pDataArrival SckLC, 0, bytesTotal

End Sub

Public Sub SckLC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim sTxt As String

If Number = modVars.WSAHOSTNOTFOUND Or Number = modVars.WSANO_DATA Then
    sTxt = "Their computer is Shut Down/In Standby"
    
ElseIf Number = modVars.WSAEADDRINUSE Then 'addr in use
    Call ErrorHandler("Address In Use", Number) ', False, True)
    
    If modVars.bRetryConnection Then
        sTxt = "Error - Can't Initialise Connection. Auto-Connect turned off"
        modVars.bRetryConnection = False
    End If
    
ElseIf Number = modVars.WSAECONNREFUSED Then 'connection is forcefully rejected
    sTxt = "Could not establish connection - Their Communicator may not be listening"
    
ElseIf Number = modVars.CustomLagError Then
    sTxt = "Error: Connection Timed Out"
    
Else
    sTxt = "Error: " & Description
    
    '10060 = timeout
End If


If LenB(sTxt) Then
    If modVars.bRetryConnection = False Then
        AddText sTxt, TxtError, True
    Else
        AutoConnect True
    End If
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

Dim SocketToUse As Integer, i As Integer, MinsUp As Single
Dim Txt As String, IP As String
Dim SystrayTxt As String, UpTimeText As String

'client initilisation
Dim bInit As Boolean


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
                AddText "Blocked IP (" & IP & ") attempted to connect - Rejected", TxtError, True
            End If
            
            'SendDevCmd edevcmds.Visible,
            
            'SendData eCommands.Info & "You have been Kicked - Your IP is blocked1", SocketToUse
            SendInfoMessage "You have been Kicked - Your IP is blocked", True, , SocketToUse
            
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
AddConsoleText Txt   '& " SocketHandle: " & SockAr(SocketToUse).SocketHandle

SetMiniInfo SystrayTxt

'if server then modmessaging.DistributeMsg "Client
'SendData eCommands.GetName, SocketCounter


'If Server Then modMessaging.DistributeMsg eCommands.Info & Txt & "0", SocketToUse
SendInfoMessage Txt, , , , SocketToUse
'no point telling guy who's connected that he's connected

frmSystray.ShowBalloonTip "New Connection Established - " & SystrayTxt, "Communicator", NIIF_INFO

MinsUp = (GetTickCount() - modLoadProgram.LoadStart) / 60000
If MinsUp >= 60 Then
    MinsUp = MinsUp / 60
    UpTimeText = Round(MinsUp, 2) & " hour" & IIf(MinsUp > 1, "s", vbNullString)
Else
    UpTimeText = Round(MinsUp, 0) & " min" & IIf(MinsUp > 1, "s", vbNullString)
End If

'#########################################################################################################
'info, etc
modMessaging.SendSetSocketMessage SocketToUse

'modMessaging.SendData _
    eCommands.Info & "Welcome to " & LastName & "'" & IIf(Right$(LastName, 1) = "s", vbNullString, "s") & _
    " Server, Version: " & GetVersion() & _
    ". Server Up Time: " & MinsUp & " min" & IIf(MinsUp > 1, "s", vbNullString) _
    & "0", SocketToUse '                                           convert into minutes ^

SendInfoMessage "Welcome to " & FormatApostrophe(LastName) & _
    " Server, Version: " & GetVersion() & _
    ". Server Up Time: " & UpTimeText, , , SocketToUse

'If LenB(ServerMsg) Then modMessaging.SendData eCommands.Info & "Server Message: " & ServerMsg & "0", SocketToUse
If LenB(ServerMsg) Then SendInfoMessage "Server Message: " & ServerMsg, , , SocketToUse


If mnuFileGameMode.Checked Then
    'modMessaging.SendData eCommands.Info & sGameModeMessage & "1", SocketToUse
    SendInfoMessage sGameModeMessage, True, , SocketToUse
    
    If modSpeech.sGameSpeak Then
        modSpeech.Say "Communicator has a New Connection Established", , , True
    End If
    
End If
'#########################################################################################################


'add to clients
On Error GoTo NotInitResume

'bInit = True
'For i = 0 To UBound(Clients)
'    If Clients(i).iSocket = SocketToUse Then 'might have been left over...?
'        bInit = False
'        Exit For
'    End If
'Next i
bInit = (FindClient(SocketToUse) = -1)


If bInit Then
    'If UBound(Clients) = 0 Then
        ''first person to connect, add ourselves'''''''or not
        'ReDim Preserve Clients(UBound(Clients) + 2)
        'Clients(0).iSocket = -1
    'Else
    ReDim Preserve Clients(UBound(Clients) + 1) 'this'll add us anyway
    'End If
    
    Clients(UBound(Clients)).iSocket = SocketToUse
End If


NotInitResume:
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

tmrMain_Timer

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

Public Sub sockAr_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim i As Integer
Dim sTxt As String
Dim MyVer As String

If Number = modVars.CustomLagError Then
    AddText "Error: Connection to Client " & CStr(Index) & " Timed Out", TxtError, True
End If

MyVer = GetVersion()

For i = 1 To UBound(Clients)
    If Clients(i).iSocket = Index Then
        sTxt = " - " & Clients(i).sName
        
        If MyVer <> Clients(i).sVersion Then
            If LenB(Clients(i).sVersion) Then
                sTxt = sTxt & " [" & Clients(i).sVersion & "] "
            End If
        End If
        
        Exit For
    End If
Next i

sTxt = "Client " & Index & sTxt
sTxt = "Error (" & sTxt & "): " & Description

AddConsoleText "SockAr " & sTxt
'append the error message in the chat buffer
AddText sTxt, TxtError, True


'If Server Then modMessaging.DistributeMsg eCommands.Info & sTxt & "1", Index
'should be true, but...
'screw it
SendInfoMessage sTxt, True, , , Index


sockAr_Close Index

End Sub

Private Sub tmrHost_Timer()

Dim JustListened As Boolean
Dim Frm As Form

InActiveTmr = InActiveTmr + 1

'frmSystray.RefreshTray

If mnuOptionsHost.Checked Then
    If InActiveTmr >= 1 Then  '30 seconds
        If (Status <> Connected) And (Status <> Connecting) And modVars.bRetryConnection = False Then
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

'If mnuOptionsAdvInactive.Checked Then
'
'    If Status = Listening Then
'
'        If InActiveTmr >= 2 Then '1 min
'            For Each Frm In Forms
'                If Frm.Name <> Me.Name Then
'                    If Frm.Name <> "frmSystray" Then
'                        If Frm.Visible Then
'                            InActiveTmr = 0
'                            Exit Sub
'                        End If
'                    End If
'                End If
'            Next Frm
'
'            InActiveTmr = 0
'            'Call ClearRtfIn
'            If Me.Visible Then ShowForm False
'        End If
'
'    'Else
'        'InActiveTmr = 0
'    End If
''Else
'    'InActiveTmr = 0
'End If

End Sub

Public Sub SetInactive()
InActiveTmr = 0
End Sub

Private Sub AutoConnect(Optional bIsRetry As Boolean = False)
Dim sTxt As String
'Const lDelay = 60000 '1min
'Dim GTC As Long

'GTC = GetTickCount()

'If LastAutoRetry + lDelay < GTC Then
    'attempt to connect
    
Connect modVars.LastIP

sTxt = " to " & modVars.LastIP & ":" & RPort & "..."

If bIsRetry Then
    SetInfo "Retrying connection" & sTxt
Else
    SetInfo "Connecting" & sTxt
End If
    
    'LastAutoRetry = GTC
'End If

End Sub

Private Sub tmrMain_Timer()
Dim SendList As String, i As Integer
Dim addr As String, Str As String
Dim sSelected As String, sTxtToAdd As String
'Dim bHad As Boolean
Dim LBTopI As Integer
Dim tPic As IPictureDisp
Static LastIconRefresh As Long

If LastIconRefresh + IconRefreshDelay < GetTickCount() Then
    RefreshIcon
    rtfIn.ForceRefresh
    
    LastIconRefresh = GetTickCount()
End If

If Status <> Connected Then
    If modVars.bRetryConnection Then
        AutoConnect True
    End If
    Exit Sub
End If

If Me.mnuDevPause.Checked And bDevMode Then Exit Sub

'If Server Then
'    addr = SockAr(1).RemoteHostIP
'Else
'    addr = SckLC.RemoteHostIP
'End If

'Call DoPing(addr)

sSelected = lstConnected.Text
LBTopI = SendMessageByNum(lstConnected.hWnd, LB_GETTOPINDEX, 0, 0)

frmMain.lstConnected.Clear
On Error GoTo EH
For i = 0 To UBound(Clients)
    
    'If Clients(i).sName = LastName And Not bHad Then
        'bHad = True
    If LenB(Clients(i).sName) Then
        If Clients(i).iSocket <> modMessaging.MySocket And modMessaging.MySocket <> 0 Then
            With lstConnected
                
                sTxtToAdd = GetClientInfo(i)
                
                .AddItem sTxtToAdd
                .ItemData(.NewIndex) = Clients(i).iSocket
                
            End With
        End If
    End If
    
    
    
    If Not Server Then
        
        SendData eCommands.SetClientVar & eClientVarCmds.SetDPSet & IIf(modDP.DP_Path_Exists(), "1", "0")
        
        If LenB(Clients(i).sIP) = 0 Then
            If Clients(i).iSocket = -1 Then
                Clients(i).sIP = SckLC.RemoteHostIP
            End If
        End If
        
    Else
        
        If LenB(Clients(i).sIP) = 0 Then
            If Clients(i).iSocket > 0 Then
                'DistributeMsg CStr(eCommands.SetClientVar & eClientVarCmds.SetSocket & Clients(i).iSocket), -1
                
                If ControlExists(SockAr(Clients(i).iSocket)) Then
                    Clients(i).sIP = SockAr(Clients(i).iSocket).RemoteHostIP
                End If
                
            End If
        End If
        
        
        SendSetSocketMessage Clients(i).iSocket
        
    End If
    
    
    If modMessaging.MySocket <> 0 Then
        If Clients(i).iSocket = modMessaging.MySocket Then
            Clients(i).sIP = SckLC.LocalIP
            Clients(i).sStatus = LastStatus
        End If
    End If
    
    
    'picture
    Str = GetClientDPStr(i)
    
'    If FileExists(Str) Then
'        'Set Clients(i).IPicture = LoadPicture(Str)
'    ElseIf Not (Clients(i).IPicture Is Nothing) Then
'        'if deleted, save it
'        On Error Resume Next
'        SavePicture Clients(i).IPicture, Str
'    End If
    
    If Clients(i).IPicture Is Nothing Then
        If FileExists(Str) Then
            Set Clients(i).IPicture = LoadPicture(Str)
        End If
        
    Else
        'if deleted, save it
        
        If FileExists(Str) = False Then
            On Error Resume Next
            SavePicture Clients(i).IPicture, Str
        End If
    End If
    
    
    Call ShowDP(i)
Next i
EH:

If modMessaging.MySocket > 0 Then
'    For i = 0 To UBound(Clients)
'        If Clients(i).iSocket = modMessaging.MySocket Then
            
            i = FindClient(modMessaging.MySocket)
            
            If i > -1 Then
                If modDP.DP_Path_Exists() Then
                    Set Clients(i).IPicture = LoadPicture(modDP.My_DP_Path)
                    ShowDP i
                End If
            End If
            
            'Exit For
'        End If
'    Next i
End If


With lstConnected
    
    If .ListCount = 0 Then
        .AddItem ConnectedListPlaceHolder
        
    ElseIf LenB(sSelected) Then
        
        For i = 0 To .ListCount - 1
            If .List(i) = sSelected Then
                .ListIndex = i
                Exit For
            End If
        Next i
            
    ElseIf LBTopI > -1 Then
        SendMessageByNum lstConnected.hWnd, LB_SETTOPINDEX, LBTopI, 0
    End If
End With




If Server Then
    
    '######################
    'send clients list
    
    If Clients(0).sName <> LastName Then
        Clients(0).sName = LastName
        Clients(0).iSocket = -1
        Clients(0).sIP = IIf(LenB(lIP) = 0, SckLC.LocalIP, lIP)
    End If
    
    Clients(0).sVersion = GetVersion()
    
    '######################
    'send client list
    SendList = modMessaging.GetClientList()
    modMessaging.DistributeMsg eCommands.ClientList & SendList, -1
    '######################
    
    '######################
    'send game list refresh
    SendList = modSpaceGame.GetGames()
    modMessaging.DistributeMsg eCommands.LobbyCmd & eLobbyCmds.Refresh & SendList, -1
    '######################
    
    '######################
    'send typing+drawing list
    
    '               add myself
    SendList = IIf(LenB(txtOut.Text), LastName & "#", vbNullString) & modMessaging.GetTypingList()
    
    modMessaging.DistributeMsg eCommands.SetTyping & SendList, -1
    '######################
Else
    '######################
    'send my name
    modMessaging.SendData eCommands.SetClientVar & eClientVarCmds.SetName & LastName
    
    'send my info
    modMessaging.SendData eCommands.SetClientVar & eClientVarCmds.SetVersion & GetVersion()
    modMessaging.SendData eCommands.SetClientVar & eClientVarCmds.SetsStatus & LastStatus
    
End If

modDP.tmrMain_Timer
tmrPing_Timer



'clients have been updated, update picClient
On Error GoTo EHContinue

If LenB(picClient.Tag) Then
    i = CInt(picClient.Tag)
    
    ShowClientInfo i
End If



EHContinue:
'    lstConnected.Clear
'    cmdRemove.Enabled = False
'    '- no need, done in modmsging.ReceivedClientList
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

Private Sub tmrPing_Timer()
Dim i As Integer
Dim GTC As Long
Static LastPingBroadcast As Long
Const PingBroadcastDelay = 5000

If Server Then
    GTC = GetTickCount()
    For i = 0 To UBound(Clients)
        If Clients(i).iSocket = -1 Then
            If Clients(i).iPing = 0 Then
                Clients(i).iPing = 1
            End If
            
        ElseIf Clients(i).iSocket > 0 Then
            
            If Clients(i).lLastPing + PingBroadcastDelay < GTC Then
            
                
                SendData eCommands.PingCmd & ePingCmds.aPing, Clients(i).iSocket
                
                Clients(i).lPingStart = GTC
                Clients(i).lLastPing = GTC
                
            End If
            
            
        End If
    Next i
End If

End Sub

Public Function ControlExists(ByRef Ctrl As Control) As Boolean
Dim sName As String
On Error GoTo EH
sName = Ctrl.Name
ControlExists = True
EH:
End Function

Private Sub tmrLog_Timer()
Static iAuto As Integer

DoLog
LogPrivate

iAuto = iAuto + 1
If iAuto >= 2 Then '20s
    AutoSave
    
    iAuto = 0
End If

End Sub

Private Sub DoLog(Optional bForce As Boolean = False)
'10 sec interval

Static T As Integer
Dim LogPath As String ', FilePath As String


If Me.mnuOptionsMessagingLoggingConv.Checked Then
    
    If bForce Then
        T = 3
    Else
        T = T + 1
    End If
    
    If T >= 3 Then '30 seconds
        T = 0
        
        
        LogPath = GetLogPath() & MakeDateFile() & "\" 'AppPath() & "Logs\"
        'If FileExists(LogPath, vbDirectory) = False Then
            'MkDir LogPath
        'End If
        'LogPath = LogPath & MakeDateFile() & "\"
        If FileExists(LogPath, vbDirectory) = False Then
            MkDir LogPath
        End If
        
        
        If LenB(CurrentLogFile) = 0 Then
            CurrentLogFile = LogPath & MakeLogFileName() & ".rtf"
        End If
        
        
        'If Status = Connected Then
        On Error GoTo EH
        rtfIn.SaveFile CurrentLogFile, rtfRTF
        'End If
        
    End If
End If

Exit Sub
EH:
If Err.Number <> err_INVALIDORNOACCESS Then
    AddText "Log Error - " & Err.Description, TxtError, True
'else
    'they are viewing the log
End If
End Sub

Private Sub LogPrivate()
Dim RootPath As String, sPath As String, sName As String
Dim Frm As Form
Dim i As Integer

If modVars.nPrivateChats Then
    If mnuOptionsMessagingLoggingPrivate.Checked Then
        
        
        RootPath = GetLogPath() & MakeDateFile() & "\"
        If FileExists(RootPath, vbDirectory) = False Then
            MkDir RootPath
        End If
        
        
        For Each Frm In Forms
            If Frm.Name = frmPrivateName Then
                
                i = FindClient(Frm.SendToSock)
                
                If i > -1 Then
                    sName = Clients(i).sName
                Else
                    sName = "Randomer"
                End If
                    
                sPath = RootPath & sName & " Private.rtf"
                
                
                On Error GoTo EH
                Frm.rtfIn.SaveFile sPath, rtfRTF
                
            End If
        Next Frm
        
    End If
End If

Exit Sub
EH:
If Err.Number <> err_INVALIDORNOACCESS Then
    AddText "Log Error - " & Err.Description, TxtError, True
'else
    'they are viewing the log
End If
End Sub

Private Function MakeLogFileName() As String
MakeLogFileName = MakeTimeFile()
End Function

Private Function MakeTimeFile() As String
MakeTimeFile = Replace$(FormatDateTime$(Time$, vbLongTime), ":", ".")
'MakeTimeFile = Replace$(Time$, ":", ".")
'Replace$(Replace$(CStr(Date & " - " & Time), "/", ".", , , vbTextCompare), ":", ".", , , vbTextCompare)
End Function

Private Function MakeDateFile() As String
MakeDateFile = GetDate() 'Replace$(CStr(Date), "/", ".")
End Function

Private Function GetLogPath() As String
Dim sPath As String

sPath = AppPath() & "Logs\"

If FileExists(sPath, vbDirectory) = False Then
    On Error Resume Next
    MkDir sPath
End If

GetLogPath = sPath

End Function

Private Sub tmrShake_Timer()
Static Count As Integer

On Error Resume Next

If Me.Visible = False Then ShowForm

Count = Count + 1

If (Count Mod 2) = 1 Then
    'Me.Top = Me.Top + 100
    'Me.Left = Me.Left + 100
    Me.Move Me.Left + 100, Me.Top + 100
Else
    'Me.Top = Me.Top - 100
    'Me.Left = Me.Left - 100
    Me.Move Me.Left - 100, Me.Top - 100
End If

If Count > 5 Then
    tmrShake.Enabled = False
    Count = 0
ElseIf Count = 1 Then
    Beep
End If

End Sub

Public Sub tmrLP_Timer()
Dim Ans As VbMsgBoxResult
Dim DaysSinceUpdate As Integer
Dim Tmp As String

Const DayDiff = "d"


On Error GoTo EH
DaysSinceUpdate = DateDiff(DayDiff, LastUpdate, Date)


AddConsoleText "Days Since Last Update: " & CStr(DaysSinceUpdate)


If DaysSinceUpdate > 5 And mnuOptionsAdvAutoUpdate.Checked Then
    
    Ans = Question("You haven't checked for an update in the past five days, check now?", mnuFileExit)
    AddText "This can be turned off - Uncheck 'Options > Advanced > Remind me to check for updates'", , True
    
    If Ans = vbYes Then
        Call CheckForUpdates
    Else
        AddText "Check Canceled", , True
    End If
ElseIf bJustUpdated Then
    
    Tmp = GetSettingsPath()
    
    If FileExists(Tmp) Then
        Ans = MsgBoxEx("Settings File is old;" & vbNewLine & "Export Settings now, to overwrite the old file?", _
        "This is because the settings file structure has been changed, so the old file needs to be overwritten", _
        vbQuestion + vbYesNo, "Export Settings?", , , Me.Icon.Handle)
        
        If Ans = vbYes Then
            modSettings.ExportSettings Tmp
        End If
    End If
    
End If

EH:
End Sub

'Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
'
'Call Form_KeyDown(KeyCode, Shift)
'
'End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 172 Then KeyAscii = 0 'prevent ¬
End Sub

Private Sub txtName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    'If Status = Connected Then
    
    txtName.Enabled = False
    'Pause 1
    txtName.Enabled = True
    
    On Error Resume Next
    txtName.SetFocus
    
    PopupMenu mnuStatus
    'End If
End If

End Sub

Private Sub txtName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HideExtras False
On Error GoTo EH
If Screen.ActiveControl.Name = txtName.Name Then
    txtStatus.Visible = True
End If
EH:
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
    
    If Not (mnuDevDataCmdsTypeShow.Checked = False And bDevMode) Then
        If SendTypeTrue Then
            Msg = eCommands.Typing & "0" & LastName
            
            If Server Then
                DistributeMsg Msg, -1
            Else
                SendData Msg
            End If
            
            SendTypeTrue = False
        End If
    End If
    
Else
    cmdSend.Enabled = True
    cmdSend.Default = True
    
    'If Len(txtOut.Text) <= 1 Then
    txtName.Enabled = False
    'End If
    
    If Not (mnuDevDataCmdsTypeShow.Checked = False And bDevMode) Then
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
End If

'Call SetInactive
'txtOut.height = TextHeight(txtOut.Text) - 90

Msg = txtOut.Text
If LenB(Msg) Then
    If LCase$(Msg) = Left$("/describe", Len(Msg)) Then
        SetInfo "Press the Right Key to add '/Describe'"
    ElseIf LCase$(Msg) = Left$("/me", Len(Msg)) Then
        SetInfo "Press the Right Key to add '/Me'"
    End If
End If

End Sub

Public Sub RefreshNetwork(Optional ByRef LB As Control = Nothing, _
                          Optional ByRef CommentLB As Control = Nothing)
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

'If modVars.IsForegroundWindow() Then
    If NewLine Then
        picDraw.Line (X, Y)-(X, Y), Colour
        NewLine = False
    End If
    
    cx = X
    cy = Y
    
    SendLine X, Y, picDraw.DrawWidth
'End If

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
            
            pDrawDrawnOn = True
            
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
        
        txtName.Enabled = False
        
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
        
        txtName.Enabled = Not SendTypeTrue
        
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
            
            pDrawDrawnOn = True
            
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
        Else
            HideExtras
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
        'dColour = picDraw.Point(cx, cy)
        dColour = GetPixel(picDraw.hdc, ScaleX(cx, vbTwips, vbPixels), ScaleY(cy, vbTwips, vbPixels))
        
        If Colour <> -1 Then
            If dColour <> -1 Then
                picColour.BackColor = dColour
                Colour = dColour
                chkPickColour.Value = 0
            End If
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
Dim Txt As String
Dim bSetFocus As Boolean

If Shift And vbAltMask Then
    If Shift And vbShiftMask Then
        If Shift And vbCtrlMask Then
            
            If KeyCode = 191 Then
                'insert upside down ?
                txtOut.SelText = Chr$(KeyCode)
            End If
            
        End If
    End If
ElseIf KeyCode = vbKeyRight Then
    'auto complete /desc or /me
    
    Txt = txtOut.Text
    
    If LenB(Txt) Then
        If LCase$(Txt) = Left$("/describe", Len(Txt)) Then
            txtOut.Text = "/describe "
            bSetFocus = True
        ElseIf LCase$(Txt) = Left$("/me", Len(Txt)) Then
            txtOut.Text = "/me "
            bSetFocus = True
        End If
        
        
        If bSetFocus Then
            txtOut.Selstart = Len(txtOut.Text)
            On Error Resume Next
            txtOut.SetFocus
        End If
    End If
    
End If

End Sub

'Private Sub txtOut_KeyDown(KeyCode As Integer, Shift As Integer)
'Call Form_KeyDown(KeyCode, Shift)
'End Sub

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
'ElseIf KeyAscii = 10 Then
    'ctrl+enter
    
    'If txtOut.height = TextBoxHeight Then
        'txtOut.height = TextBoxHeight * 4
    'End If
    
'    If txtOut.Selstart = Len(txtOut.Text) Then
'        If Right$(txtOut.Text, 2) = vbNewLine Then
'            txtOut.height = txtOut.height - 195
'        End If
'    End If
    
    
End If


End Sub

Private Sub txtOut_GotFocus()
If mnuOptionsMessagingDisplayCompact.Checked Then
    txtOut.height = TextBoxHeightInc * TextBoxHeight
End If

HideExtras

End Sub

Private Sub txtOut_LostFocus()
Dim CtrlName As String

On Error GoTo EH
If Not (Screen.ActiveControl Is Nothing) Then
    
    CtrlName = Screen.ActiveControl.Name
    
    If Not ((CtrlName = "cmdSmile") Or (CtrlName = "cmdShake")) Then
        Call LostFocus(txtOut)
        ResetTxtOutHeight
    End If
    
    txtOut.Selstart = Len(txtOut.Text)
End If

EH:
End Sub

Public Sub ResetTxtOutHeight()
If mnuOptionsMessagingDisplayCompact.Checked Then
    txtOut.height = TextBoxHeight
End If
End Sub

Private Sub txtOut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    ''don't let the control display as "Greyed"
    'LockWindowUpdate txtOut.hWnd
    'disable the textbox, so that it can't react to mouse click
    txtOut.Enabled = False
    
    
    're-enable the control
    txtOut.Enabled = True
    
    'stop freezing the update
    'LockWindowUpdate 0&
    
    On Error Resume Next
    txtOut.SetFocus
    
    PopupMenu mnuFont, , , , mnuFontColour
    
End If

End Sub

'Private Sub mnuFileDelSettings_Click()
'
'Dim Str As String
'
'Call modSettings.DelSettings
'
'Str = "Settings Deleted"
'
'AddText Str, , True
'AddConsoleText Str
'
'mnuFileSaveExit.Checked = False
'End Sub

Private Sub txtName_LostFocus()
Rename txtName.Text
End Sub

Public Function FormatApostrophe(ByVal sName As String) As String
FormatApostrophe = sName & "'" & IIf(Right$(sName, 1) = "s", vbNullString, "s")
End Function

Public Sub Rename(ByVal sNewName As String)

Dim Msg As String, NameOp As String, sTxt As String
'aka Name to be operated on

NameOp = RemoveChars(sNewName)

If Trim$(NameOp) <> Trim$(sNewName) Then
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
        
        
        Msg = LastName & " renamed to " & NameOp
        
        SendInfoMessage Msg
        'If Server Then
            'DistributeMsg Msg, -1
        'Else
            'SendData Msg
            
            'Pause 1
            
        If Not Server Then
            'tell the server to set our name to the new one
            SendData eCommands.SetClientVar & eClientVarCmds.SetName & NameOp
            'don't call the timer - no need to refresh lstConnected etc etc
            
        End If
        
        sTxt = "Renamed to " & NameOp
        AddText sTxt, , True
        SetMiniInfo sTxt
        
    End If
    
End If

LastName = NameOp
txtName.Text = NameOp 'Trim$(txtName.Text)

'Call CheckAwayChecked

End Sub

Private Sub txtStatus_GotFocus()
txtStatus.Selstart = 0
txtStatus.Sellength = Len(txtStatus.Text)
End Sub

Private Sub txtStatus_LostFocus()
ReStatus txtStatus.Text
End Sub

Public Sub ReStatus(ByVal sNewStatus As String)
Dim i As Integer
Dim sNew As String, sToSend As String
Dim bTell As Boolean
Const StatusDefault = "[Status]"

sNew = RemoveChars(sNewStatus)

If sNew = StatusDefault Then
    sNew = vbNullString
End If

Do While TextWidth(sNew) > 1850
    sNew = Left$(sNew, Len(sNew) - 1)
    bTell = True
Loop

sNew = Trim$(sNew)

If LastStatus <> sNew Then
    If LenB(sNew) Then
        txtStatus.Text = sNew
        
        
        If bTell Then
            SetInfo "Status was too long - Shortened"
        End If
        
        SetMiniInfo "Set status to: " & sNew
        
        If Status = Connected Then
            sToSend = LastName & " set their status to '" & sNew & "'"
            
'            If Server Then
'                modMessaging.DistributeMsg eCommands.Info & sToSend, -1
'            Else
'                modMessaging.SendData eCommands.Info & sToSend
'            End If
            
            AddText "Status set to '" & sNew & "'", , True
        End If
        
    Else
        txtStatus.Text = StatusDefault
        
        SetMiniInfo "Status Removed"
        
        If Status = Connected Then
            
            sToSend = LastName & " removed their status"
'            If Server Then
'                modMessaging.DistributeMsg eCommands.Info & sToSend, -1
'            Else
'                modMessaging.SendData eCommands.Info & sToSend
'            End If
            
            AddText "Status Removed", , True
        End If
    End If
    
    If LenB(sToSend) Then 'should be true
        SendInfoMessage sToSend
    End If
    
    LastStatus = sNew
    'Stick(0).sStatus set on list timer
End If

End Sub

Private Function RemoveChars(Before As String) As String
Dim NameOp As String

NameOp = Trim$(Replace$(Before, "@", vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, "#", vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, modMessaging.MsgEncryptionFlag, vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, ":", vbNullString, , , vbTextCompare)) 'for chat in game
NameOp = Trim$(Replace$(NameOp, modSpaceGame.mPacketSep, vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, modSpaceGame.UpdatePacketSep, vbNullString, , , vbTextCompare))

RemoveChars = NameOp

End Function

Private Sub ucFileTransfer_Connected(IP As String)

If LenB(modDP.sPictureToSendPath) Then
    frmMain.ucFileTransfer.SendFile modDP.sPictureToSendPath, modDP.sRemoteFileName
    
    modDP.sPictureToSendPath = vbNullString
    modDP.sRemoteFileName = vbNullString
    
    
    If modDP.bSetHostBool Then
        On Error Resume Next
        Clients(modDP.DP_iClient).bSentHostDP = True
    Else
        On Error Resume Next
        Clients(modDP.DP_iClient).sHasiDPs = Clients(modDP.DP_iClient).sHasiDPs & "," & CStr(modDP.iToAdd)
    End If
    
    
    ucFileTransfer.Disconnect
    
    If Server Then
        ucFileTransfer.Listen
    End If
End If


End Sub

Private Sub ucFileTransfer_ReceivedFile(sFileName As String)
Dim iClient As Integer, iSock As Integer, i As Integer
Dim FNameOnly As String
Dim sSavePath As String

On Error GoTo EH

FNameOnly = Mid$(sFileName, InStrRev(sFileName, "\") + 1)
iSock = val(FNameOnly)

iClient = FindClient(iSock) '-1
'For i = 0 To UBound(Clients)
'    If Clients(i).iSocket = iSock Then
'        iClient = i
'        Exit For
'    End If
'Next i


If iClient > -1 Then
    Set Clients(iClient).IPicture = LoadPicture(sFileName)
    ShowDP iClient
    
    
    If mnuOptionsDPSaveAll.Checked Then
        'save pic
        On Error GoTo FCEH
        sSavePath = FT_Path() & "\Display Pictures"
        
        If FileExists(sSavePath, vbDirectory) = False Then
            MkDir sSavePath
        End If
        
        FileCopy sFileName, sSavePath & "\" & FormatApostrophe(Clients(iClient).sName) & " DP (" & MakeTimeFile() & ").jpg"
    End If
    
End If

If modDev.bDevDataFormLoaded Then
    frmDevData.RemoveFTCap
End If

EH:
Exit Sub
FCEH:
AddText "Error Saving Display Picture - " & Err.Description, TxtError, True
End Sub

Private Sub ucFileTransfer_ReceivingFile(sFileName As String, ByVal BytesReceived As Long, ByVal BytesRemaining As Long, ByVal lTotalBytes As Long)

If modDev.bDevDataFormLoaded Then
    On Error Resume Next
    frmDevData.SetFTCap FormatNumber$(100 * (lTotalBytes - BytesRemaining) / lTotalBytes, 1, vbTrue, vbFalse, vbFalse) & "%"
    
    'lblStatus.Caption = "Receiving - " & GetFileName(sFileName) & _
        " (" & FormatNumber$(BytesReceived / 1024, 2, vbTrue, vbFalse, vbFalse) & " KB) - " & _
        FormatNumber$(100 * BytesReceived / lTotalBytes, 0, vbTrue, vbFalse, vbFalse) & "%"
    
End If

End Sub

Private Sub ucFileTransfer_SendingFile(sFileName As String, ByVal BytesSent As Long, ByVal BytesRemaining As Long, ByVal lTotalBytes As Long)

If modDev.bDevDataFormLoaded Then
    On Error Resume Next
    frmDevData.SetFTCap FormatNumber$(100 * (lTotalBytes - BytesRemaining) / lTotalBytes, 2, vbTrue, vbFalse, vbFalse) & "%"
    
    
    'lblStatus.Caption = "Sending - " & GetFileName(sFileName) & _
        " - " & Format$(100 * (lTotalBytes - BytesRemaining) / lTotalBytes, "0.00") & "%"
End If

End Sub

Private Sub ucFileTransfer_Error(Description As String, ErrNo As eFTErrors)
If modDev.bDevDataFormLoaded Then
    On Error Resume Next
    frmDevData.SetFTCap "Error - " & Description
End If
End Sub

Private Sub ucFileTransfer_Diconnected()
If modDev.bDevDataFormLoaded Then
    On Error Resume Next
    frmDevData.SetFTCap "Disconnected"
End If
End Sub

Private Sub ucFileTransfer_SentFile(sFileName As String)
If modDev.bDevDataFormLoaded Then
    frmDevData.RemoveFTCap
End If
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
            'mnuOptionsWindow2Fade.Checked = True
            mnuOptionsWindow2Slide.Checked = True
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
Dim Name As String, IPTo As String, IPFrom As String
Dim Ans As VbMsgBoxResult
Dim AutoReject As Boolean

Dim i As Integer, j As Integer

i = InStr(1, Txt, "#", vbTextCompare)
j = InStr(1, Txt, "@", vbTextCompare)

Name = Left$(Txt, i - 1)
IPTo = Mid$(Txt, i + 1, j - i - 1)
IPFrom = Mid$(Txt, j + 1)
AutoReject = mnuOptionsMessagingDisplayIgnoreInvites.Checked

'confirm receiving of invite
frmUDP.SendToSingle IPFrom, frmUDP.UDPInfo & Me.LastName & " Received the Invite", False
frmUDP.UDPListen


If Me.Visible = False Then
    Me.ShowForm
End If


If Questioning Or AutoReject Then
    Ans = vbNo
    
    Txt = "Invite Ignored, from " & Name & " (" & IPFrom & ")"
    
    If AutoReject Then
        AddConsoleText Txt
        AddText Txt, , True
    End If
ElseIf Status = Connected Or Status = Connecting Then
    Ans = vbCancel
Else
    Ans = Question(Name & " (" & IPFrom & ") sends an invite, connect to " & IPTo & "?", frmUDP.cmdInvite)
End If

If Ans = vbYes Then
    Me.CleanUp True
    Connect IPTo
    Inviter = Name
    
ElseIf Ans = vbNo Then
    
    frmUDP.SendToSingle IPFrom, _
        frmUDP.UDPInfo & "Invite to " & Me.LastName & _
        Space$(1) & IIf(AutoReject And Not Questioning, "Auto-", vbNullString) & "Rejected" & _
        IIf(Questioning And Not AutoReject, " (Answering Another Question)", vbNullString), False
    
    
    frmUDP.UDPListen
    Inviter = vbNullString
    
ElseIf Ans = vbCancel Then
    
    frmUDP.SendToSingle IPFrom, _
        frmUDP.UDPInfo & "Invite to " & Me.LastName & _
        Space$(1) & "Rejected - Already Connected/Connecting", False
    
    
    frmUDP.UDPListen
    Inviter = vbNullString
    
ElseIf Ans = vbRetry Then
    'timeout
    
    frmUDP.SendToSingle IPFrom, _
        frmUDP.UDPInfo & "Invite to " & Me.LastName & _
        Space$(1) & "Rejected - Question Timed Out", False
    
    
    frmUDP.UDPListen
    Inviter = vbNullString
    
End If

End Sub

Public Sub SetIcon(ByVal St As eStatus)

pSetIcon CInt(St) + 1, Me.mnuFileGameMode.Checked, bDevMode, modDev.bUberDevMode

End Sub

Private Sub pSetIcon(ByVal iImg As Integer, _
    Optional pbGameMode As Boolean = False, _
    Optional pbDevMode As Boolean = False, _
    Optional pbUberDev As Boolean = False)

Dim lhWndTop As Long, lhWnd As Long, lHandle As Long

If pbGameMode Then
    imgStatus.Picture = frmSystray.imgGame.ListImages(iImg).Picture
ElseIf pbUberDev Then
    imgStatus.Picture = frmSystray.imgUberDev.ListImages(iImg).Picture
ElseIf pbDevMode Then
    imgStatus.Picture = frmSystray.imgDev.ListImages(iImg).Picture
Else
    imgStatus.Picture = frmSystray.img32x32.ListImages(iImg).Picture
End If


If pbUberDev Then
    lHandle = frmSystray.img16x16UberDev.ListImages(iImg).Picture.Handle
ElseIf pbDevMode Then
    lHandle = frmSystray.img16x16Dev.ListImages(iImg).Picture.Handle
Else
    lHandle = frmSystray.img16x16.ListImages(iImg).Picture.Handle
End If

'frmMain.Icon = frmSystray.img48x48.ListImages(iImg).Picture
SendMessageByNum frmMain.hWnd, WM_SETICON, ICON_SMALL, lHandle
frmSystray.IconHandle = lHandle


lhWnd = Me.hWnd
lhWndTop = lhWnd
Do While lhWnd > 0
    lhWnd = GetWindow(lhWnd, GW_OWNER)
    
    If lhWnd > 0 Then
        lhWndTop = lhWnd
    End If
Loop
SendMessageByNum lhWndTop, WM_SETICON, ICON_BIG, imgStatus.Picture.Handle
'SendMessageByNum lhWndTop, WM_SETICON, ICON_SMALL, imgStatus.Picture.Handle

End Sub

'these are for FormLoad - Loading sub forms
Public Property Get ConnectedIcon() As IPictureDisp
Set ConnectedIcon = frmSystray.img32x32.ListImages(2).Picture
End Property
Public Property Get IdleIcon() As IPictureDisp
Set IdleIcon = frmSystray.img32x32.ListImages(1).Picture
End Property

Public Sub RefreshIcon()

'e.g. switched into dev mode
Call SetIcon(modVars.Status)

End Sub

'Private Sub tmrAnim_Timer()
'Static i As Integer
'Static bForward As Boolean
'
'If modStickGame.StickFormLoaded = False Then
'    If modSpaceGame.GameFormLoaded = False Then
'
'    i = i + IIf(bForward, 1, -1)
'
'        If i > 7 Then
'            i = 7
'            bForward = False
'        ElseIf i < 1 Then
'            i = 1
'            bForward = True
'        End If
'
'        imgStatus.Picture = imgAnim.ListImages(i).Picture
'    End If
'End If
'
'End Sub

Private Sub mnuOptionsMessagingWindowsWebcam_Click()
'Unload frmCapture
'Load frmCapture
'frmCapture.Show vbModeless, Me
End Sub

Private Sub rtfIn_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Files() As String
Dim i As Integer

ReDim Files(Data.Files.Count - 1)

For i = 0 To Data.Files.Count - 1
    Files(i) = Data.Files(i + 1)
Next i

DragDrop Files

End Sub

Private Sub rtfIn_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
SetInfo "Drag over the textbox to open file transfer"
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim Files() As String
'Dim i As Integer
'
'ReDim Files(Data.Files.Count - 1)
'
'For i = 0 To Data.Files.Count - 1
'    Files(i) = Data.Files(i + 1)
'Next i
'
'DragDrop Files

imgStatus_OLEDragDrop Data, Effect, Button, Shift, X, Y

End Sub

Private Sub DragDrop(Files() As String)

If Status = Connected Then
    If UBound(Files) >= 1 Then
        SetInfo "You can only send one file at once"
    Else
        mnuOptionsMessagingWindowsFT_Click
        frmManualFT.FilePath = Files(0)
    End If
End If

End Sub

Public Sub DP_OLEDragDrop(ByVal bEn As Boolean)
Dim i As Integer, j As Integer
i = Abs(bEn)

imgStatus.OLEDropMode = i
For j = 0 To imgDP.UBound
    imgDP(j).OLEDropMode = i
Next j
txtName.OLEDropMode = i

End Sub

Private Function IsDP_File(ByVal sPath As String) As Boolean

Select Case Right$(sPath, 4)
    Case ".jpg", ".bmp", ".jpeg"
        IsDP_File = True
    Case Else
        IsDP_File = False
End Select

End Function


Private Sub imgStatus_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Status = Connected Then
    If Data.Files.Count > 1 Then
        SetInfo "You can only have one Display Picture..."
        
    ElseIf IsDP_File(Data.Files(1)) Then
        SetMyDP Data.Files(1)
        
    Else
        SetInfo "That 'picture' isn't in the correct format"
        
    End If
End If

End Sub

Private Sub txtName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStatus_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub imgDP_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStatus_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
'########################################
Private Sub imgStatus_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
SetInfo "Drag Over to Set Display Picture"
End Sub

Private Sub txtName_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
imgStatus_OLEDragOver Data, Effect, Button, Shift, X, Y, State
End Sub

Private Sub imgDP_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
imgStatus_OLEDragOver Data, Effect, Button, Shift, X, Y, State
End Sub

