VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Communicator"
   ClientHeight    =   9180
   ClientLeft      =   75
   ClientTop       =   765
   ClientWidth     =   9795
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   9795
   Begin projMulti.ucFileTransfer ucVoiceTransfer 
      Left            =   8880
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox picVoice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4200
      ScaleHeight     =   705
      ScaleWidth      =   3585
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Timer tmrVoice 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1800
         Top             =   240
      End
      Begin projMulti.VistaProg progVoice 
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   397
      End
      Begin VB.Label lblVoice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Recording... (F2 to Send, F3 to Cancel)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdAprilFoolReset 
      Caption         =   "Back to Normal"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Text            =   "[Status]"
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "&No"
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   61
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmdReply 
      Caption         =   "&Yes"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   62
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Private Chat"
      Height          =   375
      Index           =   5
      Left            =   1680
      TabIndex        =   20
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Scan"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Kick"
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   18
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Connect"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Disconnect"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Listen"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin projMulti.XPButton cmdXPReply 
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   60
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Yes"
      CapAlign        =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXPReply 
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   59
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&No"
      CapAlign        =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      cBack           =   -2147483633
   End
   Begin VB.Timer tmrHost 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   600
      Top             =   720
   End
   Begin projMulti.XPButton cmdXPArray 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Scan"
      CapAlign        =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXPArray 
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   16
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Remove"
      CapAlign        =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXPArray 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Connect"
      CapAlign        =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXPArray 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Disconnect"
      CapAlign        =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBack           =   -2147483633
   End
   Begin projMulti.XPButton cmdXPArray 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Listen"
      CapAlign        =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBack           =   -2147483633
   End
   Begin projMulti.ucInactiveTimer ucInactiveTimer 
      Left            =   8280
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      InactiveInterval=   60000
   End
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
      HideSelection   =   0   'False
      Left            =   3360
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   54
      Top             =   3000
      Width           =   4575
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   5520
      ScaleHeight     =   1905
      ScaleWidth      =   3585
      TabIndex        =   66
      Top             =   6600
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
      TabIndex        =   58
      Top             =   3480
      Width           =   255
   End
   Begin VB.PictureBox picBig 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   5040
      ScaleHeight     =   2625
      ScaleWidth      =   3225
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
      Begin SHDocVwCtl.WebBrowser wbDP 
         CausesValidation=   0   'False
         Height          =   1335
         Left            =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         ExtentX         =   2566
         ExtentY         =   2355
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Image imgBig 
         Height          =   285
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   285
      End
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
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   5055
   End
   Begin projMulti.ucFileTransfer ucFileTransfer 
      Left            =   8160
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer tmrMain 
      Interval        =   10000
      Left            =   2880
      Top             =   1440
   End
   Begin VB.ListBox lstComputers 
      Height          =   840
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ListBox lstConnected 
      Height          =   840
      Left            =   1680
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame fraTyping 
      Height          =   615
      Left            =   120
      TabIndex        =   63
      Top             =   5040
      Width           =   3195
      Begin VB.Label lblTyping 
         Alignment       =   2  'Center
         Caption         =   "Drawing Label"
         Height          =   435
         Left            =   120
         TabIndex        =   64
         Top             =   120
         Width           =   2955
      End
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
      Height          =   1935
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Width           =   3195
      Begin VB.PictureBox picClear 
         Height          =   700
         Left            =   120
         ScaleHeight     =   645
         ScaleWidth      =   2955
         TabIndex        =   31
         Top             =   600
         Width           =   3015
         Begin VB.OptionButton optnDraw 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optnDraw 
            Caption         =   "Straight Line"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   33
            Top             =   0
            Width           =   1215
         End
         Begin VB.OptionButton optnDraw 
            Caption         =   "Pick Colour"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   34
            Top             =   310
            Width           =   1215
         End
         Begin VB.CommandButton cmdNormCls 
            Caption         =   "Clear"
            Height          =   375
            Left            =   1560
            TabIndex        =   35
            Top             =   250
            Width           =   1215
         End
         Begin projMulti.XPButton cmdXPCls 
            Height          =   255
            Left            =   1560
            TabIndex        =   36
            Top             =   360
            Width           =   1215
            _ExtentX        =   1720
            _ExtentY        =   450
            Caption         =   "Clear"
            CapAlign        =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            cBack           =   -2147483633
         End
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H000000C0&
         Height          =   255
         Index           =   8
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00004080&
         Height          =   255
         Index           =   9
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H0000C0C0&
         Height          =   255
         Index           =   10
         Left            =   600
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00008000&
         Height          =   255
         Index           =   11
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00FF0000&
         Height          =   255
         Index           =   12
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00800080&
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   13
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   14
         Left            =   1560
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   1920
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox picColour 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.ComboBox cboWidth 
         Height          =   315
         Left            =   600
         TabIndex        =   27
         Text            =   "1"
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cboRubber 
         Height          =   315
         Left            =   2160
         TabIndex        =   29
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   600
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H0000FF00&
         Height          =   255
         Index           =   3
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Index           =   4
         Left            =   1080
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00808080&
         Height          =   255
         Index           =   5
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox picColours 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   1560
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblColour 
         Caption         =   "Colour   Last Colour"
         Height          =   405
         Left            =   2280
         TabIndex        =   45
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label lblDraw 
         Caption         =   "Draw:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblRubber 
         Caption         =   "Rubber:"
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSmile 
      Height          =   375
      Left            =   9000
      Picture         =   "frmMain.frx":636A
      Style           =   1  'Graphical
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3000
      Width           =   285
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   67
      Top             =   8925
      Width           =   9795
      _ExtentX        =   17277
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
            Object.Width           =   11642
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
      TabIndex        =   65
      Top             =   5760
      Width           =   7215
   End
   Begin VB.Timer tmrLog 
      Interval        =   10000
      Left            =   3600
      Top             =   2160
   End
   Begin VB.Timer tmrShake 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8280
      Top             =   3720
   End
   Begin VB.CommandButton cmdShake 
      Caption         =   "Shake"
      Height          =   375
      Left            =   8160
      TabIndex        =   57
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   600
      MaxLength       =   20
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   8160
      TabIndex        =   55
      Top             =   3000
      Width           =   735
   End
   Begin projMulti.smRtfFBox rtfIn 
      Height          =   2400
      Left            =   3360
      TabIndex        =   25
      Top             =   600
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   4233
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HideSelection   =   0   'False
      hWnd            =   295113
      MouseIcon       =   "frmMain.frx":671C
      Text            =   "rtfIn"
      EnableTextFilter=   -1  'True
      SelRtf          =   $"frmMain.frx":6738
   End
   Begin projMulti.XPButton cmdXPArray 
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   22
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Private Chat"
      CapAlign        =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBack           =   -2147483633
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
      End
      Begin VB.Menu mnuFileOpenFolder 
         Caption         =   "Open Communicator Folder"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMini 
         Caption         =   "Show Mini Window"
      End
      Begin VB.Menu mnuFileThumb 
         Caption         =   "Show Thumbnail Window"
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
      End
      Begin VB.Menu mnuFileConnection 
         Caption         =   "Connection/IPs"
         Begin VB.Menu mnuFileClient 
            Caption         =   "Client Window..."
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
            Shortcut        =   ^K
         End
         Begin VB.Menu mnuFileConnectionSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileIPs 
            Caption         =   "View IPs/Who's Online..."
         End
         Begin VB.Menu mnuFileNetIPs 
            Caption         =   "View Network IPs..."
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
         Begin VB.Menu mnuOptionsDPClear 
            Caption         =   "Clear Display Picture"
         End
         Begin VB.Menu mnuOptionsDPSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsDPReset 
            Caption         =   "Reload All Pictures"
         End
         Begin VB.Menu mnuOptionsDPSaveAll 
            Caption         =   "Save All Pictures (Received Files)"
         End
      End
      Begin VB.Menu mnuOptionsAlerts 
         Caption         =   "Alerts"
         Begin VB.Menu mnuOptionsAlertsStyle 
            Caption         =   "Balloon Tips"
            Index           =   0
         End
         Begin VB.Menu mnuOptionsAlertsStyle 
            Caption         =   "MSN-6 Style"
            Index           =   1
         End
         Begin VB.Menu mnuOptionsAlertsStyle 
            Caption         =   "GMail Style"
            Index           =   2
         End
         Begin VB.Menu mnuOptionsAlertsStyle 
            Caption         =   "Flat Style"
            Index           =   3
         End
      End
      Begin VB.Menu mnuOptionsWindow2 
         Caption         =   "Window Display"
         Begin VB.Menu mnuOptionsAdvDisplayConn 
            Caption         =   "Connect Button Function"
            Begin VB.Menu mnuOptionsAdvDisplayConnF 
               Caption         =   "Manual Connect"
               Index           =   0
            End
            Begin VB.Menu mnuOptionsAdvDisplayConnF 
               Caption         =   "Connect to Selected"
               Index           =   1
            End
         End
         Begin VB.Menu mnuOptionsWindow2Effects 
            Caption         =   "Visual Effects"
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
         Begin VB.Menu mnuOptionsWindow2Animation 
            Caption         =   "Window Animation"
         End
         Begin VB.Menu mnuOptionsWindow2Sep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsWindow2SingleClick 
            Caption         =   "Single Click Tray Icon"
         End
      End
      Begin VB.Menu mnuOptionsMessagingDisplay 
         Caption         =   "Display"
         Begin VB.Menu mnuOptionsMessagingDisplaySmilies 
            Caption         =   "Smilies"
            Begin VB.Menu mnuOptionsMessagingDisplaySmiliesComm 
               Caption         =   "Communicator Smilies"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuOptionsMessagingDisplaySmiliesMSN 
               Caption         =   "MSN Smilies"
            End
         End
         Begin VB.Menu mnuOptionsMessagingDisplaySysUserName 
            Caption         =   "Use System User Name as Default"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsMessagingDisplayNewLine 
            Caption         =   "Separate Messages from Sender's Name"
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
         Begin VB.Menu mnuOptionsMessagingShake 
            Caption         =   "Allow Shaking"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsMessagingHurgh 
            Caption         =   "Use shake sound effect"
         End
         Begin VB.Menu mnuOptionsMessagingDisplayCompact 
            Caption         =   "Compact Typing Box"
         End
         Begin VB.Menu mnuOptionsFlashMsg 
            Caption         =   "Flash When Message Recieved"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsMessagingSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsMessagingDisplayShowHost 
            Caption         =   "Show Client's Host Names"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsMessagingDisplayIgnoreInvites 
            Caption         =   "Ignore All Invites (Auto-Reject)"
         End
         Begin VB.Menu mnuOptionsMessagingDisplayShowBlocked 
            Caption         =   "Show if a blocked IP connects"
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
               Visible         =   0   'False
            End
            Begin VB.Menu mnuOptionsMessagingWindowsFT 
               Caption         =   "Manual File Transfer..."
               Shortcut        =   ^T
            End
            Begin VB.Menu mnuOptionsMessagingWindowsSep 
               Caption         =   "-"
            End
            Begin VB.Menu mnuOptionsMessagingWindowsRecord 
               Caption         =   "Make Recording"
               Shortcut        =   {F2}
            End
            Begin VB.Menu mnuOptionsMessagingWindowsRecordCancel 
               Caption         =   "Cancel Recording/Show transfers"
               Shortcut        =   {F3}
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
            Begin VB.Menu mnuOptionsMessagingLoggingActivity 
               Caption         =   "Activity Log"
            End
         End
         Begin VB.Menu mnuOptionsMessagingCharMap 
            Caption         =   "Character Map..."
         End
         Begin VB.Menu mnuOptionsMessagingSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsMessagingVoicePreview 
            Caption         =   "Preview Voice"
            Shortcut        =   {F4}
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
         Begin VB.Menu mnuOptionsAdvNetworkRefresh 
            Caption         =   "Refresh Network-List"
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
         Begin VB.Menu mnuOptionsAdvInactive 
            Caption         =   "Inactive Timer"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsHost 
            Caption         =   "Host Mode (Auto-Listen)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptionsAdvHostMin 
            Caption         =   "Minimize to Tray When Hosting"
         End
         Begin VB.Menu mnuOptionsAdvSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOptionsStartup 
            Caption         =   "Run at System Startup"
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
      End
   End
   Begin VB.Menu mnuOnline 
      Caption         =   "Online"
      Begin VB.Menu mnuOnlineFTP 
         Caption         =   "FTP Settings"
         Begin VB.Menu mnuOnlineFTPDL 
            Caption         =   "FTP Download Settings"
            Begin VB.Menu mnuOnlineFTPDLO 
               Caption         =   "Use HTTP Download"
               Index           =   0
               Visible         =   0   'False
            End
            Begin VB.Menu mnuOnlineFTPDLO 
               Caption         =   "Use FTP Manual Download"
               Checked         =   -1  'True
               Index           =   1
            End
            Begin VB.Menu mnuOnlineFTPDLO 
               Caption         =   "Use Automatic FTP Download"
               Index           =   2
            End
         End
         Begin VB.Menu mnuOnlineFTPUL 
            Caption         =   "FTP Upload Settings"
            Begin VB.Menu mnuOnlineFTPULO 
               Caption         =   "Use FTP Manual Upload"
               Checked         =   -1  'True
               Index           =   1
            End
            Begin VB.Menu mnuOnlineFTPULO 
               Caption         =   "Use FTP Automatic Upload"
               Index           =   2
            End
         End
         Begin VB.Menu mnuOnlineFTPServer 
            Caption         =   "FTP Server"
            Begin VB.Menu mnuOnlineFTPServerAr 
               Caption         =   "Server Name etc Here"
               Index           =   0
            End
            Begin VB.Menu mnuOnlineFTPServerSep 
               Caption         =   "-"
            End
            Begin VB.Menu mnuOnlineFTPServerCustom 
               Caption         =   "Use Custom Server"
            End
            Begin VB.Menu mnuOnlineFTPServerView 
               Caption         =   "View Stored Servers"
            End
         End
         Begin VB.Menu mnuOnlineFTPSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOnlineFTPPassive 
            Caption         =   "Use Passive FTP semantics"
         End
         Begin VB.Menu mnuOnlineFTPServerMsgNow 
            Caption         =   "Download Server Message now"
         End
         Begin VB.Menu mnuOnlineFTPServerMsg 
            Caption         =   "Download Server Message when loaded"
         End
      End
      Begin VB.Menu mnuonlinesep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnlineUpdates 
         Caption         =   "Check for Updates"
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
      Begin VB.Menu mnuOnlineIPs 
         Caption         =   "View IPs/Who's Online..."
      End
      Begin VB.Menu mnuOnlinePortForwarding 
         Caption         =   "Port Forwarding..."
      End
      Begin VB.Menu mnuonlinesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnlineFileTransfer 
         Caption         =   "File Upload/Download..."
      End
      Begin VB.Menu mnuOnlineMessages 
         Caption         =   "Messages..."
      End
   End
   Begin VB.Menu mnuDev 
      Caption         =   "DevMode"
      Begin VB.Menu mnuDevForms 
         Caption         =   "Windows"
         Begin VB.Menu mnuDevForm 
            Caption         =   "Dev Main..."
         End
         Begin VB.Menu mnuDevFormsCmds 
            Caption         =   "Dev Command Window..."
            Shortcut        =   ^D
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
         Begin VB.Menu mnuDevShowCmds 
            Caption         =   "Show Recieved Dev Commands"
         End
         Begin VB.Menu mnuDevDataCmdsSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDevDataCmdsBlock 
            Caption         =   "Block Commands from X"
         End
         Begin VB.Menu mnuDevDataCmdsSetBlockMessage 
            Caption         =   "Set Block Message..."
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
         Begin VB.Menu mnuDevAdvCmdsNoFTPCallbacks 
            Caption         =   "FTP - Force Callbacks Off"
         End
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
            Shortcut        =   ^B
         End
      End
      Begin VB.Menu mnuDevSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDevChange 
         Caption         =   "Lower DevMode/Turn Off"
         Begin VB.Menu mnuDevChangeAr 
            Caption         =   "Turn Off DevMode"
            Index           =   0
         End
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
      End
      Begin VB.Menu mnuHelpSysTime 
         Caption         =   "Check System Time"
      End
      Begin VB.Menu mnuHelpBug 
         Caption         =   "Bug Report..."
      End
      Begin VB.Menu mnuHelpBugReport 
         Caption         =   "Bug Report Message"
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
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#############################################################################
' slash commands
Private Type irc_Command
    sCommand As String
    bChatMessage As Boolean
End Type
Private irc_Commands() As irc_Command


'#############################################################################
'window monitoring
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private GameWindowhWnd As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'#############################################################################
'voice recording

Private pbRecording As Boolean, bRecordingTransferInProgress As Boolean
Private time_Recorded As Long
Private recording_File_To_Send As String, recording_Remote_Filename As String
Private Const recording_Name As String = "Communicator_Voice", RecordingsViewString = " (F4 to view)"
Private Const recording_Time As Long = 10000 '10 secs

'#############################################################################

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long 'ctrl checking for ctrl+enter

Private bHadNetworkRefreshError As Boolean

Private Const DefaultPasswordMaxLen = 50
Private ServerMsg As String

Private Const TextBoxHeight = 285, TextBoxHeightInc = 3
Private Const NewLineLimit = 16

Private Const AFKStr As String = "AFK", logIndent As String = "    "

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


'menu popup
Private WithEvents HTML_Doc As HTMLDocument
Attribute HTML_Doc.VB_VarHelpID = -1

'Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Const ksAway = " (AFK)", sqB_Status_sqB = "[Status]"
Private Const ListeningStr = "Awaiting Connection..."

'font stuff
Private prtfFontName As String
Private prtfFontSize As Single
Private prtfBold As Boolean, prtfItalic As Boolean

'setting main icon (taskbar + alt tab)
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4


'logging
Private currentLogFile As String, logBasePath As String
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
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
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
Private lInactiveInterval As Long


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

'#######################################################################################
'used in getdesktoppath()
Private Declare Function SHGetSpecialFolderLocation _
    Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As _
    Long, pidl As ITEMIDLIST) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath _
    As String) As Long

Private Type SH_ITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SH_ITEMID
End Type

'################################################################################################

Private Property Let bRecording(bVal As Boolean)

picVoice.Visible = bVal
tmrVoice.Enabled = bVal
mnuOptionsMessagingWindowsRecordCancel.Enabled = bVal

If bVal Then
    'reset, since we're starting
    time_Recorded = 0
    progVoice.Value = 0
End If

pbRecording = bVal

End Property

Private Property Get bRecording() As Boolean
bRecording = pbRecording
End Property

Private Sub mnuDevChangeAr_Click(Index As Integer)
If Index = 0 Then
    modDev.setDevLevel modDev.Dev_Level_None, vbNullString
    modDev.AddDevText "DevMode Deactivated", True
Else
    'level is (index-1)
    modDev.setDevLevel Index, vbNullString
    addDevActivatedText
End If
End Sub

Private Sub mnuDevDataCmdsBlock_Click()

If mnuDevDataCmdsBlock.Checked Then
    mnuDevDataCmdsBlock.Checked = False
    
ElseIf modDev.blockCommands( _
    modVars.Password("Enter the password to block commands", Me, "Dev Block", , , 30) _
    ) Then
    
    SetInfo "Dev Commands Blocked", False
Else
    SetInfo "Incorrect Password", True
End If

End Sub

Private Sub mnuOnlineFTPServerMsgNow_Click()
DownloadServerMessage False
End Sub

Private Sub mnuOnlineFTPServerView_Click()
Unload frmFTPServers
Load frmFTPServers
frmFTPServers.Show vbModeless, Me
End Sub

Private Sub mnuOptionsMessagingHurgh_Click()
mnuOptionsMessagingHurgh.Checked = Not mnuOptionsMessagingHurgh.Checked
modSpeech.bHurgh = mnuOptionsMessagingHurgh.Checked
End Sub

Private Sub mnuOptionsMessagingLoggingActivity_Click()
mnuOptionsMessagingLoggingActivity.Checked = Not mnuOptionsMessagingLoggingActivity.Checked
End Sub

Private Sub mnuOptionsMessagingVoicePreview_Click()
Me.mnuCommandsTestSpeech_Click
End Sub

Private Sub mnuOptionsMessagingWindowsRecordCancel_Click()

If bRecording Then
    bRecording = False
    
    modSoundCap.Stop_Recording recording_Name
    modSoundCap.Close_Record recording_Name
    
    SetInfo "Recording Canceled", False
Else
    'otherwise, show the transfer window
    frmVoiceTransfers.ShowForm Not frmVoiceTransfers.Visible
End If

End Sub

Private Sub mnuOptionsMessagingWindowsRecord_Click()
Dim sTxt As String, localFileName As String, remoteFileName As String
Dim IP As String
Const ConnectTimeOut As Long = 1000
Dim Tick As Long
Dim iSocket As Integer, i As Integer

If bRecordingTransferInProgress Then
    SetInfo "Recording transfer in progress", True
    Exit Sub
End If


'do some voice shiz
If Not bRecording Then
    If modSoundCap.Open_Record(recording_Name) Then
        If modSoundCap.Start_Recording(recording_Name) Then
            
            bRecording = True
            Exit Sub
            
        Else
            sTxt = "Error Starting Recording - " & modSoundCap.GetLastErrorText()
        End If
    Else
        sTxt = "Error Initialising Recording - " & modSoundCap.GetLastErrorText()
    End If
    
    'only if error
    SetInfo sTxt, True
    AddConsoleText sTxt
    modSoundCap.Close_All
    Exit Sub
    
ElseIf modSoundCap.Stop_Recording(recording_Name) Then
    
    localFileName = GetCurrentLogFolder() & MakeRecordingFileName(True)
    remoteFileName = "Temp Recording.wav"
    
    bRecording = False
    bRecordingTransferInProgress = False
    
    recording_File_To_Send = vbNullString
    recording_Remote_Filename = vbNullString
    
    If modSoundCap.Save_And_Close_Record(recording_Name, localFileName) Then
        
        If Server Then
            'select client to send to
            If UBound(modVars.Clients) = 1 Then 'two clients - 0 = me, 1 = client
                If Clients(1).iSocket <> -1 Then
                    IP = Clients(1).iSocket
                Else
                    'shouldn't get here
                    IP = Clients(0).iSocket
                End If
            Else
                IP = modVars.IPChoice(Me, True, "Select a client to send the recording to...")
            End If
            
            If LenB(IP) = 0 Then Exit Sub
            
            On Error GoTo socketEH
            iSocket = CInt(IP)
            
            'voicePort should be open + listening, check anyway
            If ucVoiceTransfer.iCurSockStatus <> sckListening Then
                If ucVoiceTransfer.Listen(modPorts.VoicePort) = False Then
                    AddText "Error Listening on port " & modPorts.VoicePort & "(" & Err.Description & "), recording not sent", TxtError, True
                    Exit Sub
                End If
            End If
            
            frmVoiceTransfers.addToDebug "Sending " & localFileName
            frmVoiceTransfers.addToDebug logIndent & "Socket: " & IP
            frmVoiceTransfers.addToDebug logIndent & "(Waiting for them to connect)"
            
            'store the filename to send
            recording_File_To_Send = localFileName
            recording_Remote_Filename = remoteFileName
            
            SetInfo "Sending Recording... (Awaiting Connection)", False
            
            'send message, telling them to connect
            SendData eCommands.cmdOther & eOtherCmds.ConnectToServerVoicePort, iSocket
            bRecordingTransferInProgress = True
            
        Else
'##########################################################################
            'send to server
            IP = SckLC.RemoteHostIP
            
            SetInfo "Saved Recording, Connecting...", False
            bRecordingTransferInProgress = True
            
            
            frmVoiceTransfers.addToDebug "Sending " & localFileName
            frmVoiceTransfers.addToDebug logIndent & "(Connecting to server...)"
            
            ucVoiceTransfer.Connect IP, modPorts.VoicePort
            
            Tick = GetTickCount()
            
            Do
                Pause 10
            Loop While (ucVoiceTransfer.iCurSockStatus <> sckConnected) And _
                       (Tick + ConnectTimeOut > GetTickCount()) And _
                       Not modVars.Closing
            
            
            If modVars.Closing Then Exit Sub
            
            
            If frmMain.ucVoiceTransfer.iCurSockStatus = sckConnected Then
                
                frmVoiceTransfers.addToDebug logIndent & "Connected to server, sending..."
                
                SetInfo "Saved Recording, Sending (Connected)..." & RecordingsViewString, False
                frmVoiceTransfers.bRecordingCanceled = False
                
                If ucVoiceTransfer.SendFile(localFileName, remoteFileName) Then
                    
                    frmVoiceTransfers.addToDebug logIndent & "Sent"
                    
                    sTxt = GetFileName(localFileName)
                    AddText sTxt, , True
                    SetInfo sTxt & RecordingsViewString, False
                    
                    frmVoiceTransfers.AddVoiceTransfer localFileName, True
                    
                    Pause 200 'wait a while before disconnecting
                Else
                    If frmVoiceTransfers.bRecordingCanceled Then
                        frmVoiceTransfers.addToDebug logIndent & "[Canceled]"
                        SetInfo "Recording Transfer Canceled", True
                    Else
                        SetInfo "Recording Transfer Error - Error mid-transfer", True
                        frmVoiceTransfers.addToDebug logIndent & "Error - Mid-transfer disconnection"
                    End If
                End If
                
                ucVoiceTransfer.Disconnect
                frmVoiceTransfers.addToDebug logIndent & "Disconnected"
            Else
                frmVoiceTransfers.addToDebug logIndent & "Couldn't connect"
                
                ucVoiceTransfer.Disconnect
                
                SetInfo "Error Connecting to Server (Recording)", True
            End If
            
            bRecordingTransferInProgress = False
            frmVoiceTransfers.updateCurrent 0, vbNullString, False
'##########################################################################
        End If
        
    Else
        sTxt = "Error Saving Recording - " & modSoundCap.GetLastErrorText()
        SetInfo sTxt, True
        AddConsoleText sTxt
    End If
    
Else
    sTxt = "Error Stopping Recording - " & modSoundCap.GetLastErrorText()
    SetInfo sTxt, True
    AddConsoleText sTxt
End If


Exit Sub
socketEH:
frmVoiceTransfers.addToDebug logIndent & "Couldn't find target socket - Recording not sent"
SetInfo "Error - Socket Not Found, Couldn't Send Recording", True
bRecordingTransferInProgress = False
Me.ucVoiceTransfer.Disconnect
End Sub

Private Sub mnuOptionsWindow2Animation_Click()
mnuOptionsWindow2Animation.Checked = Not mnuOptionsWindow2Animation.Checked
If Me.Visible Then
    SetInfo "Window Animation " & IIf(mnuOptionsWindow2Animation.Checked, "En", "Dis") & "abled", False
End If
End Sub

Private Sub optnDraw_Click(Index As Integer)
' 0 = normal, 1 = straight, 2 = pick

optnDraw(Index).Value = True

Select Case Index
    Case 0
        PickingColour = False
        DrawingStraight = False
        
    Case 1
        PickingColour = False
        DrawingStraight = True
        
    Case 2
        PickingColour = True
        DrawingStraight = False
        
End Select


If DrawingStraight = False Then
    StraightPoint1.X = 0
    StraightPoint1.Y = 0
End If

End Sub

Private Sub ucVoiceTransfer_Connected(IP As String)
Dim sFName As String

If LenB(recording_File_To_Send) And Server Then '[And Server] just to make sure
    
    bRecordingTransferInProgress = True
    frmVoiceTransfers.bRecordingCanceled = False
    SetInfo "Client Connected, Sending Recording...." & RecordingsViewString, False
    
    frmVoiceTransfers.addToDebug logIndent & "Connection received, sending recording..."
    
    If ucVoiceTransfer.SendFile(recording_File_To_Send, recording_Remote_Filename) Then
        sFName = GetFileName(recording_File_To_Send)
        SetInfo sFName & RecordingsViewString, False
        AddText sFName, , True
        
        frmVoiceTransfers.addToDebug logIndent & "Sent"
        
        frmVoiceTransfers.AddVoiceTransfer recording_File_To_Send, True
        
        Pause 200
    Else
        If frmVoiceTransfers.bRecordingCanceled Then
            SetInfo "Recording Transfer Canceled", True
            frmVoiceTransfers.addToDebug logIndent & "[Canceled]"
        Else
            SetInfo "Recording Transfer Error - Disconnected Midway Through", True
            frmVoiceTransfers.addToDebug logIndent & "Error - Mid-transfer disconnection"
        End If
    End If
    
    recording_File_To_Send = vbNullString
    recording_Remote_Filename = vbNullString
    
    bRecordingTransferInProgress = False
    frmVoiceTransfers.updateCurrent 0, vbNullString, False
    
    frmVoiceTransfers.addToDebug logIndent & "Disconnected"
    
    ucVoiceTransfer.Disconnect
    
    If Server Then
        frmVoiceTransfers.addToDebug "Relistening on port " & modPorts.VoicePort
        ucVoiceTransfer.Listen modPorts.VoicePort
    End If
    
Else
    'we're the client, receiving a voice recording off a server..?
    SetInfo "Connected to Server, Receiving Recording..." & RecordingsViewString, False
    frmVoiceTransfers.addToDebug logIndent & "Connected to server, receiving recording..."
End If


End Sub

Private Sub ucVoiceTransfer_Diconnected()
frmVoiceTransfers.updateCurrent 0, vbNullString, False
End Sub

Private Sub ucVoiceTransfer_Error(Description As String, ErrNo As eFTErrors)
Dim sTxt As String

sTxt = "Recording Transfer Error - " & Description

SetInfo sTxt, True
AddConsoleText sTxt

frmVoiceTransfers.updateCurrent 0, vbNullString, False
End Sub

Private Function MakeRecordingFileName(bSent As Boolean) As String

MakeRecordingFileName = "Recording " & MakeTimeFile() & " (" & IIf(bSent, "Sent", "Received") & ").wav"

End Function

Private Sub ucVoiceTransfer_ReceivedFile(sFileName As String)
'add to log folder + play
Dim sPath As String, sTxt As String

sPath = GetCurrentLogFolder() & MakeRecordingFileName(False) 'move it to sPath

frmVoiceTransfers.updateCurrent 0, vbNullString, False

On Error GoTo copyEH
FileCopy sFileName, sPath
Kill sFileName

'play eet
sTxt = GetFileName(sPath) & " (Playing)" & RecordingsViewString
AddText sTxt, , True
SetInfo sTxt & RecordingsViewString, False
modAudio.PlayFileNameSound sPath

frmVoiceTransfers.AddVoiceTransfer sPath, False

Exit Sub
copyEH:
AddText Trim$(InfoStart) & vbNewLine & _
        "Received Voice Recording, Error Moving File" & vbNewLine & _
        Err.Description & vbNewLine & _
        "Recording is at " & modSettings.GetUserSettingsPath() & vbNewLine & _
        Trim$(InfoEnd), TxtError, , True

End Sub

Private Sub ucVoiceTransfer_ReceivingFile(sFileName As String, ByVal BytesReceived As Long, ByVal BytesRemaining As Long, ByVal lTotalBytes As Long)
frmVoiceTransfers.updateCurrent 100 * BytesReceived / lTotalBytes, sFileName, False
End Sub
Private Sub ucVoiceTransfer_SendingFile(sFileName As String, ByVal BytesSent As Long, ByVal BytesRemaining As Long, ByVal lTotalBytes As Long)
frmVoiceTransfers.updateCurrent 100 * (lTotalBytes - BytesRemaining) / lTotalBytes, GetFileName(sFileName), True
End Sub

Private Sub tmrVoice_Timer()

time_Recorded = time_Recorded + tmrVoice.Interval

If time_Recorded > recording_Time Then
    
    'stop the recording
    mnuOptionsMessagingWindowsRecord_Click
    
Else
    progVoice.Value = 100 * time_Recorded / recording_Time
End If

End Sub

'################################################################################################
Public Sub SetInfo(ByVal sTxt As String, ByVal bError As Boolean)

If Me.Visible = False Then Exit Sub

On Error Resume Next
'lblInfo.Caption = sTxt
picInfo.Visible = True
picInfo.WhatsThisHelpID = Abs(bError)
sInfoText = sTxt

iInfoTimer = 0

'picInfo.Left = 3480 - don't need to, plus it messes up with vista border
picInfo.height = 255
picInfo.width = Me.ScaleWidth - picInfo.Left - 100

tmrInfoHide.Enabled = True
End Sub

Private Sub cmdAprilFoolReset_Click()

cmdAprilFoolReset.Visible = False
modDisplay.Mirror Me, , False
Me.Hide
Me.Show

End Sub

Private Sub cmdSlash_Click()
cmdSlash.Enabled = False
frmSystray.mnuCommandsTestSpeech.Enabled = Len(txtOut.Text)
frmSystray.mnuCommandsStopSpeech.Enabled = (modSpeech.nSpeechStatus = sSpeaking)
PopupMenu frmSystray.mnuCommands, , cmdSlash.Left, cmdSlash.Top + cmdSlash.height
cmdSlash.Enabled = True '(Status = Connected)
End Sub

Private Sub cmdSmile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If rtfIn.SmileyBoxVisible Then
    'cancel, allow it to lose focus
    ReleaseCapture
End If
End Sub

Public Sub mnuCommandsDescribe_Click()
Const kAction = "Run for the hills!", kCommand = "/Describe"

InsertAction kAction, kCommand
End Sub

Public Sub mnuCommandsMe_Click()
Const kAction = "likes cake", kCommand = "/Me"

InsertAction kAction, kCommand
End Sub

Private Sub InsertAction(kAction As String, kCommand As String)
Dim i As Integer

txtOut.Selstart = 0
If LenB(txtOut.Text) = 0 Then
    txtOut.SelText = kCommand & vbSpace & kAction
Else
    txtOut.SelText = kCommand & vbSpace
End If

txtOut.Selstart = Len(kCommand) + 1
i = Len(txtOut.Text) - Len(kCommand)
If i > 0 Then txtOut.Sellength = i

SetFocus2txtOut

End Sub

'Public Sub mnuCommandsGregory_Click()
'AddTags UCase$(modMessaging.GregTag) '"LANCS"
'End Sub
'
'Public Sub mnuCommandsJeffery_Click()
'AddTags UCase$(modMessaging.JeffTag) '"BLOCKEDNOSE"
'End Sub

Public Sub mnuCommandsBold_Click()
AddTags BoldTag
End Sub
Public Sub mnuCommandsUnderline_Click()
AddTags UnderLineTag
End Sub
Public Sub mnuCommandsItalic_Click()
'Dim iStart As Integer, iLen As Integer
'
'iStart = txtOut.Selstart
'iLen = txtOut.Sellength
'
'AddTags ItalicTag
'
'txtOut.Selstart = iStart + 3
'txtOut.Sellength = iLen
'
'AddTags "EMPH"
AddTags ItalicTag
End Sub

Public Sub mnuCommandsSpeechEmph_Click()
AddTags "e"
End Sub

Public Sub mnuCommandsSpeechPause_Click()
Const Quote = """"
Const tagToAdd = "<SILENCE MSEC =" & Quote & "250" & Quote & "/>"

If txtOut.Sellength > 0 Then txtOut.Sellength = 0

txtOut.SelText = tagToAdd
End Sub

Public Sub mnuCommandsSpeechPitch_Click()
'AddTags "pitch middle=""5"""
AddTags "p5"
End Sub

Public Sub mnuCommandsSpeechSpeed_Click()
'AddTags "rate speed=""5"""
AddTags "s5"
End Sub

Public Sub mnuCommandsSpeechVolume_Click()
'AddTags "volume level=""50"""
AddTags "v50"
End Sub

Public Sub mnuCommandsSpeechActorsHL_Click()
AddTags "hl"
End Sub

Public Sub mnuCommandsSpeechActorsJW_Click()
AddTags "hi"
End Sub

Public Sub mnuCommandsStopSpeech_Click()
modSpeech.StopSpeech
End Sub

Public Sub mnuCommandsTestSpeech_Click()
modSpeech.Say txtOut.Text, , , True
End Sub

Public Sub mnuCommandsBugAlert_Click()

txtOut.Text = "/describe BUG ALERT"

On Error Resume Next
txtOut.Selstart = Len(txtOut.Text)
SetFocus2txtOut

End Sub

Private Sub AddTags(ByVal sTag As String)

Dim iSelEnd As Integer

sTag = LCase$(sTag)

With txtOut
    iSelEnd = .Selstart + .Sellength
    .Sellength = 0
    .SelText = MakeTag(sTag)
    
    .Selstart = iSelEnd + Len("<" & sTag & ">")
    .SelText = MakeTag(sTag, True)
    
    .Selstart = iSelEnd + 2 + Len(sTag)
    
    
    SetFocus2txtOut
End With

End Sub

Private Sub lstConnected_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstConnected.Text = ConnectedListPlaceHolder Then
    lstConnected.ListIndex = -1
End If
End Sub

Private Sub lstConnected_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lbIndex As Long, lTopIndex As Long, lSingleHeight As Long


ListBoxMouseInfo Y, lstConnected, lbIndex, lTopIndex, lSingleHeight


If Not (lbIndex > -1 And lbIndex < lstConnected.ListCount) Then
    'user click on an empty space
    lstConnected.ListIndex = -1
End If


End Sub

Private Sub mnuDevAdvCmdsNoFTPCallbacks_Click()
mnuDevAdvCmdsNoFTPCallbacks.Checked = Not mnuDevAdvCmdsNoFTPCallbacks.Checked
End Sub

'Private Sub mnuFileShortcut_Click()
'Dim CmdLine As String
'
'CmdLine = modVars.Password("Enter the command line for the shortcut", Me, "Shortcut command line", _
'          "/console /host 1", False)
'
'
'
'If LenB(CmdLine) Then
'    Create_Shortcut AppPath() & App.EXEName, GetDesktopPath(), "Shortcut to Communicator", CmdLine
'    SetInfo "Shortcut Placed on your Desktop", False
'End If
'
'End Sub
'Private Sub Create_Shortcut(ByVal sTargetPath As String, ByVal sShortCutPath As String, ByVal sShortCutName As String, _
'                            Optional ByVal sArguments As String, Optional ByVal sWorkPath As String, _
'                            Optional ByVal eWinStyle As WshWindowStyle = vbNormalFocus, Optional ByVal iIconNum As Integer)
'
''Requires reference to Windows Script Host Object Model
'Dim oShell As IWshRuntimeLibrary.WshShell
'Dim oShortCut As IWshRuntimeLibrary.WshShortcut
'
'Set oShell = New IWshRuntimeLibrary.WshShell
'Set oShortCut = oShell.CreateShortcut(oShell.SpecialFolders(sShortCutPath) & _
'                                      "\" & sShortCutName & ".lnk")
'With oShortCut
'    .TargetPath = sTargetPath
'    .Arguments = sArguments
'    .WorkingDirectory = sWorkPath
'    .WindowStyle = eWinStyle
'    .IconLocation = sTargetPath & "," & iIconNum
'    .Save
'End With
'
'Set oShortCut = Nothing
'Set oShell = Nothing
'
'End Sub
Private Function GetDesktopPath() As String
GetDesktopPath = GetSpecialFolder(0)
End Function
Private Function GetSpecialFolder(CSIDL As Long) As String
Dim lR As Long
Dim IDL As ITEMIDLIST
Dim Path As String

lR = SHGetSpecialFolderLocation(Me.hWnd, CSIDL, IDL)

If lR = 0 Then
    Path = Space$(512)
    lR = SHGetPathFromIDList(IDL.mkid.cb, Path)
    GetSpecialFolder = Trim0(Path) 'Left$(Path, InStr(Path, Chr$(0)) - 1)
    
    modDirBrowse.CoTaskMemFree VarPtr(IDL)
Else
    GetSpecialFolder = vbNullString
End If

End Function

Private Sub mnuFileThumb_Click()

If modLoadProgram.frmThumbNail_Loaded Then
    Unload frmThumbnail
Else
    Load frmThumbnail
End If

End Sub

Private Sub mnuHelpBug_Click()
Unload frmBug
Load frmBug
frmBug.Show vbModeless, Me
End Sub

Public Sub mnuFileMini_Click()

mnuFileMini.Checked = Not mnuFileMini.Checked

Unload frmMini
If mnuFileMini.Checked Then
    'if msgboxex("This can cause Communicator to crash, since it measures bytes sent on the network, so " & _
        "the variables can't hold numbers that are too large. Are you sure?"
    
    On Error GoTo EH
    Load frmMini
End If

Exit Sub
EH:
Unload frmMini
mnuFileMini.Checked = False
mnuFileMini.Enabled = False
AddText "Error Loading Mini Window: " & Err.Description, TxtError, True
End Sub

Private Sub mnuHelpBugReport_Click()

SendInfoMessage BugReportStr, , True
AddText BugReportStr, TxtError, True

modAudio.PlayBugReport

End Sub

Private Sub mnuHelpSysTime_Click()
Dim lSocket As Long, lR As Long, lStart As Long
Dim sIP As String, sData As String
Dim bConnected As Boolean, bReceived As Boolean, bTimeOut As Boolean

AddText "Checking NIST Time...", , True


lSocket = modWinsock.TCP_CreateSocket()
If lSocket <> WINSOCK_ERROR Then
    lR = modWinsock.TCP_BindSocket(lSocket)
    
    If lR <> WINSOCK_ERROR Then
        
        sIP = modWinsock.HostNameToIP("time.nist.gov")
        
        lR = modWinsock.TCP_Connect(lSocket, sIP, 13)
        
        If lR = WINSOCK_ERROR Then
            
            bTimeOut = False
            lStart = GetTickCount()
            Do
                
                If modWinsock.TCP_Connected(lSocket) Then
                    bConnected = True
                    bTimeOut = False
                    
                    lStart = GetTickCount()
                    Do
                        If modWinsock.TCP_ReceiveData(lSocket, sData) Then
                            bReceived = True
                            
                            If modVars.Closing = False Then
                                Call ProcessReceivedDateTime(sData)
                            End If
                        ElseIf lStart + 5000 < GetTickCount() Then
                            bTimeOut = True
                        End If
                        
                    Loop Until bReceived Or modVars.Closing Or bTimeOut
                    
                    
                    
                    DoEvents
                ElseIf lStart + 5000 < GetTickCount() Then
                    bTimeOut = True
                End If
                
                
                DoEvents
            Loop Until bConnected Or modVars.Closing Or bTimeOut
            
            If modVars.Closing Then Exit Sub
            
            If bTimeOut Then
                AddText "Time Check - Error, Connection Timed Out", TxtError, True
            End If
            
        Else
            AddText "Error Connecting to time.nist.gov - " & Err.LastDllError, TxtError, True
        End If
    Else
        AddText "Time Check - Error Binding Socket", TxtError, True
    End If
Else
    AddText "Time Check - Error Creating Socket - " & Err.LastDllError, TxtError, True
End If

If lSocket Then modWinsock.TCP_CloseSocket lSocket

End Sub

Private Sub ProcessReceivedDateTime(sData As String)
Dim dDate As Date
Dim lStart As Long
Dim msAdj As Long

lStart = GetTickCount()

If FormatNetTime(sData, dDate, msAdj) Then
    
    AddText "NIST Server Time is " & CStr(dDate), , True
    AddText "(This time will be adjusted, since asking the question takes time)", , True
    
    If Question("Set time to " & CStr(dDate) & "?", mnuHelpSysTime) = vbYes Then
        
        If modVars.SetSystemTime(dDate, msAdj + GetTickCount() - lStart) Then
            AddText "System Time Set (Time may have been adjusted - daylight saving)", , True
        Else
            AddText InfoStart & vbNewLine & _
                "Error Setting System Time" & vbNewLine & _
                modVars.DllErrorDescription() & vbNewLine & _
                InfoEnd, TxtError
            
        End If
    Else
        AddText "Canceled", TxtInfo, True
    End If
'Else
    'error text added
End If

End Sub

Private Function FormatNetTime(ByVal NetTime As String, ByRef dDate As Date, _
    msAdj As Long) As Boolean  'format the string received from time server

Dim strDate As String
Dim strTime As String

'Received string example
'JJJJJ YR-MO-DA HH:MM:SS TT L H msADV UTC(NIST) OTM
'52587 02-11-12 22:05:25 00 0 0 636.1 UTC(NIST) *

On Error GoTo DateTimeEH
'Extract the Date from the received string
strDate = Mid$(NetTime, 11, 5) & "-" & Mid$(NetTime, 8, 2)

'Extract the time from the received string
strTime = Mid$(NetTime, 17, 8)

'Check that extracts are suitable date and time - then convert using CDate
If IsDate(strDate) And IsDate(strTime) Then
    dDate = CDate(strDate & vbSpace & strTime)
    
    'Extract the downloaded millisec time offset
    'this is the transmission line time delay the server calculated and offset the time by
    msAdj = val(Mid$(NetTime, 33, 3) + Mid$(NetTime, 37, 1))
    
    'Server health is bad, server actually reports its own condition, do not set if bad
    If Mid$(NetTime, 31, 1) <> "0" Then 'is servers' health bad?, then do not set
        AddText "The server reports that its time may not be accurate at the moment", TxtError, True
        
        FormatNetTime = False
    Else
        FormatNetTime = True
    End If
Else
    'Date / Time format is wrong
DateTimeEH:
    
    AddText InfoStart & vbNewLine & _
        "The data string from the server caused a Time and/or Date formatting error." & vbNewLine & _
        "Try again later." & vbNewLine & _
        InfoEnd, TxtInfo
    
    FormatNetTime = False
End If
  
End Function

'######################################################################
Public Sub mnuOnlineFTPDLO_Click(Index As Integer)
Dim i As Integer
'Static bTold As Boolean

If Index = eFTP_Methods.FTP_HTTP Then
    If modFTP.iCurrent_FTP_Details > 0 Then
        Show_Server_HTTP_Error
        Exit Sub
    ElseIf Question("The HTTP Method locks up Communicator while it downloads, are you sure?", _
            mnuOnlineFTPDL) = vbNo Then
        
        Exit Sub
    'Else
        'AddText "HTTP Method Selected", , True
    End If
End If

For i = mnuOnlineFTPDLO.LBound To mnuOnlineFTPDLO.UBound
    mnuOnlineFTPDLO(i).Checked = False
Next i

On Error Resume Next 'in case someone messes with the settings file
mnuOnlineFTPDLO(Index).Checked = True

'If Not bTold Then
'    AddText "Upload will still use the Manual Method", , True
'    bTold = True
'End If

End Sub

Private Sub mnuOnlineFTPPassive_Click()
mnuOnlineFTPPassive.Checked = Not mnuOnlineFTPPassive.Checked
End Sub

Private Sub mnuOnlineFTPServerCustom_Click()
Dim i As Integer

'If mnuOnlineFTPServerCustom.Checked Then Exit Sub
'^ allow them to reconfigure

Unload frmFTPServer
Load frmFTPServer
frmFTPServer.Show vbModal, Me

If modFTP.FTP_iCustomServer > -1 Then
    modFTP.iCurrent_FTP_Details = modFTP.FTP_iCustomServer
    
    mnuOnlineFTPServerCustom.Checked = True
    
    For i = 0 To mnuOnlineFTPServerAr.UBound
        mnuOnlineFTPServerAr(i).Checked = False
    Next i
End If

End Sub

Private Sub mnuOnlineFTPServerMsg_Click()
mnuOnlineFTPServerMsg.Checked = Not mnuOnlineFTPServerMsg.Checked

'If mnuOnlineFTPServerMsg.Checked Then
    'If Question("Download now?", mnuOnlineFTPServerMsg) = vbYes Then
        'DownloadServerMessage
    'End If
'End If

End Sub

Public Sub mnuOnlineFTPULO_Click(Index As Integer)
Dim i As Integer

For i = mnuOnlineFTPULO.LBound To mnuOnlineFTPULO.UBound
    mnuOnlineFTPULO(i).Checked = False
Next i

On Error Resume Next 'in case someone messes with the settings file
mnuOnlineFTPULO(Index).Checked = True
End Sub

Public Function GetServerName() As String
Dim i As Integer

For i = 0 To mnuOnlineFTPServerAr.UBound
    If mnuOnlineFTPServerAr(i).Checked Then
        GetServerName = Mid$(mnuOnlineFTPServerAr(i).Caption, 5)
        Exit For
    End If
Next i

If i = mnuOnlineFTPServerAr.UBound + 1 Then
    GetServerName = "Custom Server"
End If

End Function

Public Sub mnuOnlineFTPServerAr_Click(Index As Integer)
Dim i As Integer
'Static bTold As Boolean
'Dim Ans As VbMsgBoxResult

If mnuOnlineFTPServerAr(Index).Checked Then Exit Sub

If Index > 0 Then
    If mnuOnlineFTPDLO(eFTP_Methods.FTP_HTTP).Checked Then
        Show_Server_HTTP_Error
        Exit Sub
    Else
        'If Not bTold And Me.Visible Then
            'AddText "The backup server is only used if the primary one isn't working", TxtQuestion, True
            'bTold = True
        'End If
        
        'If Me.Visible Then
            'Ans = Question("Are you sure you want to use a secondary server?", _
                mnuOnlineFTPServer)
        'Else
            'Ans = vbIgnore
        'End If
        
        'If Ans = vbNo Then
            'Exit Sub
        'ElseIf Ans = vbYes Then
            'AddText "Secondary Server Selected", , True
            SetInfo "Secondary/Backup Server Selected", False
        'End If
        
        
    End If
Else
    SetInfo "Primary Server Selected", False
End If

mnuOnlineFTPServerCustom.Checked = False
For i = 0 To mnuOnlineFTPServerAr.UBound
    mnuOnlineFTPServerAr(i).Checked = (Index = i)
Next i


modFTP.iCurrent_FTP_Details = Index

End Sub

Private Sub Show_Server_HTTP_Error()
If Me.Visible Then SetInfo "Error - Only the Primary Server supports HTTP Download", True
'in case at startup
End Sub

'######################################################################

Public Sub mnuOptionsAdvDisplayConnF_Click(Index As Integer)

mnuOptionsAdvDisplayConnF(Index).Checked = True
mnuOptionsAdvDisplayConnF(1 - Index).Checked = False


If Index = 1 Then
    'select
    cmdArray(2).Caption = "Connect"
Else
    'manual
    cmdArray(2).Caption = "Manual" ' Connect"
End If


cmdXPArray(2).Caption = cmdArray(2).Caption


End Sub

Private Sub mnuOptionsAdvNetworkRefresh_Click()
mnuOptionsAdvNetworkRefresh.Checked = Not mnuOptionsAdvNetworkRefresh.Checked
End Sub

Public Sub mnuOptionsAlertsStyle_Click(Index As Integer)
Dim i As Integer

'##########
'menu checkage
For i = mnuOptionsAlertsStyle.LBound To mnuOptionsAlertsStyle.UBound
    mnuOptionsAlertsStyle(i).Checked = False
Next i
mnuOptionsAlertsStyle(Index).Checked = True
'##########

'##########
'variable settage
modAlert.bBalloonTips = (Index = 0)
If Index > 0 Then modAlert.AlertStyle = Index - 1
'##########

'##########
'displayage
If Me.Visible And (modSettings.bLoadingSettings = False) Then
    Select Case Index
        Case 0
            'balloon
            frmSystray.ShowBalloonTip "Balloon tips will be shown", , NIIF_INFO, , True
            
            
        Case Else
            
            frmSystray.HideBalloon
            modAlert.ShowAlert "Hello there", "This type of alert will be shown"
            
    End Select
End If
'##########

End Sub

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

Private Sub mnuOptionsMessagingDisplayShowHost_Click()
mnuOptionsMessagingDisplayShowHost.Checked = Not mnuOptionsMessagingDisplayShowHost.Checked
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
picInfo.ForeColor = IIf(picInfo.WhatsThisHelpID = 0, vbBlue, vbRed)
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
If sF >= 6 Then
    prtfFontSize = sF
    'rtfIn.Font.Size = sF
End If
End Property

Public Property Get rtfFontName() As String
rtfFontName = prtfFontName
End Property

Public Property Let rtfFontName(sF As String)
If LenB(sF) Then
    prtfFontName = sF
    txtOut.Font = sF
End If
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
    If mnuOptionsAdvDisplayConnF(1).Checked Then
        'primary function = Selected, so 2nrdy = manual
        mnuFileManual_Click
    Else
        cmdAdd_Proper_Click
    End If
End If
End Sub

Private Sub cmdShake_KeyPress(KeyAscii As Integer)
SetFocus2txtOut
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
Dim iClient As Integer

If imgBig.tag <> CStr(Index) Then
    
    imgBig.Stretch = False
    wbDP.Visible = False
    Set imgBig.Picture = imgDP(Index).Picture
    
    
    'assume can do
    picBig.Left = imgDP(Index).Left + imgDP(Index).width / 2 - imgBig.width / 2
    picBig.width = imgBig.width
    picBig.height = imgBig.height - 10
    
    
    If picBig.Left < MinLeft Then picBig.Left = MinLeft
    
    
    If (picBig.Left + picBig.width) > (Me.width - 20) Then
        'no can do
        imgBig.Stretch = True
        
        SetInfo "Picture is too large, resize Communicator to see fully", True
        
        ResetpicBigXY
        
        picBig.Left = imgDP(Index).Left + imgDP(Index).width - picBig.width / 2
    End If
    
    If Index = 0 Then
        iClient = FindClient(-1)
    Else
        iClient = FindClient(Index)
    End If
    
    If iClient > -1 Then
        If Clients(iClient).bDPIsGIF Then
            
            ShowGIFPicture iClient
            
            wbDP.width = imgBig.width + 200
            wbDP.height = imgBig.height + 300
            
            
            imgBig.Visible = False
        Else
            imgBig.Visible = True
        End If
    Else
        imgBig.Visible = True
    End If
    
    picBig.Visible = True
    
    imgBig.tag = CStr(Index)
End If

End Sub

Private Function HTML_Doc_oncontextmenu() As Boolean
HTML_Doc_oncontextmenu = False

imgDP_MouseDown iCurrentMouseOver, vbRightButton, 0, 0, 0
End Function
Private Function HTML_Doc_ondblclick() As Boolean
imgDP_DblClick iCurrentMouseOver
End Function

Private Function ShowGIFPicture(ByVal iClient As Integer) As Boolean
Dim Path As String
Const BeforeStr = "about:<html><body scroll='no'><img src='", _
      AfterStr = "'></img></body></html>"

Path = modDP.GetClientDPStr(iClient)

If FileExists(Path) Then
    
    wbDP.Navigate BeforeStr & Path & AfterStr
    wbDP.Visible = True
    
Else
    ShowGIFPicture = False
    wbDP.Visible = False
End If

End Function

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
imgBig.tag = vbNullString
End Sub

Private Sub HideExtras() 'Optional ByVal bStatusAndName As Boolean = True)
HidePicClient
HidePicBig

On Error GoTo EH
If Not (Screen.ActiveControl Is Nothing) Then
    If Screen.ActiveControl.Name <> txtStatus.Name Then
        txtStatus.Visible = False
    End If
End If

'If bStatusAndName Then
'    txtStatus.Visible = False
'
'    On Error GoTo EH
'    If Screen.ActiveControl.Name = txtName.Name Then
'        Rename txtName.Text
'
'        On Error GoTo EH
'        If Status = Connected Then
'            txtOut.SetFocus
'        Else
'            rtfIn.SetFocus
'        End If
'
'    End If
'End If

EH:
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

Public Sub mnuDPOpen_Click()

modVars.OpenFolder vbNormalFocus, modDP.DP_Dir_Path

End Sub

Public Sub mnuDPView_Click()
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
modSettings.SaveUsedIPs

SetInfo "Settings Saved", False
End Sub

Private Sub mnuFileSettingsUserProfileLoad_Click()

If modSettings.ImportSettings(modSettings.GetSettingsFile(), False, True) Then
    modSettings.LoadUsedIPs
    SetInfo "Loaded Settings", False
'Else
    'error text added
End If

End Sub

Private Sub SaveUserProfileSettings()
modSettings.ExportSettings modSettings.GetSettingsFile(), False
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

Public Sub mnuFontColour_Click()
txtOut_DblClick
End Sub

Public Sub mnuFontCopy_Click()
Clipboard.SetText txtOut.Text
End Sub

Public Sub mnuFontDialog_Click()

With Me.Cmdlg
    .FontName = rtfFontName
    .FontSize = rtfFontSize
    .FontItalic = rtfItalic
    .FontBold = rtfBold
    
    
    .flags = cdlCFForceFontExist Or cdlCFLimitSize Or cdlCFBoth 'Or cdlCFScalableOnly
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

If Not Server Then
    'yo server, we don't have a DP
    SendData eCommands.SetClientVar & eClientVarCmds.SetDPSet & "0"
End If


ResetImgDP 0
For i = 1 To imgDP.UBound
    Unload imgDP(i)
Next i

modDP.DelPics

SetInfo "Pictures will be updated soon", False

End Sub

Private Sub mnuOptionsMessagingLoggingAutoSave_Click()
mnuOptionsMessagingLoggingAutoSave.Checked = Not Me.mnuOptionsMessagingLoggingAutoSave.Checked
'tmrAutoSave.Enabled = mnuOptionsMessagingAutoSave.Checked
End Sub

Private Sub mnuOptionsMessagingDisplaySmiliesMSN_Click()
mnuOptionsMessagingDisplaySmiliesMSN.Checked = Not mnuOptionsMessagingDisplaySmiliesMSN.Checked

If mnuOptionsMessagingDisplaySmiliesMSN.Checked Then
    If mnuOptionsMessagingDisplaySmiliesComm.Checked Then
        mnuOptionsMessagingDisplaySmiliesComm.Checked = False
        SetInfo "Switched to MSN Smilies", False
    Else
        SetInfo "MSN Smilies Enabled", False
    End If
Else
    SetInfo "Smilies Disabled", False
End If

ApplySmileySettings
End Sub

Private Sub mnuOptionsMessagingDisplaySmiliesComm_Click()

mnuOptionsMessagingDisplaySmiliesComm.Checked = Not mnuOptionsMessagingDisplaySmiliesComm.Checked

If mnuOptionsMessagingDisplaySmiliesComm.Checked Then
    If mnuOptionsMessagingDisplaySmiliesMSN.Checked Then
        mnuOptionsMessagingDisplaySmiliesMSN.Checked = False
        SetInfo "Switched to Communicator Smilies", False
    Else
        SetInfo "Communicator Smilies Enabled", False
    End If
Else
    SetInfo "Smilies Disabled", False
End If

ApplySmileySettings
End Sub

Public Sub ApplySmileySettings()

If mnuOptionsMessagingDisplaySmiliesComm.Checked Or Me.mnuOptionsMessagingDisplaySmiliesMSN.Checked Then
    rtfIn.EnableSmiles = True
    cmdSmile.Enabled = (Status = Connected)
    
    rtfIn.ShowNewSmilies = mnuOptionsMessagingDisplaySmiliesComm.Checked
    
    If mnuOptionsMessagingDisplaySmiliesComm.Checked And mnuOptionsMessagingDisplaySmiliesMSN.Checked Then
        mnuOptionsMessagingDisplaySmiliesMSN.Checked = False
    End If
Else
    rtfIn.EnableSmiles = False
    cmdSmile.Enabled = False
    
    rtfIn.HideSmilies
End If

End Sub

Public Sub mnuStatusAway_Click()

If UCase$(LastStatus) = AFKStr Then
    ReStatus vbNullString
    
    If Not modVars.Closing Then modSpeech.Say "Status Removed"
Else
    ReStatus AFKStr
    
    If Not modVars.Closing Then modSpeech.Say "Status set to " & AFKStr
End If

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

'LastStatus = "AFK"
'txtStatus.Text = LastStatus


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

'Private Function RemoveAwayStatus() As String
'RemoveAwayStatus = Left$(LastName, Len(LastName) - Len(ksAway))
'End Function

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

Public Sub mnuStatusResetName_Click()

If Me.mnuOptionsMessagingDisplaySysUserName.Checked Then
    Rename modVars.User_Name
Else
    Rename SckLC.LocalHostName
End If

End Sub

Private Sub AutoSave(Optional bForce As Boolean = False)
'Dim Path As String
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
    'Path = GetLogPath()
    
    On Error GoTo EH
'    If FileExists(Path, vbDirectory) = False Then
'        MkDir Path
'    End If
    
    rtfIn.SaveFile GetLogPath() & "AutoSave.rtf", rtfRTF 'rtfText
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

Public Sub ShowSB_IP()

Dim brIP As Boolean, rIP As String

'If LenB(lIP) = 0 Then lIP = frmMain.SckLC.LocalIP
'If LenB(rIP) = 0 Then rIP = frmMain.GetIP()
'brIP = Not CBool(InStr(1, rIP, "Error:", vbTextCompare))

'##########################################################################
rIP = modWinsock.RemoteIP

If Len(rIP) > Len("xxx.xxx.xxx.xxx") Then
    brIP = False
ElseIf LenB(rIP) = 0 Then
    brIP = False
Else
    brIP = True
End If

'##########################################################################
If modWinsock.LocalIP = "0.0.0.0" Then modWinsock.SetLocalIP Me.SckLC.LocalIP


SetPanelText LIPHeading & modWinsock.LocalIP & Space$(3) & _
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
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If frmMain.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessageByLong frmMain.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
'ElseIf Button = vbRightButton Then
    'PopupMenu
End If

End Sub

Public Sub Form_Terminate()

If modLoadProgram.frmMain_Loaded = False Then
    If modLoadProgram.bIsIDE = False Then 'otherwise the IDE would close
        'end api equivalent
        Debug.Assert False
        Call ExitProcess(e_Normal_Unload)
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
        PopupMenu frmSystray.mnuDP, , , , frmSystray.mnuDPView
    End If
End If

End Sub

Private Sub imgStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub
Private Sub imgStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtName_MouseMove Button, Shift, X, Y
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
            
            If picClient.tag <> CStr(iClient) Then
                With picClient
                    
                    .Move .Left, (lbIndex - lTopIndex) * lSingleHeight + lstConnected.Top
                    
                    
                    'highlighted top bit
                    '13 px = 195 twips
                    ShowClientInfo iClient
                    
                    
                    If picBig.Visible Then picBig.Visible = False
                    
                    
                    .tag = CStr(iClient)
                End With
            End If
            
            
            
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
Const YSep = 250, RedX = "REDX"

With picClient
    .Left = lstConnected.Left + lstConnected.width
    
    
    .Cls
    .ZOrder vbBringToFront
    
    
    BackGroundpicClientLine
    
    
    .CurrentX = 75
    .CurrentY = 10
    .ForeColor = vbWhite
    picClient.Print "User Details for " & IIf(LenB(Clients(iClient).sName), Clients(iClient).sName, "[?]")
    
    
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
    
    
    
    
    '.ToolTipText = GetClientInfo(iClient)
    
    
    
    
    
    
    
    '##############################################################
    'Set Picture
    If Clients(iClient).IPicture Is Nothing Then
        'prevent flickering
        If imgClientDP.tag <> RedX Then
            'Red X
            Set imgClientDP.Picture = GetRedX()
            imgClientDP.tag = RedX
        End If
        
    Else
        Set imgClientDP.Picture = Clients(iClient).IPicture
        imgClientDP.tag = vbNullString
    End If
    '##############################################################
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
lSingleHeight = SendMessageByLong(lst.hWnd, LB_GETITEMHEIGHT, 0, 0) * Screen.TwipsPerPixelY
lTopIndex = SendMessageByLong(lst.hWnd, LB_GETTOPINDEX, 0, 0)

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
picClient.tag = vbNullString
End Sub

Private Sub mnuDevAdvCmdsDebug_Click()
'mnuDevAdvCmds.Checked = Not mnuDevAdvCmds.Checked
'menu is auto-set
modVars.bDebug = Not modVars.bDebug
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
Dim Ans As VbMsgBoxResult
Dim bReShow As Boolean
Dim cPos As PointAPI

Const Max_Notification_Len = 150

bSpeak = True
If mnuFileGameMode.Checked = False Then
    'going to be checked...
    
    sTmp = modVars.Password("Enter a custom notification (Sent to clients)", Me, "Custom Message", _
        modSpaceGame.sGameModeMessage, False, Max_Notification_Len)
    
    
    If LenB(sTmp) Then
        If LenB(sTmp) > Max_Notification_Len Then
            sTmp = Left$(sTmp, Max_Notification_Len)
        End If
        
        modSpaceGame.sGameModeMessage = sTmp
        mnuFileGameMode.Checked = True
        
        'get game window
        Ans = MsgBoxEx("Ensure your mouse is over the game window, and Communicator will monitor it" & _
                       " so Game Mode can be deactivated when it closes" & vbNewLine & vbNewLine & _
                       "Select No to stop Communicator monitoring the game window." & vbNewLine & _
                       "Cancel will deactivate Game Mode", "See the above message" _
                       , vbYesNoCancel, "Select Game Window - Communicator")
        
        
        
        If Ans = vbYes Then
            'find game window
            If Me.Visible Then
                Me.ShowForm False, False
                bReShow = True
            End If
            
            'Get cursor position
            GetCursorPos cPos
            'Get window handle from point
            GameWindowhWnd = WindowFromPoint(cPos.X, cPos.Y)
            
            If bReShow Then
                Me.ShowForm , False
            End If
            
        ElseIf Ans = vbCancel Then
            mnuFileGameMode.Checked = False
            bSpeak = False
        End If
        
        
        
        If bSpeak Then
            If Status = Connected Then
                SendInfoMessage sTmp
                AddText sTmp, , True
            End If
        End If
    Else
        mnuFileGameMode.Checked = False
        bSpeak = False
    End If
Else
    mnuFileGameMode.Checked = False
End If

frmSystray.mnuPopupGameMode.Checked = mnuFileGameMode.Checked
Call RefreshIcon
Call GetTrayText '+set

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
        
        modSettings.ExportSettings Path, False
        SetInfo "Exported Settings", False
        
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
        
        If modSettings.ImportSettings(Path, False, True) Then
            SetInfo "Imported Settings", False
        End If
        
    End If
End If
End Sub

'Private Sub mnuFileSettingsRMenu_Click()
'
'If mnuFileSettingsRMenu.Checked Then
'    modVars.RemoveFromRightClick RightClickExt, RightClickMenuTitle
'
'    AddText "Communicator removed from right click menu", , True
'Else
'    modVars.AddToRightClick RightClickExt, RightClickMenuTitle, AppPath() & App.EXEName
'
'    AddText "Communicator added to right click menu", , True
'End If
'
'mnuFileSettingsRMenu.Checked = Not mnuFileSettingsRMenu.Checked
'
'End Sub

Private Sub mnuFileStealth_Click()
StealthMode = True
End Sub

Private Sub mnuOnlineFileTransfer_Click()
Load frmUpload
frmUpload.Show vbModeless, Me
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

'Private Sub mnuOnlineManual_Click()
''ShellExecute 0&, vbNullString, modFTP.UpdateZip, vbNullString, vbNullString, vbNormalNoFocus
'Dim Ret As Long
'Dim Ans As VbMsgBoxResult
'Dim Path As String
'
'Ans = Question("Download via HTTP Protocol?", mnuOnlineManual)
'
'If Ans = vbYes Then
'    Path = AppPath() & "New Communicator.zip"
'
'    AddText "Downloading...", , True
'    Me.Refresh
'
'    Ret = URLDownloadToFile(0, modFTP.UpdateZip, Path, 0, 0)
'
'    If Ret = 0 And Dir$(Path) <> vbNullString Then
'        AddText "Downloaded Successfully", , True
'
'
'        Call ZipFileExtractQuestion(Path)
'    '    Ans = Question("Open Folder?", mnuOnlineManual)
'    '    If Ans = vbYes Then
'    '        On Error Resume Next
'    '        'Shell "explorer.exe " & Left$(Path, InStrRev(Path, "\", , vbTextCompare)), vbNormalFocus
'    '        OpenFolder (vbNormalFocus)
'    '        AddText "Folder Opened", , True
'    '    End If
'        AddText "Download Complete", , True
'    Else
'        AddText "Download Unsuccessful", , True
'    End If
'Else
'    AddText "Download Canceled", , True
'End If
'
'End Sub

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

Public Function CheckForUpdates(Optional ByVal bStealth As Boolean = False) As Boolean

Load frmUpdate
On Error Resume Next 'in case object is unloaded
CheckForUpdates = frmUpdate.Show_Update_Check(bStealth)

End Function

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


If modLoadProgram.bVistaOrW7 = False Then
    mnuOptionsAdvDisplayGlassBG.Checked = False
    mnuOptionsAdvDisplayGlassBG.Enabled = False
    'caption already set, since the OS version won't have changed
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
        
        SetControlsLeft bTmp
        lblBorder.Visible = bTmp
        Call Form_Resize
        
    Else
        mnuOptionsAdvDisplayGlassBG.Checked = False
        mnuOptionsAdvDisplayGlassBG.Enabled = False
        
        SetControlsLeft False
        
        If Not modLoadProgram.bLoading Then
            AddText "Error - Desktop Composition Not Enabled", TxtError, True
        End If
        
    End If
End If

End Sub

Private Sub ActivateGlass()

modDisplay.SetGlassBorders Me.hWnd, , , _
    GetMenuHeight() - GetBorderHeight() / 2 ', _
    ScaleY(sbMain.height, vbTwips, vbPixels) + GetBorderHeight() * 1.5


End Sub

Private Sub SetControlsLeft(bMoveLeft As Boolean)
Dim formLeft As Integer, i As Integer
Static bMovedLeft As Boolean

If bMoveLeft = bMovedLeft Then Exit Sub

formLeft = IIf(bMoveLeft, modDisplay.Glass_Border_Indent, -modDisplay.Glass_Border_Indent)

SetLeft picDraw, formLeft
SetLeft lblName, formLeft
SetLeft txtName, formLeft
SetLeft txtStatus, formLeft
SetLeft lstComputers, formLeft
'SetLeft lstConnected, FormLeft
SetLeft fraDrawing, formLeft
SetLeft fraTyping, formLeft
SetLeft rtfIn, formLeft
SetLeft imgStatus, formLeft
SetLeft txtOut, formLeft
SetLeft picInfo, formLeft

Set_lstConnected_Left

For i = 0 To imgDP.UBound
    SetLeft imgDP(i), formLeft
Next i

For i = 0 To cmdArray.UBound
    SetLeft cmdArray(i), formLeft
    SetLeft cmdXPArray(i), formLeft
Next i

bMovedLeft = bMoveLeft

End Sub

Public Sub Set_lstConnected_Left()

If mnuOptionsAdvDisplayGlassBG.Checked Then
    lstConnected.Left = lstConnectedNormLeft + ScaleX(modDisplay.Glass_Border_Indent, vbPixels, vbTwips)
Else
    lstConnected.Left = lstConnectedNormLeft
End If

End Sub

Private Sub SetLeft(Ctrl As Control, formLeft As Integer)
Ctrl.Left = Ctrl.Left + ScaleX(formLeft, vbPixels, vbTwips)
End Sub

Private Function GetMenuHeight() As Long
'gets menu height IN PIXELS

GetMenuHeight = GetSystemMetrics(SM_CYCAPTION)

End Function

Private Function GetBorderHeight() As Long
GetBorderHeight = GetSystemMetrics(SM_CYFIXEDFRAME)
End Function

Public Sub mnuOptionsAdvDisplayVistaControls_Click()
Dim bDisable As Boolean

If modLoadProgram.bVistaOrW7 = False Then
    mnuOptionsAdvDisplayVistaControls.Checked = False
    mnuOptionsAdvDisplayVistaControls.Enabled = False
Else
    mnuOptionsAdvDisplayVistaControls.Checked = Not mnuOptionsAdvDisplayVistaControls.Checked
    
    If mnuOptionsAdvDisplayVistaControls.Checked Then
        If modDisplay.VisualStyle() Then
            
            If modDisplay.CompositionEnabled() Then
                
                Call SetVistaControls
                
            Else
                'no composition
                Call SetVistaControls(False)
                bDisable = True
                If Not modLoadProgram.bLoading Then
                    AddText "Error - Desktop Composition Not Enabled", TxtError, True
                End If
            End If
            
        Else
            'no visual style
            Call SetVistaControls(False)
            bDisable = True
            If Not modLoadProgram.bLoading Then
                AddText "Error - Visual Styles aren't enabled", TxtError, True
            End If
        End If
    Else
        'just turning off
        Call SetVistaControls(False)
    End If
End If

If bDisable Then
    mnuOptionsAdvDisplayVistaControls.Checked = False
    mnuOptionsAdvDisplayVistaControls.Enabled = False
End If

End Sub

Public Function GetCommandIconHandle() As Long
GetCommandIconHandle = frmSystray.imgButton.ListImages(1).Picture.Handle
End Function

Private Sub SetVistaControls(Optional bEnable As Boolean = True)
Dim hPic As Long
'Dim B As Boolean

'AddConsoleText "SVC Called - bEnable: " & bEnable, , True
'B = modLoadProgram.bVistaOrW7


If bEnable Then
    
    hPic = GetCommandIconHandle()
    
    modDisplay.SetButtonIcon cmdArray(0).hWnd, hPic
    modDisplay.SetButtonIcon cmdArray(1).hWnd, hPic
    modDisplay.SetButtonIcon cmdArray(2).hWnd, hPic
    
    'modDisplay.MakeCommandLink cmdScan
    'modDisplay.MakeCommandLink cmdPrivate
Else
    modDisplay.SetButtonIcon cmdArray(0).hWnd, 0
    modDisplay.SetButtonIcon cmdArray(1).hWnd, 0
    modDisplay.SetButtonIcon cmdArray(2).hWnd, 0
    
    'modDisplay.RemoveCommandLink cmdScan
    'modDisplay.RemoveCommandLink cmdPrivate
End If

'AddConsoleText "Exiting SVC Proc", , , True

'Exit Sub
'EH:
'AddConsoleText "SVC Proc Error - " & Err.Description, , , True
End Sub

Private Sub SetTBBanners(Optional ByVal bSet As Boolean = True)

If bSet Then
    modDisplay.SetTextBoxBanner txtName.hWnd, "Enter Your Name"
    modDisplay.SetTextBoxBanner txtOut.hWnd, "Enter a Message"
    modDisplay.SetTextBoxBanner txtStatus.hWnd, "Your Status Here"
Else
    modDisplay.SetTextBoxBanner txtName.hWnd, vbNullString
    modDisplay.SetTextBoxBanner txtOut.hWnd, vbNullString
    modDisplay.SetTextBoxBanner txtStatus.hWnd, vbNullString
End If

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

If modDisplay.setVisualStyle(Not mnuOptionsAdvDisplayStyles.Checked) Then
    mnuOptionsAdvDisplayStyles.Checked = Not mnuOptionsAdvDisplayStyles.Checked
    If Not Told Then
        AddText "You need to restart this program for changes to take place", TxtError, True
        Told = True
    End If
Else
    AddText "Error creating visual style file - " & Err.Description, TxtError, True
End If

End Sub

Private Sub mnuOptionsDPSet_Click()
Dim Path As String, IDir As String
Dim Er As Boolean
Dim i As Integer, iClient As Integer
Const ImgFilter = "All Images|*.jpeg;*.jpg;*.bmp;*.gif|Bitmap (*.bmp)|*.bmp|Jpeg (*.jpeg,*.jpg)|*.jpeg;*.jpg|GIF (*.gif)|*.gif"

iClient = FindClient(modMessaging.MySocket)

If iClient = -1 Then
    'AddConsoleText "Clients not init'd"
    If Server Then
        SetInfo DP_Error_Server, True
    Else
        SetInfo DP_Error_Clients, True
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
SetInfo "Error - " & Err.Description, True
End Sub

Private Function SetMyDP(ByVal Path As String) As Boolean
Const iMaxFileSize_MB As Integer = 3 'aka 3MB
Const iMaxFileSize As Long = 1024 * iMaxFileSize_MB
Dim iMaxSize As Integer
Dim IDir As String
Dim i As Integer, iClient As Integer
Dim bIsGIF As Boolean

If Not Server Then
    If modMessaging.MySocket = 0 Then
        'shouldn't get here - server sends socket on connect
        
        'AddConsoleText "Socket = 0 - Exiting Sub..."
        
        SetInfo DP_Error_Socket, True
        Exit Function
    End If
End If


iClient = FindClient(modMessaging.MySocket)

If iClient > -1 Then
    If LenB(Path) Then
        
        bIsGIF = (Right$(Path, 3) = "gif")
        
        'check size
        iMaxSize = IIf(bIsGIF, iMaxFileSize * 4, iMaxFileSize)
        If (FileLen(Path) / 1024) > iMaxSize Then
            'KB
            SetInfo "File is too large. It must be smaller than " & CStr(iMaxSize) & "KB", True
            Exit Function
        End If
        
        
        IDir = DP_Dir_Path
        
        If FileExists(IDir, vbDirectory) = False Then
            On Error Resume Next
            MkDir IDir
        End If
        IDir = IDir & "\Local." & IIf(bIsGIF, "gif", "jpg")
        If FileExists(IDir) Then
            On Error Resume Next
            Kill IDir
        End If
        
        FileCopy Path, IDir
        
        
        Set Clients(iClient).IPicture = LoadPicture(IDir)
        Clients(iClient).bDPIsGIF = bIsGIF
        
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
        mnuOptionsDPClear.Enabled = True
        
        
        SetInfo "Loaded Picture (" & GetFileName(Path) & ")", False
        'SendInfoMessage LastName & " set their display picture", , True
        
    End If
Else
    'AddConsoleText "Clients not init'd"
    If Server Then
        SetInfo DP_Error_Server, True
    Else
        SetInfo DP_Error_Clients, True ', TxtError, True
    End If
End If

End Function

Private Sub mnuOptionsDPClear_Click()
Dim i As Integer
Dim sPath As String

mnuOptionsDPClear.Enabled = False


If modMessaging.MySocket <> 0 Then
    i = FindClient(modMessaging.MySocket)
    
    If i > -1 Then
        With Clients(i)
            .bDPSet = False
            .bSentHostDP = False
            '.sHasiDPs = vbNullString 'Keep other's DPs
            Set .IPicture = Nothing
        End With
        
        If Not Server Then
            'yo server, we don't have a DP
            SendData eCommands.SetClientVar & eClientVarCmds.SetDPSet & "0"
        End If
        
        
        ResetImgDP 0
        
        
        'delete DP
        If modDP.DP_Path_Exists() Then
            On Error Resume Next
            Kill modDP.My_DP_Path
            modDP.My_DP_Path = vbNullString
        End If
        
        SetInfo "Display Picture will disappear shortish", False
    Else
        If Server Then
            SetInfo DP_Error_Server, True
        Else
            SetInfo DP_Error_Clients, True ', TxtError, True
        End If
    End If
Else
    SetInfo DP_Error_Socket, True
End If

End Sub

Private Sub mnuOptionsDPSaveAll_Click()
mnuOptionsDPSaveAll.Checked = Not mnuOptionsDPSaveAll.Checked
End Sub

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
        .ToolTipText = vbNullString
    End With
    
    j = j + 1
Loop
j = j - 1

'##################################################################

If Not (Clients(iClient).IPicture Is Nothing) Then
    
    Set imgDP(j).Picture = Clients(iClient).IPicture
    
    If imgDP(j).BorderStyle <> 1 Then imgDP(j).BorderStyle = 1
    
    
'    If LenB(Clients(iClient).sName) Then
'        sTxt = FormatApostrophe(Clients(iClient).sName) & " Display Picture"
'        If imgDP(j).ToolTipText <> sTxt Then
'            imgDP(j).ToolTipText = sTxt
'        End If
'    End If
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
    'If modLoadProgram.bSafeMode Then
        'SetInfo "Error - Can't Open Game Lobby in Safe Mode", True
    'Else
        Load frmLobby
        'frmLobby.Show vbModeless, Me
    'End If
Else
    SetInfo "Error - A Game Window is Open", True
    
    If modSpaceGame.GameFormLoaded Then
        SetFocus2 frmGame
    Else 'if modStickGame.StickFormLoaded then
        SetFocus2 frmStickGame
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
            SetFocus2 Frm
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

Private Sub mnuOptionsVoice_Click()
Load frmSpeech
frmSpeech.Show vbModeless, Me
End Sub

Public Sub mnuSBCopylIP_Click()

On Error Resume Next
With Clipboard
    .Clear
    .SetText modWinsock.LocalIP
    Beep
End With

End Sub

Public Sub mnuSBCopyrIP_Click()

On Error Resume Next
With Clipboard
    .Clear
    .SetText modWinsock.RemoteIP
    Beep
End With

End Sub

Public Sub mnuSBObtain_Click()
'sbMain.Panels(1).Text = "Obtaining External IP..."
'modVars.lIP = Me.SckLC.LocalIP

SetPanelText "Obtaining External IP...", 1

If modWinsock.ObtainRemoteIP() Then
    If modLoadProgram.bSlow Then modLogin.AddToFTPList
    SetInfo "Obtained External IP", False
Else
    AddText "Error Obtaining External IP - Are you connected to the internet?", TxtError, True
End If

ShowSB_IP

modVars.GetTrayText '+ set
End Sub

Public Sub mnuSBObtainLocal_Click()
'sbMain.Panels(1).Text = "Obtaining External IP..."
'modVars.lIP = Me.SckLC.LocalIP

SetPanelText "Obtaining Internal IP...", 1

modWinsock.ObtainLocalIP
SetInfo "Obtained Internal IP", False

ShowSB_IP

modVars.GetTrayText '+ set
End Sub

Private Sub picColour_Change()
picColours(7).BackColor = picColour.BackColor
picColour.BackColor = colour
End Sub

Private Sub picColours_Click(Index As Integer)
If Index <> 7 Then
    colour = picColours(Index).BackColor
    picColour_Change
    picColour.BackColor = picColours(Index).BackColor
Else
    colour = picColours(7).BackColor
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
    PopupMenu frmSystray.mnuSB, , , , frmSystray.mnuSBCopyrIP
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

If mnuOptionsAdvDisplayConnF(1).Checked Then
    cmdAdd_Proper_Click
Else
    mnuFileManual_Click
End If

End Sub
Private Sub cmdAdd_Proper_Click()
If Connect() Then
    'AddText "To save time, double click the listbox instead", TxtError, True
    SetInfo "To save time, double click the listbox instead", True
End If
End Sub

Public Function Reconnect() As Boolean

If LenB(modMessaging.LastIP) Then
    SetInfo "Reconnecting to " & modMessaging.LastIP, False
    Reconnect = Connect(modMessaging.LastIP)
Else
    SetInfo "Nowhere to reconnect to...", True
    Reconnect = False
End If

End Function

Public Function Connect(Optional ByVal Name As String = vbNullString) As Boolean
Dim sRemoteHost As String, Text As String
Dim i As Integer

AddConsoleText "Beginning Connecting...", , True, , True

Connect = True


On Error GoTo EH
'txtName_LostFocus


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
    modLogging.addToActivityLog "Connecting to " & sRemoteHost
    
    modMessaging.LastIP = sRemoteHost
    modMessaging.AddUsedIP sRemoteHost
    
    modMessaging.CurIPIndex = -1
    
    For i = 0 To UBound(UsedIPs)
        If UsedIPs(i).sIP = sRemoteHost Then
            CurIPIndex = i
            Exit For
        End If
    Next i
    
    
    SckLC.RemotePort = MainPort
    SckLC.LocalPort = 0 'LPort
    
    Text = "Connecting to " & sRemoteHost & ":" & MainPort & "..."
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
Dim n As Integer
Dim Frm As Form

If bCleanedUp Then Exit Sub

AddConsoleText "Cleaning Up...", , True ', , True

'txtOut.Text = vbNullString
'txtOut_Change
Pause 1

If SckLC.state <> sckClosed Then
    SckLC_Close 'we close it in case it was trying to connect or whatever
Else
    'autosave the picture
    If SavePic And pDrawDrawnOn Then
        Call SaveLastPic
    End If
    'end autosave
End If


'autosave convo if needed
If Me.mnuOptionsMessagingLoggingAutoSave.Checked Then AutoSave True
If Me.mnuOptionsMessagingLoggingConv.Checked Then DoLog True
'If Me.mnuOptionsMessagingLoggingPrivate.Checked Then LogPrivate


Inviter = vbNullString
Server = False 'must be after scklc_close
SendTypeTrue = False 'for typingstr
SendTrueDraw = False 'for drawingstr

mnuOptionsMessagingPrivate.Caption = "Private Chat with..."

lstConnected.Clear
cmdArray(3).Enabled = False
cmdXPArray(3).Enabled = False


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
For n = 0 To imgDP.UBound
    
    ResetImgDP n
    
    If CBool(n) Then
        Unload imgDP(n)
    End If
Next n

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
If modDev.bDevCmdFormLoaded Then
    Unload frmDevCmd
End If


'mnuOptionsMessagingDrawingOff.Checked = False

'close and unload all previous sockets
For n = 1 To SockAr.UBound '(SockAr.Count - 1)
    If ControlExists(SockAr(n)) Then
        On Error Resume Next
        SockAr(n).Close
        Unload SockAr(n)
    End If
Next n

Cmds Idle

SocketCounter = 0
iSelectedClientSock = 0
modMessaging.MySocket = 0

AddConsoleText "Cleaned Up", , , True

modLogging.addToActivityLog "Connections Reset"


'frmSystray.ShowBalloonTip "All Connections Closed", "Communicator", NIIF_INFO
If modVars.nPrivateChats > 0 Then
    For Each Frm In Forms
        If Frm.Name = frmPrivateName Then
            Unload Frm
        End If
    Next Frm
End If

Me.ucFileTransfer.Disconnect
Me.ucVoiceTransfer.Disconnect
bRecordingTransferInProgress = False
If pbRecording Then
    'stop eet
    modSoundCap.Stop_Recording recording_Name
    modSoundCap.Close_Record recording_Name
End If
bRecording = False
recording_Remote_Filename = vbNullString
recording_File_To_Send = vbNullString


Unload frmManualFT
picClient.Visible = False

modDP.DelPics

'SetFocus2 txtOut 'NOOOOOOOAAAA - STEALS FOCUS = GHEY

bCleanedUp = True

End Sub

Public Function GetSelectedClient() As Integer
GetSelectedClient = iSelectedClientSock
End Function

Public Sub SaveLastPic(Optional ByVal bTell As Boolean = True)
Dim FilePath As String

'FilePath = GetLogPath() & "Drawings\"
FilePath = GetCurrentLogFolder() 'GetLogPath() & MakeDateFile() & "\"

If FileExists(FilePath, vbDirectory) = False Then
    On Error Resume Next
    MkDir FilePath
End If

FilePath = FilePath & MakeTimeFile() & ".bmp" 'always saves in .bmp

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
Dim sTime As String
sTime = " (" & FormatDateTime$(Time$, vbLongTime) & ")"

If modVars.bRetryConnection Then
    modVars.bRetryConnection = False
    AddText "Canceled Auto-Connect Attempt" & sTime, TxtError, True, True
Else
    'AddText "Connection Closed" & sTime, , True, True
    'text added by scklc_close
End If

Call CleanUp(True)

'reset host timer
tmrHost.Enabled = False
tmrHost.Enabled = True

End Sub

Private Sub cmdCls_Click()
Dim Ans As VbMsgBoxResult
Dim Msg As String

Ans = Question("Clear Board, Are You Sure?", IIf(modLoadProgram.bVistaOrW7, cmdNormCls, cmdXPCls))

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
        EnableCmdCls True
    Else
        EnableCmdCls False
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
    ByVal Text As String, ByVal bHide As Boolean)

', Optional ByVal SocketSendTo As Integer = -1)
Dim dMsg As String

If iCmd Then
    modDev.AddDevLog "Sent Command" & vbNewLine & _
                     "    Command: " & modDev.GetDevCommandName(iCmd) & vbNewLine & _
                     "    To: " & SendTo & vbNewLine & _
                     "    Parameter(s): " & Text
    
    'format = SendToName # FromName @ Command Parameter [[1|0]OVERRIDE]
    
    'dMsg = eCommands.DevSend & SendTo & "#" & Trim$(LastName) & "@" & iCmd & _
                Text & IIf(Override, IIf(bHide, "1", "0") & modDev.DevOverride, vbNullString)
    
    dMsg = modDev.createDevCommand(SendTo, Trim$(LastName), bHide, CStr(iCmd) & Text)
    
    AddDevText "Sent (to " & SendTo & ") - Command: " & GetDevCommandName(iCmd) & ", Parameter: " & Text, True
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

SckLC.Close
SckLC.LocalPort = MainPort
SckLC.RemotePort = 0 'LPort

On Error Resume Next
SckLC.Listen

If SckLC.state <> sckListening Then GoTo EH

Cmds Listening

AddText ListeningStr, , True
modLogging.addToActivityLog "Listening"

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

Private Sub cmdXPReply_Click(Index As Integer)
cmdReply_Click Index
End Sub

Private Sub cmdReply_Click(Index As Integer)
QuestionReply = Index
cmdReply(0).Visible = False
cmdReply(1).Visible = False
cmdXPReply(0).Visible = False
cmdXPReply(1).Visible = False
End Sub

Private Sub cmdShake_Click()

Dim timeToShake As Long

If LastShake = 0 Then
    On Error Resume Next 'just incase GTC = largest -ve value
    
    LastShake = GetTickCount() - Shake_Delay - 10
End If

timeToShake = LastShake + Shake_Delay - GetTickCount()

If timeToShake > 0 Then
    SetInfo CStr(Round(timeToShake / 1000, 1)) & " seconds until you can shake again", True
Else
    If Server Then
        DistributeMsg eCommands.Shake & LastName, -1
    Else
        SendData eCommands.Shake & LastName
    End If
    AddText "Shake Sent by " & LastName, TxtSent, True
    
    LastShake = GetTickCount()
End If

SetFocus2txtOut

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'all keypress events come through here <= me.keypreview = True

If KeyCode = 223 Then
    If Shift = 1 Then
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
    #If HOOK_TXTOUT Then
        If modSubClass.txtOut_Hooked() Then
            modSubClass.txtOut_Hooked = False
        End If
    #End If
    
    If Not App.PrevInstance Then
        If modSpeech.sHiBye Then
            If modSpeech.sBye Then 'start saying here - terminated in _Unload
                
                If Me.mnuOptionsMessagingDisplaySysUserName.Checked Then
                    Tmp = modVars.User_Name
                Else
                    Tmp = Me.LastName
                End If
            
                modSpeech.Say "Goodbye " & Tmp
            End If
        End If
    End If
    
    modNetwork.InitBandwidthStuff False
    
    If ConsoleShown Then
        ShowConsole False
    End If
    
    If modSpaceGame.GameFormLoaded Then
        frmGame.bRunning = False
    End If
    If modStickGame.StickFormLoaded Then
        frmStickGame.bRunning = False
    End If
    
    If frmSystray.mnuStatusAway.Checked Then
        mnuStatusAway_Click
    End If
    
    'If mnuFileSaveExit.Checked Then
    If mnuFileSettingsUserProfileExportOnExit.Checked Then
        SaveUserProfileSettings 'file
        SaveUsedIPs
        SaveSettings 'reg
    End If
    
    If InTray Then
        DoSystray False
    End If
    
    Set HTML_Doc = Nothing
    
    AddConsoleText "Goodbye!", , , True
    modLogging.addToActivityLog "Exiting" & vbNewLine
    
    'If Me.Visible Then ImplodeFormToMouse Me.hwnd 'done after all others closed
Else
    Cancel = True
    
    If frmUDP.Visible Then Unload frmUDP
    '                      only actually unload is modVars.bClosing = T
    
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
    
    ShowCmdReplys
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

ShowCmdReplys

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

ShowCmdReplys False

Caller.Enabled = True
Questioning = False

End Function

Private Sub Form_Unload(Cancel As Integer)
Dim Frm As Form

For Each Frm In Forms
    
    If Frm.Name <> "frmMain" Then Unload Frm
    
    'Set Frm = Nothing
    
    Me.Refresh
Next Frm

Me.Refresh

If Me.Visible Then modImplode.AnimateAWindow hWnd, aRandom, True 'ImplodeFormToMouse Me.hWnd
Me.Hide
'modLogging.LogEvent "Unloaded Main Window"

modLogin.RemoveFromFTPList

modAudio.StopSound
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

Private Sub lstConnected_DblClick()
Dim iSock As Integer, iClient As Integer

On Error GoTo EH

iSock = lstConnected.ItemData(lstConnected.ListIndex)
iClient = FindClient(iSock)
If iClient > -1 Then
    With Clients(iClient)
        If LenB(.sIP) Then
            Clipboard.Clear
            Clipboard.SetText .sIP
            SetInfo "IP (" & .sIP & ") Copied to Clipboard", False
        Else
            SetInfo "IP Unknown", True
        End If
    End With
End If


EH:
End Sub

Private Sub lstConnected_Click()
Dim sName As String
Dim iSock As Integer, iClient As Integer
Const kCaption = "Private Chat with..."
Dim bEnable As Boolean

On Error Resume Next
iSock = lstConnected.ItemData(lstConnected.ListIndex)
sName = Trim$(lstConnected.Text)


If iSock <> 0 And iSock <> modMessaging.MySocket Then
    'If Tmp <> ConnectedListPlaceHolder Then
    
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



bEnable = CBool(iSock)

mnuOptionsMessagingPrivate.Enabled = bEnable
EnableCmd 5, bEnable
EnableCmd 3, bEnable And Server

End Sub

Public Sub mnuDevClient_Click()
'don't remove this, called by mnufileclient
frmClients.Show vbModeless, Me
End Sub

'Private Sub mnuConsoleOff_Click()
'ShowConsole False
'End Sub
'
'Private Sub mnuConsoleType_Click()
'
'Static Told As Boolean
'
'If Not Told Then
'    AddText "Type into the Console", , True
'    rtfIn.Refresh
'    Me.Refresh
'    Told = True
'End If
'
'modConsole.ProcessConsoleCommand
'
'End Sub
'
'Private Sub mnuConsoleTypeLots_Click()
'
'Static Told As Boolean
'
'If Not Told Then
'    AddText "Type into the Console", , True
'    AddConsoleText "For assistance, type help"
'    rtfIn.Refresh
'    Me.Refresh
'    Told = True
'End If
'
'modConsole.ProcessConsoleCommand True
'
'End Sub

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
frmDev.ShowForm True
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
        "-----", DevCol1, False

End Sub

Private Sub mnuDevMaintenanceTimers_Click()
Dim b As Boolean
Const kTxt = "You should restart me to get everything back to normal"
'tmrMain.Enabled = False
'tmrHost.Enabled = False
'tmrShake.Enabled = False
''tmrInactive.Enabled = False
'tmrLog.Enabled = False

b = mnuDevMaintenanceTimers.Checked
mnuDevMaintenanceTimers.Checked = Not b

DisableAllTimers b

AddText kTxt, TxtError, True
SetInfo kTxt, False
End Sub

Private Sub DisableAllTimers(Optional bEnable As Boolean = False)
Dim Tmr As Control

For Each Tmr In Controls
    If TypeOf Tmr Is Timer Then
        Tmr.Enabled = bEnable
    End If
Next Tmr

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
If modVars.OpenNewCommunicator("/forceopen") = False Then
    AddText "Error - " & Err.Description, TxtError, True
End If
End Sub

Private Sub mnuFileRefresh_Click()
Dim sError As String

bHadNetworkRefreshError = False 'allow it to try

If RefreshNetwork(sError) Then
    AddText "Refreshed Network List", , True
Else
    AddText "Error - " & sError, TxtError, True
End If
End Sub

Public Sub mnuFileSaveCon_Click()
mnuRtfPopupSaveAs_Click
End Sub

Private Sub mnuFileSaveDraw_Click()
Dim Path As String, IDir As String
Dim Er As Boolean
'Dim i As Integer

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
        
        
        AddText "Saved Drawing (" & GetFileName(Path) & ")", , True
    End If
End If

Exit Sub
EH:
AddText "Error Saving Drawing: " & Err.Description, , True
End Sub

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

SetInfo "Manual Options Configured", False

End Sub

Private Sub mnuOptionsAdvPresetReset_Click()

'Me.mnuFileSaveExit.Checked = True
'-
'Call AnimClick(Me.mnuOptionsWindow2All)
mnuOptionsWindow2Animation.Checked = True
Me.mnuOptionsWindow2SingleClick.Checked = False
'Me.mnuOptionsBalloonMessages.Checked = True
mnuOptionsAlertsStyle_Click 0
'-
Me.mnuOptionsTimeStamp.Checked = False
Me.mnuOptionsTimeStampInfo.Checked = False
Me.mnuOptionsFlashMsg.Checked = True
'Me.mnuOptionsFlashInvert.Checked = False
'Me.mnuOptionsMessagingColours.Checked = True
Me.mnuOptionsMessagingDisplaySmiliesComm.Checked = True
Me.mnuOptionsMessagingDisplaySmiliesMSN.Checked = True
ApplySmileySettings

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

'modSpeech.Vol = 100
'modSpeech.Speed = 0

SetInfo "Reset to Original Settings", False

End Sub

Private Sub mnuOptionsAdvPresetServer_Click()

'Me.mnuOptionsAdvInactive.Checked = True
Me.mnuOptionsHost.Checked = True
Me.mnuOptionsStartup.Checked = True
mnuOptionsAdvHostMin.Checked = True

'AddText "Server Options Configured", , True
SetInfo "Server Options Configured", False

End Sub

Private Sub mnuOptionsFlashMsg_Click()
mnuOptionsFlashMsg.Checked = Not mnuOptionsFlashMsg.Checked
'mnuOptionsFlashInvert.Enabled = mnuOptionsFlashMsg.Checked
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
    SetInfo "Type in this textbox", False
    Told = True
End If

AddText vbNullString 'Will add new line

If mnuOptionsMatrix.Checked Then SetFocus2 rtfIn

End Sub

Public Sub mnuOptionsMessagingClearTypeList_Click()
Dim i As Integer
lblTyping.Caption = vbNullString

modMessaging.TypingStr = vbNullString
modMessaging.DrawingStr = vbNullString

ReDim modMessaging.Typers(0)
ReDim modMessaging.Drawers(0)

End Sub

Private Sub mnuOptionsMessagingLoggingConv_Click()
mnuOptionsMessagingLoggingConv.Checked = Not mnuOptionsMessagingLoggingConv.Checked
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
ElseIf Ans = vbNo Then
    AddText "Clear Text Canceled", , True
End If
'cmdCls_Click

End Sub

Public Sub ClearRtfIn()

DoLog True
rtfIn.Text = vbNullString
currentLogFile = vbNullString

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

Private Sub mnuRtfPopupSaveAs_Click()
Dim Path As String, IDir As String
Dim Er As Boolean

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
        
        
        AddText "Saved Conversation (" & GetFileName(Path) & ")", , True
        
    End If
End If

End Sub

Public Sub CommonDPath(ByRef Path As String, ByRef Er As Boolean, _
    ByVal Title As String, Optional ByVal Filter As String = _
    "Rich Text Format (*.rtf)|*.rtf|Text File (*.txt)|*.txt", _
    Optional ByVal InitDir As String = vbNullString, _
    Optional ByVal OpenFile As Boolean = False) ', _
    Optional ByVal bDirChangable As Boolean = True)

Dim TmpPath As String

'API version: http://vbnet.mvps.org/index.html?code/hooks/fileopensavedlghooklvview.htm

Er = False

If LenB(InitDir) = 0 Then
    If modLoadProgram.bVistaOrW7 Then
        InitDir = Environ$("USERPROFILE")
    Else
        InitDir = Environ$("USERPROFILE") & "\My Documents"
    End If
End If

Cmdlg.Filter = Filter
'Cmdlg.FilterIndex = 2 'Start at the second filter

Cmdlg.DialogTitle = Title
Cmdlg.CancelError = True


If LenB(Path) Then
    On Error Resume Next
    Path = Right$(Path, Len(Path) - InStrRev(Path, "\", , vbTextCompare))
    'just the filename
End If

Cmdlg.fileName = Path 'vbNullString
Cmdlg.InitDir = InitDir
Cmdlg.flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or _
    cdlOFNFileMustExist Or cdlOFNOverwritePrompt ' _
    Or IIf(bDirChangable, 0, cdlOFNNoChangeDir)


On Error GoTo CancelError
If OpenFile Then
    Cmdlg.ShowOpen
Else
    Cmdlg.ShowSave
End If

TmpPath = Cmdlg.fileName

If LenB(TmpPath) Then
    Path = Trim$(TmpPath)
Else
    Path = vbNullString
End If

Exit Sub
CancelError:
If Err.Number = cdlCancel Then
    Er = True
Else
    MsgBoxEx "Error: " & Err.Description, "A random error occured showing the 'Save As'/'Open' dialog", _
        vbExclamation, "Error"
    
    Er = True
End If
End Sub

Private Sub PicColour_Click()

Cmdlg.flags = cdlCCFullOpen + cdlCCRGBInit
Cmdlg.Color = colour

On Error GoTo Err
Cmdlg.ShowColor

colour = Cmdlg.Color
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
        
        
        Ans = Question("Kick " & sTargetToDisplay & _
            IIf(LenB(Clients(iRemove).sIP), " (" & Clients(iRemove).sIP & ")", vbNullString) & _
            ", are you sure?", IIf(modLoadProgram.bVistaOrW7, cmdArray(3), cmdXPArray(3)))
        
        
        If Ans = vbYes Then
            Call Kick(iRemove, sTarget)
        Else
            AddText "Kick Canceled", , True
        End If
    End If
    
Else
    cmdXPArray(3).Enabled = False
    cmdArray(3).Enabled = False
    AddText "Only the server/host can remove people", TxtError, True
End If

End Sub

Public Sub Kick(ByVal iSocket As Integer, ByVal sTarget As String, Optional ByVal bTell As Boolean = True, _
    Optional ByVal ConnectAttempt As Boolean = False)
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
    If ConnectAttempt Then
        Str = "'" & sTarget & "' attempted to connect - rejected and kicked"
    Else
        Str = "'" & sTarget & "' was kicked"
    End If
    
    'modMessaging.DistributeMsg eCommands.Info & Str & "1", -1
    SendInfoMessage Str, False, True, , , iSocket
    SendInfoMessage "You were kicked from the server", , True, , iSocket
    Pause 100
    
    AddText Str, TxtError, True
End If

On Error Resume Next
'sockAr_Close iRemove
sockClose iSocket, bTell

DataArrival eCommands.Typing & "0" & sTarget
DataArrival eCommands.Drawing & "0" & sTarget

End Sub

Private Function CountNewLines(ByVal sTxt As String) As Integer
Dim i As Integer, j As Integer

For i = 1 To Len(sTxt)
    If Mid$(sTxt, i, 2) = vbNewLine Then
        j = j + 1
    End If
Next i

CountNewLines = j

End Function

Private Function RemoveMessageSeps(ByVal sTxt As String) As String

If InStr(1, sTxt, modMessaging.MessageSeperator) Then
    sTxt = Replace$(sTxt, modMessaging.MessageSeperator, vbNullString)
End If

If InStr(1, sTxt, modMessaging.MessageStart) Then
    sTxt = Replace$(sTxt, modMessaging.MessageStart, vbNullString)
End If

RemoveMessageSeps = sTxt

End Function

Private Function processIRCCommand(sText As String) As String
Dim sCmd As String, sExtra As String
Dim i As Integer

i = InStr(1, sText, vbSpace)
If i Then
    sCmd = Left$(sText, i - 1)
    sExtra = Trim$(Mid$(sText, i + 1))
Else
    sCmd = sText
End If

Select Case LCase$(sCmd)
'-------------------------------------------------------------------------------------------------
    'Chat IRC commands return empty string on failure, string to add to the log on success
    
    Case "me"
        If LenB(sExtra) Then
            processIRCCommand = InfoStart & LastName & vbSpace & sExtra & InfoEnd
            'use trim, in case they don't have a space after /me
        Else
            SetInfo "You what? You need an action to do to someone", True
            Beep
            processIRCCommand = vbNullString
            Exit Function
        End If
        
        
    Case "agree"
        If LenB(sExtra) Then
            processIRCCommand = InfoStart & Trim$(LastName & " agrees: " & sExtra) & InfoEnd
        Else
            processIRCCommand = InfoStart & Trim$(LastName & " agrees") & InfoEnd
        End If
        
        
    Case "summon"
        processIRCCommand = InfoStart & Trim$("Summoning " & String2(20, "soy ") & vbSpace & sExtra) & InfoEnd
        
        
    Case "describe"
        If LenB(sExtra) Then
            processIRCCommand = InfoStart & sExtra & InfoEnd
        Else
            SetInfo "You need an action - Slap someone with a fish?", True
            processIRCCommand = vbNullString
            Exit Function
        End If
        
'-------------------------------------------------------------------------------------------------
    'Non chat IRC commands return empty string on success, non-empty on failure
    'Should return empty string, since that'll exit cmdSend_Click immediately
    
    Case "connect"
        If LenB(sExtra) Then
            modVars.bRetryConnection = False
            Connect sExtra
            txtOut.Text = vbNullString
            SetFocus2 txtOut
        Else
            SetInfo "Enter an IP to connect to, e.g. /connect " & modMessaging.LastIP, True
        End If
        processIRCCommand = vbNullString
        
    Case "disconnect"
        If Status <> Idle Then
            CleanUp True
            txtOut.Text = vbNullString
            SetFocus2 txtOut
        Else
            SetInfo "Error - Already Disconnected", True
            Beep
            txtOut.Text = vbNullString
        End If
        processIRCCommand = vbNullString
        
        
    Case "reconnect"
        If Reconnect() Then
            txtOut.Text = vbNullString
            SetFocus2 txtOut
        End If
        processIRCCommand = vbNullString
        
'-------------------------------------------------------------------------------------------------
    Case Else
        If Status = Connected Then
            SetInfo "Command not recognised: /" & sCmd & " - sending plain text", True
            processIRCCommand = LastName & MsgNameSeparator & "/" & sText
        Else
            'shouldn't get here, since cmdSend will be disabled
            SetInfo "Command not recognised: /" & sCmd, True
            processIRCCommand = vbNullString
        End If
        
End Select

End Function

Public Sub cmdSend_Click()
Dim StrOut As String
Dim colour As Long
Dim txtOutText As String
Dim sDataToSend As String
Dim sTmp As String
Dim sFont As String
Dim i As Integer
Static LastTick As Long


If LCase$(LastStatus) = LCase$(AFKStr) Then
    'toggle
    mnuStatusAway_Click
    SetInfo "Auto-Removed AFK Status", False
End If

On Error GoTo EH
If LastTick = 0 Then
    On Error Resume Next 'just incase GTC = largest -ve value
    
    LastTick = GetTickCount() - MsMessageDelay - 10
End If

sTmp = TrimNewLine$(Trim$(txtOut.Text))

If (Not modMessaging.bReceivedWelcomeMessage And Not Server) Or (Status <> Connected) Then
    If Left$(sTmp, 1) = "/" Then
        If LenB(processIRCCommand(Mid$(sTmp, 2))) = 0 Then
            'processed a non-chat command, exit, since we're not connected
            Exit Sub
        End If
    End If
    
    SetInfo "Can't send yet - Waiting for welcome message", True
    Beep
    Exit Sub
End If

trySend:

If (LastTick + MsMessageDelay) < GetTickCount() Then
    
    txtOutText = RemoveMessageSeps(sTmp)
    If txtOutText <> sTmp Then
        AddText "Certain Characters have been removed", TxtError, True
    End If
    sTmp = vbNullString
    
    
    If CountNewLines(txtOutText) > NewLineLimit Then
        
        SetInfo "Too many lines! Reduce them. (Limit is " & NewLineLimit & ")", True
        Beep
        
        Exit Sub
    End If
    
    
    If (Right$(txtOutText, 1) = "/") And Me.mnuOptionsMessagingReplaceQ.Checked Then
        txtOutText = Left$(txtOutText, Len(txtOutText) - 1) & "?"
    End If
    
    
    If LenB(txtOutText) Then
        
        If Left$(txtOutText, 1) = "/" Then
            StrOut = processIRCCommand(Mid$(txtOutText, 2))
            sFont = DefaultFontName
        Else
            StrOut = LastName & MsgNameSeparator & txtOutText
            sFont = rtfFontName
        End If
        
        'txtOut.height = TextBoxHeight
        txtOut.Text = vbNullString
        
        Pause 100 'otherwise, below data gets sent with 30Name, and it doesn't get received
        
        colour = txtOut.ForeColor 'is changed by mnuoptionsthing_click, to be either txtsent or txtforecolour
        
        If mnuOptionsMessagingEncrypt.Checked = False Then
            sDataToSend = colour & "#" & StrOut
        Else
            sDataToSend = colour & "#" & modMessaging.MsgEncryptionFlag & CryptString(StrOut)
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
                    sTmp = "[" & FormatDateTime$(Time$, vbLongTime) & "] " & LastName
                Else
                    sTmp = LastName
                End If
                
                AddText sTmp & MsgNameSeparator & vbNewLine & Space$(4) & txtOutText, colour, , True, sFont
                
            Else
                If frmMain.mnuOptionsTimeStamp.Checked Then
                    sTmp = "[" & FormatDateTime$(Time$, vbLongTime) & "] " & StrOut
                Else
                    sTmp = StrOut
                End If
                
                AddText sTmp, colour, , True, sFont
                
            End If
            
            If modSpeech.sSent Then
                modSpeech.Say LastName & MsgNameSeparator & txtOutText
            End If
            
        End If
        Pause 50
    Else
        'AddText "Type something to send...", TxtError, True
        SetInfo "Type something to send...", True
        txtOut.Text = vbNullString
    End If
    LastTick = GetTickCount()
Else
    'AddText "No one likes a Spammer - Wait at least half a second", TxtError, True
    SetInfo "No one likes a Spammer - Wait at least half a second", True
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

Private Sub txtStatus_Change()
Dim Txt As String

On Error GoTo EH

If Screen.ActiveControl.Name = txtStatus.Name Then
    Txt = Trim$(txtStatus.Text)
    
    If Txt <> LastStatus Then
        If LenB(Txt) Then
            modDisplay.ShowBalloonTip txtStatus, "Status Setting", "Click off the textbox to set your status"
        Else
            modDisplay.ShowBalloonTip txtStatus, "Status Setting", "Click off the textbox to remove your status"
        End If
    End If
End If

EH:
End Sub

Private Sub txtStatus_KeyPress(KeyAscii As Integer)
Dim sStatus As String
Static bTold As Boolean


If KeyAscii = 13 Then
    KeyAscii = 0
    txtStatus_LostFocus
    
ElseIf KeyAscii <> 8 Or bTold Then
    
    sStatus = txtStatus.Text & Chr$(KeyAscii)
    
    If TruncateStatus(sStatus) <> sStatus Then
        'too long
        SetInfo "Status is too long - shorten it", True
        bTold = True
    ElseIf bTold Then
        SetInfo "That's more like it", False
        bTold = False
    End If
End If

End Sub

Public Sub ResetFocus()
SetFocus2 txtOut
End Sub

'############################################################################################
'############################################################################################
'############################################################################################

Private Sub ucInactiveTimer_UserInactive()
'fires after 1 minute of inactivity
Static iCount As Long
Static bChecked_For_Updates As Boolean


If LenB(modWinsock.RemoteIP) = 0 Then
    If modWinsock.ObtainRemoteIP() Then
        ShowSB_IP
        modVars.GetTrayText
    End If
End If

If Not bChecked_For_Updates Then
    'iCount = 0
    'Load frmUpdate
    bChecked_For_Updates = CheckForUpdates(True)
End If


If mnuOptionsAdvInactive.Checked Then
    iCount = iCount + 1
    
    If iCount >= lInactiveInterval Then
        iCount = 0
        If Status <> Connected Then
            If Me.Visible Then
                ShowForm False
                
                AddText "Window Hidden - You've been inactive for " & _
                    CStr(ucInactiveTimer.InactiveInterval / 1000) & " seconds.", , True
                
            End If
            
            
            If Me.mnuOptionsHost.Checked Then
                If Status <> Listening Then
                    Listen False
                End If
            End If
        Else
            If LenB(LastStatus) = 0 Then
                
                AddText "You've been inactive for " & _
                    CStr(ucInactiveTimer.InactiveInterval / 1000) & " seconds. AFK Status Set", , True
                
                SendInfoMessage LastName & " has been AFK for " & _
                    CStr(ucInactiveTimer.InactiveInterval / 1000) & " seconds"
                
                modVars.bDisableAddText = True
                mnuStatusAway_Click
                modVars.bDisableAddText = False
            End If
        End If
    End If
End If


End Sub

Private Sub mnuOptionsAdvInactive_Click()
Dim b As Boolean
Dim iMins As Long
Dim sInterval As String

b = Not mnuOptionsAdvInactive.Checked


If b Then
    sInterval = modVars.Password("Enter a time to wait for (in minutes)", Me, "Inactive Timer", _
        CStr(lInactiveInterval), False, , True)
    
    If LenB(sInterval) Then
        iMins = val(sInterval)
        
        
        If SetInactiveInterval(iMins) Then
            SetInfo "Delay set to " & CStr(iMins) & " minute" & IIf(iMins > 1, "s", vbNullString), False
            mnuOptionsAdvInactive.Checked = True
        Else
            SetInfo "Delay must be between 1 minute and 5 minutes", True
            mnuOptionsAdvInactive.Checked = False
        End If
    End If
    
Else
    mnuOptionsAdvInactive.Checked = False
End If

End Sub

Public Function SetInactiveInterval(ByVal iMins As Integer) As Boolean

If 1 <= iMins And iMins <= 5 Then
    lInactiveInterval = iMins '* 60000
    SetInactiveInterval = True
    'mnuOptionsAdvInactive.Checked = CBool(iMins)
End If

End Function
Public Property Get Inactive_Interval() As Long
Inactive_Interval = lInactiveInterval
End Property

'############################################################################################
'############################################################################################
'############################################################################################

Private Sub InitVars()

Dim i As Integer
Dim bVisualStyles As Boolean, bComp_En As Boolean
Const Vista_Only_Cap = " (Vista Only)", _
      Desk_Comp_Cap = " (Desktop Composition is Off)", _
      XP_Only_Cap = " (XP Only)"

Dim sTmp As String

modLoadProgram.SetSplashInfo "Disabling/Enabling Certain Menus..."

lblBorder.width = ScaleX(modDisplay.Glass_Border_Indent, vbPixels, vbTwips)
lblTyping.Caption = vbNullString

lstComputers.ZOrder vbSendToBack



bVisualStyles = modDisplay.VisualStyle()
If modLoadProgram.bVistaOrW7 Then
    
    bComp_En = modDisplay.CompositionEnabled()
    mnuOptionsAdvDisplayGlassBG.Enabled = bComp_En
    
    If bComp_En Then
        'enable by default
        'or not
    Else
        mnuOptionsAdvDisplayVistaControls.Caption = mnuOptionsAdvDisplayVistaControls.Caption & Desk_Comp_Cap
        mnuOptionsAdvDisplayGlassBG.Caption = mnuOptionsAdvDisplayGlassBG.Caption & Desk_Comp_Cap
    End If
    
    If bVisualStyles Then
        mnuOptionsAdvDisplayVistaControls.Enabled = bComp_En
        
        'enable by default
        mnuOptionsAdvDisplayVistaControls.Checked = False
        mnuOptionsAdvDisplayVistaControls_Click
    Else
        mnuOptionsAdvDisplayVistaControls.Enabled = False
        mnuOptionsAdvDisplayVistaControls.Caption = mnuOptionsAdvDisplayVistaControls.Caption & " (Visual Styles must be on)"
    End If
    
    
    'Disable XP-only menus ############################################################################
    'mnuFileSettingsRMenu.Enabled = False
    mnuOptionsAdvNoStandby.Enabled = False: mnuOptionsAdvNoStandby.Checked = False
    mnuOptionsAdvNoStandbyConnected.Enabled = False: mnuOptionsAdvNoStandbyConnected.Checked = False
    
    'mnuFileSettingsRMenu.Caption = mnuFileSettingsRMenu.Caption & XP_Only_Cap
    mnuOptionsAdvNoStandby.Caption = mnuOptionsAdvNoStandby.Caption & XP_Only_Cap
    mnuOptionsAdvNoStandbyConnected.Caption = mnuOptionsAdvNoStandbyConnected.Caption & XP_Only_Cap
    '##################################################################################################
    
Else
    mnuOptionsAdvDisplayGlassBG.Enabled = False
    mnuOptionsAdvDisplayVistaControls.Enabled = False
    
    
    mnuOptionsAdvDisplayGlassBG.Caption = mnuOptionsAdvDisplayGlassBG.Caption & Vista_Only_Cap
    mnuOptionsAdvDisplayVistaControls.Caption = mnuOptionsAdvDisplayVistaControls.Caption & Vista_Only_Cap
    
    'If Not modLoadProgram.bIsIDE Then
        mnuFileThumb.Enabled = False
        mnuFileThumb.Caption = mnuFileThumb.Caption & Vista_Only_Cap
    'End If
End If
SetTBBanners
If modDisplay.VisualStyle Then
    txtStatus.Text = vbNullString
Else
    txtStatus.Text = sqB_Status_sqB
End If

Me.mnuDev.Visible = False
Me.mnuDevDataCmdsBlock.Checked = False
Me.mnuDevDataCmdsSetBlockMessage.Enabled = False
Me.mnuDevPriNormal.Checked = True

'Me.mnuConsole.Visible = False
Me.mnuRtfPopup.Visible = False
'frmSystray.mnuSB.Visible = False
'frmSystray.mnuSBObtain.Visible = False
'frmSystray.mnuStatus.Visible = False
mnuFileSettingsUserProfileExportOnExit.Checked = True
'Me.mnuOnlineManual.Visible = False
'Me.sbMain.Panels(3).Visible = True
'frmSystray.mnuFont.Visible = False
mnuOptionsAdvAutoUpdate.Checked = True
Me.mnuOptionsMessagingLoggingConv.Checked = True
'mnuOptionsDPEnable.Checked = True
mnuOptionsMessagingDisplaySmiliesComm.Checked = True
mnuOptionsMessagingDisplaySmiliesMSN.Checked = True
ApplySmileySettings

mnuOptionsMessagingHurgh.Checked = True
modSpeech.bHurgh = True

'Me.mnuOptionsWindow2Implode.Checked = True
'frmSystray.mnuDP.Visible = False
mnuDevDataCmdsTypeShow.Checked = True
'frmsystray.mnuCommands.Visible = False
Me.mnuOptionsMessagingLoggingDrawing.Checked = True
mnuOptionsAdvInactive.Checked = True
mnuOptionsAdvDisplayConnF_Click 1
'SetMenuColour
cmdReply(0).Visible = False
cmdReply(1).Visible = False
txtStatus.Visible = False
mnuOptionsMessagingDisplayShowHost.Checked = False
mnuOptionsAdvNetworkRefresh.Checked = True

mnuOnlineFTPDLO_Click eFTP_Methods.FTP_Default
mnuOnlineFTPULO_Click eFTP_Methods.FTP_Default

mnuOptionsAlertsStyle_Click 1
SetInactiveInterval 5
frmMain.mnuOptionsAdvInactive.Checked = True


ucVoiceTransfer.SaveDir = modSettings.GetUserSettingsPath()
ucFileTransfer.OverwriteFiles = True 'overwrite old DPs


imgClientDP.tag = "00"
picClient.Top = 1080
picBig.BackColor = vbWhite
ResetpicBigXY
picBig.ZOrder vbBringToFront
wbDP.Move -200, -300

picClear.BorderStyle = 0

Me.rtfIn.Text = vbNullString
Me.rtfIn.Locked = True
Me.rtfIn.EnableSmiles = True
Me.rtfIn.EnableTextFilter = False
'Me.rtfIn.OLEDropMode = vbManual
'Me.txtOut.OLEDropMode = vbManual

modLoadProgram.SetSplashInfo "Setting Variables..."

modTrig.TrigInit
modSpaceGame.InitVars
modStickGame.InitVars

LastName = modVars.User_Name

modSettings.ProcessTodoList

'#################################################################
'load FTP details
modFTP.iCurrent_FTP_Details = 0
For i = 0 To UBound(modFTP.FTP_Details)
    If i Then
        Load mnuOnlineFTPServerAr(i)
    End If
    
    With mnuOnlineFTPServerAr(i)
        .Visible = True
        .Checked = (i = 0)
        
        'If bDevMode Then
            '.Caption = modFTP.FTP_Details(i).FTP_User_Name & "@" & _
                modFTP.FTP_Details(i).FTP_Host_Name & _
                IIf(i = 0, " (Primary)", " (Backup)")
        'Else
        .Caption = "Use " & IIf(i = 0, "Primary", "Backup") & " Server" & IIf(i > 0, vbSpace & CStr(i), vbNullString)
        'End If
        
    End With
Next i
'#################################################################

SetSplashProgress 35

modNetwork.InitBandwidthStuff True

rtfFontName = rtfIn.FontName
prtfFontSize = rtfIn.Font.Size
prtfBold = False
prtfItalic = False


Set HTML_Doc = wbDP.Document

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

ReDim modMessaging.UsedIPs(0)
modMessaging.CurIPIndex = -1
'modMessaging.AddUsedIP "127.0.0.1"
'modMessaging.AddUsedIP SckLC.LocalIP


modConsole.frmMainhWnd = Me.hWnd
modPorts.Init

NewLine = True

If modLoadProgram.bSlow And modLoadProgram.bSafeMode = False Then
    modLoadProgram.SetSplashInfo "Updating Network List..."
    RefreshNetwork vbNullString
End If
SetSplashProgress 45


modLoadProgram.SetSplashInfo "Setting Variables..."
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


modLoadProgram.SetSplashInfo "Adding Main Form Script Object..."
SC.AddObject "frmMain", frmMain, True
'SC.AddObject "frmSystray", frmSystray, True

If bIsIDE = False Then
    modLoadProgram.SetSplashInfo "Adding URL Detection to rtfMain"
    rtfIn.EnableURLDetection rtfIn.hWnd
End If

bRecording = False

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


SetSplashProgress 55
'###################

'check if in right click menu
'Me.mnuFileSettingsRMenu.Checked = modVars.InRightClickMenu(RightClickExt, RightClickMenuTitle)

'file transfer stuff
'DP_Path = AppPath() & "Communicator Files"
Me.DP_Path = modSettings.GetUserSettingsPath() & "DP Files"

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

modSpaceGame.sGameModeMessage = modSpaceGame.ksGameModeMessage


SetPanelText "Version: " & GetVersion(), 3
Form_MouseMove 0, 0, 0, 0 'set panel 3 text



optnDraw(2).ToolTipText = "When this is selected, you can select a colour on the board to use"
optnDraw(1).ToolTipText = "Click two points on the board to draw a straight line between them"

ReDim irc_Commands(0 To 6)
irc_Commands(0).sCommand = "agree"
irc_Commands(0).bChatMessage = True
irc_Commands(1).sCommand = "describe"
irc_Commands(1).bChatMessage = True
irc_Commands(2).sCommand = "me"
irc_Commands(2).bChatMessage = True
irc_Commands(3).sCommand = "connect"
irc_Commands(3).bChatMessage = False
irc_Commands(4).sCommand = "disconnect"
irc_Commands(4).bChatMessage = False
irc_Commands(5).sCommand = "reconnect"
irc_Commands(5).bChatMessage = False
irc_Commands(6).sCommand = "summon"
irc_Commands(6).bChatMessage = True


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
modLoadProgram.SetSplashInfo "Initialising Variables..."
Call InitVars

ClosedWell = True

modLoadProgram.SetSplashInfo "Loading Systray..."
DoSystray True


'If modLoadProgram.bSlow Then
    modLoadProgram.SetSplashInfo "Obtaining IP Addresses..."
    
    modWinsock.ObtainRemoteIP
'Else
    'frmSystray.mnuSBObtain.Visible = True
    'lIP = frmMain.SckLC.LocalIP
'End If


Call ShowSB_IP 'add ip to status bar

SetSplashProgress 60

'modLoadProgram.SetSplashInfo "Checking FTP Connection..."
'Call AddToFTPList 'tell FTP server we are on


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

If InStr(1, Command(), "/reset", vbTextCompare) Then
    ClosedWell = False
End If

'############################################################################
modLoadProgram.SetSplashInfo "Loading Settings..."

Tmp = AppPath() & "Settings.cfg"

If FileExists(Tmp) Then
    modSettings.LoadSettings 'load some missed by below
    modSettings.ImportSettings Tmp, True, True
Else
    Tmp = modSettings.GetSettingsFile()
    
    If FileExists(Tmp) Then
        modSettings.LoadSettings 'load some missed by below
        modSettings.ImportSettings Tmp, True, True
        
    ElseIf modSettings.LoadSettings = False Or ClosedWell = False Then
        Call SetDefaultColours
        mnuOptionsAdvPresetReset_Click
        
    End If
    
'Else
    'AddText "UserProfile Settings Not Found (Registry Settings Used)", TxtError, True
End If
modSettings.LoadUsedIPs
'############################################################################



modLoadProgram.SetSplashInfo "Initialising Sound..."
modAudio.InitSounds 'bug report only

SetSplashProgress 70

'If ClosedWell = False Then
    'AddText "Last Communicator Crash Detected", , True
'End If

If modSpeech.sHiBye Then
    If modSpeech.sHi Then
        If modVars.bStartup = False Then
            
            If Me.mnuOptionsMessagingDisplaySysUserName.Checked Then
                Tmp = modVars.User_Name()
            Else
                Tmp = Me.LastName
            End If
            
            modSpeech.Say "Hello " & Tmp '& ", from" & IIf(modVars.bStealth, " stealthy", vbNullString) & " Communicator"
        End If
    End If
End If

'On Error Resume Next
'Kill SF
'On Error GoTo 0
modLoadProgram.SetSplashInfo "Checking Internet Status..."
If Not OnTheNet() Then
    'If App.PrevInstance Then
        'ExitProgram
        'Exit Sub
    'Else
    AddText "Internet Not Connected", , True
    'AddText "You May Close Me", , True
    Startup = False
    'End If
End If

modLoadProgram.SetSplashInfo "Processing Command Line..."
Call ProcessCmdLine(Startup, NoSubClass)

Call SetCmdButtons

Call CleanUp(False)

If App.PrevInstance And Not modLoadProgram.bJustUpdated Then
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

CheckAprilFools 'must be before show, so the reset_cmd is hidden

If Startup Then
    If Listen() Then
        ShowForm False, False
    Else
        ShowForm
    End If
ElseIf modVars.bStealth = False Then
    modLoadProgram.SetSplashInfo "Imploding Form..."
    'ImplodeFormToMouse Me.hWnd, True, True
    modImplode.AnimateAWindow Me.hWnd, aRandom
End If

If Startup = False And Not modVars.bStealth Then
    modLoadProgram.SetSplashInfo "Showing Form..."
    
    On Error GoTo LoadEH
    '##################################################################### FORM SHOWN
    ShowForm , False
    
    If modLoadProgram.frmSplash_Loaded Then
        
        frmSplash.Show vbModeless, frmMain
        
        'frmSplash.Move frmMain.Left + frmMain.width / 2 - frmSplash.width / 2, _
                       frmMain.Top + frmMain.height / 2 - frmSplash.height / 2
    
    End If
    
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
modLogging.addToActivityLog "Loaded Main Window"
modLoadProgram.SetSplashInfo Tmp
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


SetSplashProgress 80

LoadEH:
End Sub

Private Sub CheckAprilFools()

If modLoadProgram.bAprilFools Then
    AddText "April Fool!", TxtError, True
    
    modDisplay.Mirror Me
    modDisplay.MirrorhWnd rtfIn.hWnd, False
End If

cmdAprilFoolReset.Visible = modLoadProgram.bAprilFools

End Sub

Private Function GetSettingsPath() As String
GetSettingsPath = modSettings.GetSettingsFile() 'AppPath() & "Settings." & modVars.FileExt
End Function

Private Sub addDevActivatedText()
modDev.AddDevText "DevMode Activated - Level: " & modDev.getDevLevelName(), True
End Sub

Private Sub ProcessCmdLine(ByRef Startup As Boolean, ByRef NoSubClass As Boolean) ', _
            'ByRef ResetFlag As Boolean)

Dim CommandLine() As String
Dim i As Integer
Dim cmd As String, Param As String

Dim ClsFlag As Boolean
'Dim bVistaFlag As Boolean, sVistaParam As String

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
            
            'check param against all passwords...
            If modDev.devLogin(Param) Then
                addDevActivatedText
            Else
                AddText "DevMode Password Incorrect", TxtError, True
            End If
            
        'Case "vista"
            'bVistaFlag = True
            'sVistaParam = Param
            
        Case "xpbuttons"
            If LenB(Param) = 0 Then Param = "1"
            
            modLoadProgram.bAllowXPButtons = CBool(Param)
            AddText "XP Buttons Forced " & IIf(modLoadProgram.bAllowXPButtons, "On", "Off"), , True
            
            
        Case "startup"
            Startup = True
            
            
        Case "host"
            If Param <> vbNullString Then
                mnuOptionsHost.Checked = CBool(Param)
            Else
                mnuOptionsHost.Checked = True
            End If
            
            AddText "Host Mode " & IIf(mnuOptionsHost.Checked, "On", "Off"), , True
        
            
        Case "subclass"
            
            If Param <> vbNullString Then
                NoSubClass = Not CBool(Param)
            Else
                NoSubClass = False
            End If
            
            AddText "Subclassing " & IIf(NoSubClass, "Off", "On"), , True
            
            
        Case "cls"
            ClsFlag = True
            
        'Case "quick"
            'AddText "Started in Quick Mode, this is not recommended", TxtError, True
            
        Case "internet"
            
            If LenB(Param) = 0 Then Param = "1"
            modVars.bNoInternet = Not CBool(Param)
            
            
        Case "aprilfools"
            
            If LenB(Param) = 0 Then Param = "1"
            
            modLoadProgram.bAprilFools = CBool(Param)
            
            
        Case "log"
            If LenB(Param) = 0 Then Param = "1"
            Me.mnuOptionsMessagingLoggingConv.Checked = CBool(Param)
            AddText "Logging " & IIf(Me.mnuOptionsMessagingLoggingConv.Checked, "Enabled", "Disabled"), , True
            
            
        Case "logall"
            If LenB(Param) = 0 Then Param = "1"
            Me.mnuOptionsMessagingLoggingConv.Checked = CBool(Param)
            Me.mnuOptionsMessagingLoggingAutoSave.Checked = CBool(Param)
            Me.mnuOptionsMessagingLoggingDrawing.Checked = CBool(Param)
            Me.mnuOptionsMessagingLoggingPrivate.Checked = CBool(Param)
            'Me.mnuOptionsMessagingLoggingActivity.Checked = CBool(Param)
            
            
        Case "autosave"
            
            If LenB(Param) = 0 Then Param = "1"
            
            Me.mnuOptionsMessagingLoggingAutoSave.Checked = CBool(Param)
            
            AddText "AutoSave " & IIf(mnuOptionsMessagingLoggingAutoSave.Checked, "Enabled", "Disabled"), , True
            
            
        Case "killold"
            
            KillOldVersion
            
            
        Case "gamemode"
            
            mnuFileGameMode_Click
            
        Case "console", "debug", "stealth", "safemode", "skipdll", _
                "forceopen", "instanceprompt", "killold"
            
            'donothing
            
        Case Else
            
            AddText "-----" & vbNewLine & _
                "Commandline Command not recognised:" & vbNewLine & _
                "'" & CommandLine(i) & "'" & vbNewLine & _
                "-----", TxtError
            
    End Select
    
Next i


'If bVistaFlag Then
'    If LenB(sVistaParam) Then
'        On Error Resume Next
'        modLoadProgram.bVistaOrW7 = CBool(sVistaParam)
'    Else
'        modLoadProgram.bVistaOrW7 = True
'    End If
'
'    AddText "Vista mode forced " & IIf(modLoadProgram.bVistaOrW7, "on", "off"), , True
'
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
        If Frm.tag = "visible" Then
            Call FormLoad(Frm, , False, False)
            Frm.Visible = True
            Frm.tag = vbNullString
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
                            Frm.tag = "visible"
                            
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
        
        frmSystray.ShowBalloonTip "You are still connected" & vbNewLine & _
                                  "Right Click and Select 'Close Connection' to disconnect", , NIIF_WARNING
        
        
    ElseIf Status = Idle Then
        
        If Me.mnuOptionsHost.Checked Then
            If Listen(False) Then
                frmSystray.ShowBalloonTip ListeningStr & vbNewLine & _
                                          "Right click the tray icon and select Close Connection to stop" & vbNewLine & _
                                          vbNewLine & "Click this to get rid of it", , NIIF_INFO, 500
            End If
        End If
        
    End If
    
End If

If modLoadProgram.frmMini_Loaded Then
    'frmMini.chkComm.Value = Abs(Me.Visible)
    frmMini.setchkComm_Value Abs(Me.Visible)
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

With cmdXPReply(0)
    .Top = rtfIn.Top + 100
    .Left = rtfIn.Left + rtfIn.width - .width - 350
End With
With cmdXPReply(1)
    .Top = cmdXPReply(0).Top
    .Left = cmdXPReply(0).Left - .width + 375
End With

cmdReply(0).Left = cmdXPReply(0).Left + cmdReply(0).width / 2
cmdReply(1).Left = cmdReply(0).Left - cmdReply(1).width
cmdReply(0).Top = cmdXPReply(0).Top
cmdReply(1).Top = cmdXPReply(1).Top

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

SetInfoPanel CStr(ScaleX(Me.width, vbTwips, vbPixels)) & " x " & CStr(ScaleY(Me.height, vbTwips, vbPixels))


If modLoadProgram.frmThumbNail_Loaded Then
    frmThumbnail.RefreshThumbNailRect False
End If

End Sub

Private Sub mnuFileExit_Click()

If modFTP.bFTP_Doing Then
    AddText "Can't Exit - FTP Transfer in Progress", TxtError, True
Else
    If Question("Exit, Are You Sure?", mnuFileExit) = vbYes Then
        ExitProgram
    Else
        AddText "Exit Canceled", , True
    End If
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
Static Current As String
Const devWord As String = "DevMode"
Dim sPass As String

If mnuOptionsMatrix.Checked = False Then
    If Chr$(KeyAscii) = Mid$(devWord, Len(Current) + 1, 1) Then
        Current = Current & Chr$(KeyAscii)
        
        If Current = devWord Then
            sPass = modVars.Password("Enter a dev password", Me, _
                "Dev Password", , True, 10)
            
            
            If LCase$(sPass) = "off" Then
                mnuDevChangeAr_Click 0
            ElseIf modDev.devLogin(sPass) Then
                addDevActivatedText
            Else
                AddText "Incorrect Password", TxtError, True
                
                Select Case LCase$(sPass)
                    Case "rob"
                        modSpeech.Say "Rob is well awesome<SILENCE MSEC=""5000"" />The Game."
                    Case "steven"
                        modSpeech.Say "Steven sucks<SILENCE MSEC=""5000"" />The Game."
                    Case "steve"
                        modSpeech.Say "Oh, is that err... Steve? <SILENCE MSEC=""5000"" />The Game."
                        
                    Case "dogsbody"
                        modSpeech.Say "Oh is that errr... bodsdogy? I err... quite like it....<SILENCE MSEC=""3500"" />" & _
                            "<RATE SPEED=""5"">Dogsbody invented and created by Groin.</RATE>"
                    Case "cj"
                        modSpeech.Say "oh, CJ! oh, oh, oh, oh, oh, oh, <PITCH MIDDLE=""9"">oh, oh, oh, oh," & _
                            "<SILENCE MSEC=""1000"" /> ah yeah, all riight.</PITCH><SILENCE MSEC=""3500"" />" & _
                            "Yeah. That's the ticket."
                            
                    Case "dogs"
                        modSpeech.Say "and chickens, and dogs and chicken's dogs.<SILENCE MSEC=""5000"" />" & _
                            "But is the chicken plural?<SILENCE MSEC=""5000"" />The Game."
                    Case "meat"
                        modSpeech.Say "I do like a bit of tasty meat, eeeeeeeeee."
                        
                    Case "tim"
                        modSpeech.Say "Back in time Timmeh was 'ear, oh-nine."
                    Case "timmeh"
                        modSpeech.Say "Timmeh is not available right nao.<SILENCE MSEC=""5000"" />The Game."
                        
                    Case "bojo", "boje", "bojé"
                        modSpeech.Say "Oh err... are you err.... on err.... MSN? I err... quite like it err... What do you think of MSN?" & _
                            "<SILENCE MSEC=""1000"" /><RATE SPEED=""-5""><PITCH MIDDLE=""-10"">MACE</PITCH></RATE>"
                            
                        
                End Select
            End If
            
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
                
                MidText StrOut, TxtForeGround
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
        
        'mnuRtfPopupDelSel.Enabled = bFlag
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
SetFocus2txtOut
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
    
    Msg = Msg & " Disconnected (" & FormatDateTime$(Time$, vbLongTime) & ")"
    
    If bTell Then
        AddText Msg, , True, True
    End If
    
    'used to be .count -1, but since it is unloaded above, it is .count -1 + 1 = .count... or not
    For i = 1 To SockAr.UBound '(SockAr.Count - 1) '- unreliable, could have a low one d/c and etc but errorless
        
        If i <> Index Then
            If ControlExists(SockAr(i)) Then
                If SockAr(i).state = sckConnected Then
                    Ctd = True
                    Exit For
                End If
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
            
            If Me.Visible Then ShowForm False, False
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
            
            modMessaging.EvalTyping "0" & Clients(i).sName 'remove from typing list
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
                Clients(i).bDPIsGIF = False
                Clients(i).iSocket = 0
                Clients(i).lLastPing = 0
                Clients(i).lPingStart = 0
                Clients(i).sStatus = vbNullString
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
Else
    Pause 50
    SockAr(Index).Close
End If

End Sub

Private Sub SckLC_Close()
'handles the closing of the connection

'Static Cleaned As Boolean
Dim Str As String, IP As String ', Msg As String
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

Str = Str & " (" & FormatDateTime$(Time$, vbLongTime) & ")"

AddConsoleText Str
If modVars.bRetryConnection = False Then
    
    If (Not modVars.IsForegroundWindow()) And (Not Closing) Then
        frmSystray.ShowBalloonTip Str & " - " & FormatDateTime$(Time$, vbLongTime), , NIIF_INFO
    End If
    
    AddText Str, , True, True
    
    'If Server Then
        'Msg = "All Connections Closed"
    'Else
        'Msg = "Disconnected from Server"
    'End If
End If

'If SendTypeTrue Then
    'txtOut.Text = vbNullString
    'DoEvents
'End If

SckLC.Close  'close connection


If modVars.bRetryConnection = False Then
    Call CleanUp(True)
'Else
    'timer'll take care of it
End If

'Cmds Idle

End Sub

Private Sub SckLC_Connect()
Dim TimeTaken As Long
Dim Text As String
Dim Server_NameIP As String, HostIP As String
Dim bAutoRetry As Boolean

'txtLog is the textbox used as our
'chat buffer.

'SckLC.RemoteHost returns the hostname( or ip ) of the host
'SckLC.RemoteHostIP returns the IP of the host
Cmds Connected

If modVars.bRetryConnection Then
    bAutoRetry = True
End If
modVars.bRetryConnection = False


TimeTaken = GetTickCount() - ConnectStartTime
HostIP = SckLC.RemoteHostIP

If mnuOptionsMessagingDisplayShowHost.Checked Then
    Server_NameIP = SckLC.RemoteHost
    
    If LenB(Server_NameIP) = 0 Or Server_NameIP = HostIP Then
        Server_NameIP = HostIP
    Else
        Server_NameIP = Server_NameIP & " [" & HostIP & "] "
    End If
Else
    Server_NameIP = HostIP
End If

Text = "Connected to " & Server_NameIP & " in " & _
    CStr(TimeTaken / 1000) & _
    " seconds (" & FormatDateTime$(Time$, vbLongTime) & ")"

AddText Text, , True, True
AddConsoleText Text
modLogging.addToActivityLog Text

SetMiniInfo "Connected to " & Server_NameIP

Pause 25 'let the other info get through...?
Text = FormatApostrophe(LastName) & " Version: " & GetVersion() & _
    IIf(bAutoRetry, ". (Auto-Connected)", vbNullString)

'SendData eCommands.Info & Text & "0"
SendInfoMessage Text


If LenB(Inviter) > 0 Then
    Pause 25 'let the other info get through...?
    Text = LastName & " was invited by " & Trim$(Inviter)
    'SendData eCommands.Info & Text
    SendInfoMessage Text
End If


If frmMain.Visible = False Then
    'frmSystray.ShowBalloonTip "Communicator has connected to " & SckLC.RemoteHostIP & ":" & CStr(RPort), , NIIF_INFO, , True
    frmSystray.ShowBalloonTip "Communicator has connected to " & _
        Server_NameIP & " on port " & CStr(MainPort), _
        "Communicator Connection Established", NIIF_INFO, , True
    
    ShowForm
End If


SetInfo "Connected, Awaiting Welcome Message...", False

SetFocus2txtOut
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
    sTxt = "Their computer is shut down/in standby or the computer doesn't exist"
    
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


SckLC_Close

AddConsoleText "SckLC Error: " & Description
modLogging.addToActivityLog "Main socket error: " & Description

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
Dim Txt As String, IP As String, IP_And_Name As String
Dim SystrayTxt As String, UpTimeText As String, ServerName As String

'client initilisation
Dim bInit As Boolean, bBlockThisIP As Boolean


SocketToUse = -1

For i = 1 To SockAr.Count - 1
    If SockAr(i).state = sckClosed Then
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
SockAr(SocketToUse).accept requestID

IP = SockAr(SocketToUse).RemoteHostIP 'SckLC.RemoteHostIP

If mnuOptionsMessagingDisplayShowHost.Checked And LenB(SockAr(SocketToUse).RemoteHost) Then
    IP_And_Name = SockAr(SocketToUse).RemoteHost & " [" & IP & "] "
Else
    IP_And_Name = IP
End If

If modMessaging.bAllBlocked Then
    'see if we can allow this guy
    bBlockThisIP = True 'assume blocked
    
    For i = 0 To UBound(modMessaging.BlockedIPs)
        If LenB(modMessaging.BlockedIPs(i)) Then
            If IP = modMessaging.BlockedIPs(i) Then
                'allowed, exit prematurely
                bBlockThisIP = False
                Exit For
            End If
        End If
    Next i
Else
    'see if this guy's blocked
    
    For i = 0 To UBound(modMessaging.BlockedIPs)
        If LenB(modMessaging.BlockedIPs(i)) Then
            If IP = modMessaging.BlockedIPs(i) Then
                bBlockThisIP = True
            End If
        End If
    Next i
End If


If bBlockThisIP Then
    AddConsoleText "Blocked IP (" & IP_And_Name & ") was kicked"
    
'    If mnuOptionsMessagingDisplayShowBlocked.Checked Then
'        Txt = "Blocked IP (" & IP_And_Name & ") attempted to connect - Rejected"
'
'        SendInfoMessage Txt, True, , , SocketToUse
'        If modSpeech.sSayInfo Then modSpeech.Say Txt
'        AddText Txt, TxtError, True
'        Txt = vbNullString
'    End If
    
    'SendData eCommands.Info & "You have been Kicked - Your IP is blocked1", SocketToUse
    SendInfoMessage "You have been Kicked - Your IP is blocked", , True, , SocketToUse
    
    Kick SocketToUse, "Blocked IP (" & IP_And_Name & ")", mnuOptionsMessagingDisplayShowBlocked.Checked, True
    
    Exit Sub
End If


If Status <> Connected Then Cmds Connected


SystrayTxt = IP_And_Name & " connected." & vbNewLine & "(Client " & CStr(SocketToUse) & ")"
Txt = IP_And_Name & " (Client " & CStr(SocketToUse) & ") Connected" & " (" & FormatDateTime$(Time$, vbLongTime) & ")"

'add to the log
AddText Txt, , True, True
AddConsoleText Txt   '& " SocketHandle: " & SockAr(SocketToUse).SocketHandle
modLogging.addToActivityLog "New connection " & Txt

SetMiniInfo IP_And_Name & " connected."

'if server then modmessaging.DistributeMsg "Client
'SendData eCommands.GetName, SocketCounter


'If Server Then modMessaging.DistributeMsg eCommands.Info & Txt & "0", SocketToUse
SendInfoMessage Txt, , , , , SocketToUse
'no point telling guy who's connected that he's connected

'frmSystray.ShowBalloonTip sName, "Communicator - New Connection", NIIF_INFO
'don't show here, wait until we get a name

UpTimeText = FormatTimeElapsed((GetTickCount() - modLoadProgram.LoadStart) / 1000)


'#########################################################################################################
'info, etc
modMessaging.SendSetSocketMessage SocketToUse

'modMessaging.SendData _
    eCommands.Info & "Welcome to " & LastName & "'" & IIf(Right$(LastName, 1) = "s", vbNullString, "s") & _
    " Server, Version: " & GetVersion() & _
    ". Server Up Time: " & MinsUp & " min" & IIf(MinsUp > 1, "s", vbNullString) _
    & "0", SocketToUse '                                           convert into minutes ^

ServerName = FormatApostrophe(LastName) & " server"

SendInfoMessage "Welcome to " & ServerName & ", Version: " & GetVersion() & _
    ". Server Up Time: " & UpTimeText, , , , SocketToUse

'If LenB(ServerMsg) Then modMessaging.SendData eCommands.Info & "Server Message: " & ServerMsg & "0", SocketToUse
If LenB(ServerMsg) Then
    SendInfoMessage "Server Message: " & ServerMsg, , , , SocketToUse
End If


SendData eCommands.cmdOther & eOtherCmds.SetServerName & ServerName, SocketToUse


If mnuFileGameMode.Checked Then
    'modMessaging.SendData eCommands.Info & sGameModeMessage & "1", SocketToUse
    SendInfoMessage sGameModeMessage, True, , , SocketToUse
    
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
sTxt = FormatDateTime$(Time$, vbLongTime) & vbNewLine & "Error with Client on Socket " & CStr(Index)

i = FindClient(Index)
If i > -1 Then
    If LenB(Clients(i).sName) Then
        sTxt = sTxt & vbNewLine & "Name: " & Clients(i).sName
        If LenB(Clients(i).sVersion) Then
            sTxt = sTxt & ",  Version: " & Clients(i).sVersion
        End If
    End If
End If


sTxt = sTxt & vbNewLine & Description & " (" & Number & ")"


'append the error message in the chat buffer
AddText Trim$(InfoStart) & vbNewLine & sTxt & vbNewLine & Trim$(InfoEnd), TxtError, , True

'AddConsoleText "SockAr_Error:" & sTxt


'If Server Then modMessaging.DistributeMsg eCommands.Info & sTxt & "1", Index
'should be true, but...
'screw it
sTxt = "Error with Client on Socket " & CStr(Index)
If i > -1 Then
    If LenB(Clients(i).sName) Then
        sTxt = sTxt & " (" & Clients(i).sName & ")"
    End If
End If
SendInfoMessage sTxt, False, True, , , Index


sockClose Index, False

End Sub

Private Sub tmrHost_Timer()

Dim JustListened As Boolean
Dim Frm As Form

If modVars.bModalFormShown Then Exit Sub


InActiveTmr = InActiveTmr + 1

'frmSystray.RefreshTray

If mnuOptionsHost.Checked Then
    If InActiveTmr >= 1 Then  '30 seconds
        If (Status <> Connected) And (Status <> Connecting) And modVars.bRetryConnection = False Then
            If SckLC.state <> sckListening Then
                'Call CleanUp 'handled below
                'frmMain.ClearRtfIn
                Call Listen
                JustListened = True
            End If
        End If
    End If
End If

If (Status <> Connected) And (Not JustListened) Then
    If mnuOptionsAdvNetworkRefresh.Checked Then
        RefreshNetwork vbNullString
    End If
End If

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

If LenB(modMessaging.LastIP) = 0 Then
    modVars.bRetryConnection = False
    AddText "Error - IP was lost. You should never see this text", TxtError, True
    If Me.Visible = False Then ShowForm
    Exit Sub
End If

Connect modMessaging.LastIP

sTxt = " to " & modMessaging.LastIP & ":" & MainPort & "..."

If bIsRetry Then
    SetInfo "Retrying connection" & sTxt, True
Else
    SetInfo "Connecting" & sTxt, False
End If
    
    'LastAutoRetry = GTC
'End If

End Sub

Private Sub tmrMain_Timer()
Dim SendList As String, i As Integer
Dim Addr As String, Str As String
Dim sSelected As String, sTxtToAdd As String
'Dim bHad As Boolean
Dim LBTopI As Integer
Dim tPic As IPictureDisp
Static LastIconRefresh As Long

MonitorGameWindow

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
LBTopI = SendMessageByLong(lstConnected.hWnd, LB_GETTOPINDEX, 0, 0)

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
        
        SendData eCommands.SetClientVar & eClientVarCmds.SetDPSet & CStr(Abs(modDP.DP_Path_Exists()))
        
        If LenB(Clients(i).sIP) = 0 Then
            If Clients(i).iSocket = -1 Then
                Clients(i).sIP = SckLC.RemoteHostIP
            End If
        End If
        
    Else
        
        '                                 V, check if i=0, in case we have sent our internal ip and now have an external ip to send
        If LenB(Clients(i).sIP) = 0 Or i = 0 Then
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
            If LenB(Clients(i).sIP) = 0 Then
                Clients(i).sIP = SckLC.LocalIP
            End If
            
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
        SendMessageByLong lstConnected.hWnd, LB_SETTOPINDEX, LBTopI, 0
    End If
End With




If Server Then
    
    '######################
    'send clients list
    
    If Clients(0).sName <> LastName Then
        Clients(0).sName = LastName
        Clients(0).iSocket = -1
    End If
    
    
    Str = modWinsock.RemoteIP
    
    If LenB(Str) Then
        Clients(0).sIP = Str
    Else
        Clients(0).sIP = modWinsock.LocalIP
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
    
    If ucVoiceTransfer.iCurSockStatus <> sckListening Then
        If ucVoiceTransfer.iCurSockStatus <> sckConnected Then
            ucVoiceTransfer.Listen VoicePort
        End If
    End If
Else
    '######################
    'send my name
    modMessaging.SendData eCommands.SetClientVar & eClientVarCmds.SetName & LastName
    
    'send my info
    modMessaging.SendData eCommands.SetClientVar & eClientVarCmds.SetVersion & GetVersion()
    modMessaging.SendData eCommands.SetClientVar & eClientVarCmds.SetsStatus & LastStatus
    
End If

'#################################
modDP.tmrMain_Timer
tmrPing_Timer


'clients have been updated, update picClient
On Error GoTo EHContinue

If LenB(picClient.tag) Then
    i = CInt(picClient.tag)
    
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

Private Sub MonitorGameWindow()
If GameWindowhWnd > 0 Then
    If Me.mnuFileGameMode.Checked Then
        If IsWindow(GameWindowhWnd) = 0 Then
            GameWindowhWnd = 0
            'deactivate game mode
            Me.mnuFileGameMode_Click
        End If
    Else
        'shouldn't get here
        GameWindowhWnd = 0
    End If
End If
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
'LogPrivate

iAuto = iAuto + 1
If iAuto >= 2 Then '20s
    AutoSave
    
    iAuto = 0
End If

End Sub

Private Sub DoLog(Optional bForce As Boolean = False)
'10 sec interval

Static T As Integer
Dim logPath As String ', FilePath As String


If Me.mnuOptionsMessagingLoggingConv.Checked Then
    
    If bForce Then
        T = 3
    Else
        T = T + 1
    End If
    
    If T >= 3 Then '30 seconds
        T = 0
        
        
        logPath = GetCurrentLogFolder() 'AppPath() & "Logs\"
        'If FileExists(LogPath, vbDirectory) = False Then
            'MkDir LogPath
        'End If
        'LogPath = LogPath & MakeDateFile() & "\"
        If FileExists(logPath, vbDirectory) = False Then
            MkDir logPath
        End If
        
        
        If LenB(currentLogFile) = 0 Then
            currentLogFile = logPath & MakeLogFileName() & ".rtf"
        End If
        
        
        'If Status = Connected Then
        On Error GoTo EH
        rtfIn.SaveFile currentLogFile, rtfRTF
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

'Public Sub LogPrivate()
'Dim RootPath As String, sPath As String, sName As String
'Dim Frm As Form
'Dim i As Integer, f As Integer
'
'If modVars.nPrivateChats Then
'    If mnuOptionsMessagingLoggingPrivate.Checked Then
'
'
'        RootPath = GetCurrentLogFolder() 'GetLogPath() & MakeDateFile() & "\"
'        If FileExists(RootPath, vbDirectory) = False Then
'            MkDir RootPath
'        End If
'
'
'        For Each Frm In Forms
'            If Frm.Name = frmPrivateName Then
'
'                i = FindClient(Frm.SendToSock)
'
'                If i > -1 Then
'                    sName = Clients(i).sName
'                Else
'                    sName = "Randomer"
'                End If
'
'                sPath = RootPath & sName & " Private " & MakeTimeFile() & ".rtf"
'
'
'                On Error GoTo EH
'                Frm.rtfIn.SaveFile sPath, rtfRTF
'
'            End If
'        Next Frm
'
'    End If
'End If
'
'Exit Sub
'EH:
'If Err.Number <> err_INVALIDORNOACCESS Then
'    AddText "Log Error - " & Err.Description, TxtError, True
''else
'    'they are viewing the log
'End If
'End Sub

Private Function MakeLogFileName() As String
MakeLogFileName = MakeTimeFile()
End Function

Public Function MakeTimeFile() As String
MakeTimeFile = Replace$(FormatDateTime$(Time$, vbLongTime), ":", ".")
'MakeTimeFile = Replace$(Time$, ":", ".")
'Replace$(Replace$(CStr(Date & " - " & Time), "/", ".", , , vbTextCompare), ":", ".", , , vbTextCompare)
End Function

Private Function MakeDateFile() As String
MakeDateFile = GetDate() 'Replace$(CStr(Date), "/", ".")
End Function

Public Function GetCurrentLogFolder() 'path of today's log folder

GetCurrentLogFolder = GetLogPath() & MakeDateFile() & "\"

End Function

Public Function GetLogPath() As String 'root path of all logs
'Dim sPath As String
'sPath = AppPath() & "Communicator Logs\"

checkLogPath
If FileExists(logBasePath, vbDirectory) = False Then
    On Error Resume Next
    MkDir logBasePath
End If

GetLogPath = logBasePath

End Function

Private Sub checkLogPath()
Dim bReset As Boolean

If LenB(logBasePath) = 0 Then
    bReset = True
ElseIf FileExists(logBasePath, vbDirectory) = False Then
    bReset = True
End If

If bReset Then
    'logBasePath = Environ$("HOMEPATH")
    'logBasePath = logBasePath & IIf(Right$(logBasePath, 1) <> "\", "\", vbNullString) & "My Documents\Communicator Logs\"
    logBasePath = modPaths.logPath
End If

If Right$(logBasePath, 1) <> "\" Then logBasePath = logBasePath & "\"

End Sub

Public Sub resetLogPath()
logBasePath = vbNullString
checkLogPath
End Sub

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
    If modSpeech.bHurgh Then
        modAudio.PlayHurgh
    Else
        Beep
    End If
End If

End Sub

'########################################################################################################

Public Sub tmrLP_Timer()

If modLoadProgram.bSlow Then
    modLogin.AddToFTPList
    If mnuOnlineFTPServerMsg.Checked Then DownloadServerMessage True
End If


ProcessUpdate

End Sub

Private Sub DownloadServerMessage(bStealth As Boolean)
Dim sMsg As String, sError As String
Dim eError As eFTPCustErrs
Dim iFTP_Setting As Integer

If bStealth Then
    iFTP_Setting = modFTP.iCurrent_FTP_Details
    modFTP.iCurrent_FTP_Details = 0 'force default
    modFTP.FTP_StealthMode = True
End If

modFTP.GetFileStr sMsg, eError, _
    modFTP.FTP_Root_Location & "/Messages/Server Message" & Dot & FileExt, _
    mnuFileExit, sError, False

If eError = cSuccess Then
    AddText "FTP Server Message - " & sMsg, TxtReceived, True
ElseIf bStealth Then
    AddConsoleText "Error (" & CStr(eError) & ") Downloading Server Message - " & sError
Else
    AddText "Error (" & CStr(eError) & ") Downloading Server Message - " & sError, TxtError, True
End If

If bStealth Then
    modFTP.FTP_StealthMode = False
    modFTP.iCurrent_FTP_Details = iFTP_Setting
End If
End Sub

Private Sub ProcessUpdate()
Dim Ans As VbMsgBoxResult
Dim DaysSinceUpdate As Integer
Dim Tmp As String

Const DayDiff = "d"


On Error GoTo EH
DaysSinceUpdate = DateDiff(DayDiff, LastUpdate, Date)

AddConsoleText "Days Since Last Update: " & CStr(DaysSinceUpdate)


If DaysSinceUpdate > 5 And mnuOptionsAdvAutoUpdate.Checked Then
    
    If Me.Visible = False Then
        Me.ShowForm
    End If
    
    Ans = Question("You haven't checked for an update in the past five days, check now?", mnuFileExit)
    AddText "If this is annoying, turn it off - 'Options > Advanced > Remind me...'", , True
    
    If Ans = vbYes Then
        Call CheckForUpdates
    Else
        AddText "Check Canceled", , True
    End If
    
'ElseIf bJustUpdated Then
'
'    Tmp = GetSettingsPath()
'
'    If FileExists(Tmp) Then
'        Ans = MsgBoxEx("Settings File is old;" & vbNewLine & "Export Settings now, to overwrite the old file?", _
'            "This is because the settings file structure has been changed, so the old file needs to be overwritten", _
'            vbQuestion + vbYesNo, "Export Settings?", , , Me.Icon.Handle)
'
'        If Ans = vbYes Then
'            modSettings.ExportSettings Tmp
'        End If
'    End If
    
End If

EH:
End Sub

'########################################################################################################

'Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
'
'Call Form_KeyDown(KeyCode, Shift)
'
'End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

If KeyAscii = 172 Then
    KeyAscii = 0 'prevent ¬
ElseIf KeyAscii = 13 Then
    KeyAscii = 0
    Rename txtName.Text
End If

End Sub

Private Sub txtName_Change()
Dim Txt As String

On Error GoTo EH

If Screen.ActiveControl.Name = txtName.Name Then
    Txt = Trim$(txtName.Text)
    If LenB(Txt) Then
        If Txt <> LastName Then
            modDisplay.ShowBalloonTip txtName, "Name Setting", "Click off the textbox to set your name"
        End If
    End If
End If

EH:
End Sub

Private Sub txtName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    'If Status = Connected Then
    
    txtName.Enabled = False
    'Pause 1
    txtName.Enabled = True
    
    SetFocus2 txtName
    
    PopupMenu frmSystray.mnuStatus
    'End If
End If

End Sub

Private Sub txtName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer

HidePicBig

If Status = Connected Then
    With picClient
        
        i = FindClient(modMessaging.MySocket)
        
        
        If .tag <> CStr(i) Then
            If i > -1 Then
                If LenB(Clients(i).sName) Then
                    
                    .Move txtName.Left + txtName.width, txtName.Top
                    .Visible = True
                    
                    ShowClientInfo i
                    
                    If picBig.Visible Then picBig.Visible = False
                    
                    .tag = CStr(i)
                    
                End If
            End If
        End If
        
    End With
End If


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
    
    If Status = Connected Then
        txtName.Enabled = True '(modVars.nPrivateChats = 0)
        
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
        txtName.Enabled = True
    End If
    
Else
    If Status = Connected Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False '(Left$(txtOut.Text, 1) = "/")
    End If
    cmdSend.Default = cmdSend.Enabled
    
    If Status = Connected Then
        txtName.Enabled = False
        
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
    Else
        txtName.Enabled = True
    End If
    
End If

'Call SetInactive
'txtOut.height = TextHeight(txtOut.Text) - 90
Dim i As Integer, j As Integer

Msg = txtOut.Text
If Len(Msg) = 1 Then
    If Left$(Msg, 1) = "/" Then
        Dim S As String
        
        For i = 0 To UBound(irc_Commands)
            S = S & irc_Commands(i).sCommand & ", "
        Next i
        
        SetInfo "Commands: " & Left$(S, Len(S) - 2), False
    End If
    
ElseIf Len(Msg) >= 3 Then 'Len("/de")
    If Left$(Msg, 1) = "/" Then
        Msg = Mid$(Msg, 2)
        If matches_ircCommand(Msg, i, j) Then
            If j >= Len(irc_Commands(i).sCommand) Then
                'something like "/describe greetings"
                If Status = Connected Or Not irc_Commands(i).bChatMessage Then
                    cmdSend.Enabled = True
                    cmdSend.Default = True
                Else
                    cmdSend.Enabled = False
                    SetInfo "Can't execute chat command - not connected", True
                End If
            Else
                'something like "/de" or "/dehello"
                SetInfo "Press the Right Key to add '/" & irc_Commands(i).sCommand & "'", False
            End If
        End If
    End If
End If

End Sub

Public Function RefreshNetwork(ByRef sError As String, Optional ByRef LB As Control = Nothing, _
                          Optional ByRef CommentLB As Control = Nothing, _
                          Optional ByVal bForce As Boolean = False) As Boolean

Dim Svr As ListOfServer
Dim i As Integer
Dim S As String
Dim AddComment As Boolean
Dim bIsfrmMain As Boolean


If bHadNetworkRefreshError And Not bForce Then Exit Function
If modLoadProgram.bIsIDE Then Exit Function

If LB Is Nothing Then
    Set LB = lstComputers
    bIsfrmMain = True
End If
If Not (CommentLB Is Nothing) Then
    AddComment = True
    CommentLB.Clear
End If


LB.Clear
Me.MousePointer = vbHourglass

'S = "Refreshing server list..."
'
'If Right$(modConsole.ConsoleText, Len(S) + 2) <> (S & vbNewLine) Then
'    AddConsoleText S
'End If

Svr = EnumServer(SRV_TYPE_ALL, sError)

If Svr.Init Then
    For i = 1 To UBound(Svr.List)
        'If InviteBox = False Then
        LB.AddItem Svr.List(i).ServerName
        
        If AddComment Then
            
            With Svr.List(i)
                
                S = .Comment
                
                If LenB(S) Then
                    If .Type = 6 Then
                        S = S & vbSpace & "(Vista)"
                    ElseIf .Type = 5 Then
                        S = S & vbSpace & "(XP)"
                    End If
                End If
                
                CommentLB.AddItem S
                
                    'modVars.TranslateWindowsVer(.VerMajor, .VerMinor, ((.PlatformId And &H80000000) = 0), .Type = 6)
                
            End With
            
        End If
        
        'Else
            'frmInvite.lstComputers.AddItem Svr.List(i).ServerName
        'End If
    Next i
    'AddConsoleText "Done"
Else
    If Svr.LastErr > 0 Then
        bHadNetworkRefreshError = True
        AddConsoleText "Error refreshing server list:" & vbNewLine & "    " & sError
    End If
End If


RefreshNetwork = Svr.Init

Me.MousePointer = vbNormal

End Function

'------------DRAWING------------------

Public Sub DoLine(ByVal X As Integer, ByVal Y As Integer)

'If modVars.IsForegroundWindow() Then
    If NewLine Then
        picDraw.Line (X, Y)-(X, Y), colour
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
                optnDraw_Click 0
            End If
            
            Dim TmpColour As Long, TmpWidth As Integer
            TmpColour = colour
            TmpWidth = picDraw.DrawWidth
            colour = picDraw.BackColor
            picDraw.DrawWidth = RubberWidth
            
            aX = CInt(X)
            aY = CInt(Y)
            
            Call DoLine(aX, aY)
            
            'SendLine X, Y, picDraw.DrawWidth
            
            colour = TmpColour
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
            
            picDraw.Line (cx, cy)-(aX, aY), colour
            picDraw.Refresh
            'Remember where the mouse is so new lines can be drawn connecting to this point.
            
            tcX = cx
            tcY = cy
            
            cx = aX
            cy = aY
            
            SendLine tcX, tcY, picDraw.DrawWidth, aX, aY
            
        
        ElseIf Button = vbRightButton Then
            
            Dim TmpColour As Long, TmpWidth As Integer
            TmpColour = colour
            TmpWidth = picDraw.DrawWidth
            
            colour = picDraw.BackColor
            picDraw.DrawWidth = RubberWidth
            
            aX = CInt(X)
            aY = CInt(Y)
            
            picDraw.Line (cx, cy)-(aX, aY), colour
            picDraw.Refresh
            
            'docxcy
            tcX = cx
            tcY = cy
            
            cx = aX
            cy = aY
            
            SendLine tcX, tcY, picDraw.DrawWidth, aX, aY
            
            
            'Remember where the mouse is so new lines can be drawn connecting to this point.
            
            
            colour = TmpColour
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
        dColour = GetPixel(picDraw.hDC, ScaleX(cx, vbTwips, vbPixels), ScaleY(cy, vbTwips, vbPixels))
        
        If colour <> -1 Then
            If dColour <> -1 Then
                picColours(7).BackColor = picColour.BackColor
                picColour.BackColor = dColour
                colour = dColour
                optnDraw_Click 0 'pick colour off
            End If
        End If
        
    ElseIf DrawingStraight Then
        
        If pDrawDrawnOn = False Then pDrawDrawnOn = True
        
        If CBool(StraightPoint1.X) Or CBool(StraightPoint1.Y) Then
            StraightPoint2.X = cx
            StraightPoint2.Y = cy
            picDraw.Line (StraightPoint1.X, StraightPoint1.Y)- _
                (StraightPoint2.X, StraightPoint2.Y), colour
            
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

'If mnuOptionsMessagingColours.Checked Then
    
    Cmdlg.flags = cdlCCFullOpen + cdlCCRGBInit
    Cmdlg.Color = TxtForeGround
    
    On Error GoTo Err
    Cmdlg.ShowColor
    
    TxtForeGround = Cmdlg.Color
    
    'txtOut.ForeColor = TxtForeGround
    
'End If

Err:
End Sub

Private Sub txtOut_KeyDown(KeyCode As Integer, Shift As Integer)
'Debug.Print "KC: " & KeyCode & ", Shift: " & Shift

'shift = 2, kc = 66 - ctrl+b

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
    txtOut_Autocomplete
    
ElseIf KeyCode = vbKeyReturn Then
    'If (Shift Or vbCtrlMask) = vbCtrlMask Then
    If (Shift And vbShiftMask) = vbShiftMask Then
        'ctrl+ent (=1 for shift)
        modSpeech.Say txtOut.Text, , , True
    End If
End If

End Sub

Private Sub txtOut_Autocomplete()
Dim Txt As String
Dim iCmd As Integer, len_match As Integer

Txt = LCase$(txtOut.Text)

If LenB(Txt) Then
    If Left$(Txt, 1) = "/" Then
        Txt = Mid$(Txt, 2)
        
        If Len(Txt) > 1 Then
            If matches_ircCommand(Txt, iCmd, len_match) Then
                If len_match < Len(irc_Commands(iCmd)) Then
                    Txt = Mid$(Txt, len_match + 1)
                    If Left$(Txt, 1) <> vbSpace Then Txt = vbSpace & Txt
                    
                    txtOut.Text = "/" & irc_Commands(iCmd).sCommand & Txt
                    
                    If Status = Connected Or Not irc_Commands(iCmd).bChatMessage Then
                        cmdSend.Enabled = True 'will be disabled on _Change
                        cmdSend.Default = True
                    Else
                        cmdSend.Enabled = False
                    End If
                    txtOut.Selstart = Len(txtOut.Text)
                    SetFocus2txtOut
                End If
            End If
        End If
    End If
End If

End Sub

Private Function matches_ircCommand(Txt As String, ByRef iCommand As Integer, ByRef len_match As Integer) As Boolean
Dim i As Integer, j As Integer
Const MIN_MATCH_LEN = 2


For i = 0 To UBound(irc_Commands)
    For j = Len(irc_Commands(i).sCommand) To MIN_MATCH_LEN Step -1
        If Left$(irc_Commands(i).sCommand, j) = Left$(Txt, j) Then
            len_match = j
            iCommand = i
            matches_ircCommand = True
            Exit Function
        End If
    Next j
Next i

iCommand = -1
matches_ircCommand = False
End Function

'Private Sub txtOut_KeyDown(KeyCode As Integer, Shift As Integer)
'Call Form_KeyDown(KeyCode, Shift)
'End Sub

Public Sub txtOut_KeyPress(KeyAscii As Integer)
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
    
ElseIf KeyAscii = 13 Then
    If (GetAsyncKeyState(vbKeyShift) And &H8000) Or Status <> Connected Then
        KeyAscii = 0
    End If
    
ElseIf KeyAscii = vbKeyTab Then
    txtOut_Autocomplete
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
Dim iStart As Integer, iLen As Integer

On Error GoTo EH
If Not (Screen.ActiveControl Is Nothing) Then
    
    CtrlName = Screen.ActiveControl.Name
    
    If Not ((CtrlName = "cmdSmile") Or (CtrlName = "cmdShake")) Then
        iStart = txtOut.Selstart
        iLen = txtOut.Sellength
        
        Call LostFocus(txtOut)
        ResetTxtOutHeight
        
        On Error Resume Next
        txtOut.Selstart = iStart
        txtOut.Sellength = iLen
    End If
    
    'txtOut.Selstart = Len(txtOut.Text)
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
    
    SetFocus2txtOut
    
    PopupMenu frmSystray.mnuFont, , , , frmSystray.mnuFontColour
    
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
Rename txtName.Text, True

'No, because this prevents txtStatus from getting focus
'ResetFocus
End Sub

Public Function FormatApostrophe(ByVal sName As String) As String
FormatApostrophe = sName & "'" & IIf(Right$(sName, 1) = "s", vbNullString, "s")
End Function

Public Sub Rename(ByVal sNewName As String, Optional bLostFocus As Boolean = False)

Const MaxNameLen = 20
Dim bHadError As Boolean
Dim Msg As String, NameOp As String, sTxt As String
'aka Name to be operated on

If sNewName = LastName Then Exit Sub

NameOp = RemoveChars(sNewName)

If Trim$(NameOp) <> Trim$(sNewName) Then
    AddText "Certain characters can't be used in your name" & _
        " (" & _
        "@, #, " & modMessaging.MsgEncryptionFlag & ", :, " & modSpaceGame.mPacketSep & ", " & modSpaceGame.UpdatePacketSep & _
        ")", TxtError, True
    
    bHadError = True
End If

If LenB(NameOp) = 0 Then
    If Me.mnuOptionsMessagingDisplaySysUserName.Checked Then
        NameOp = modVars.User_Name
    Else
        NameOp = SckLC.LocalHostName
    End If
End If


If Len(NameOp) > MaxNameLen Then
    AddText "Name is Too Long - Truncated", TxtError, True
    NameOp = Left$(NameOp, MaxNameLen)
    bHadError = True
End If


NameOp = Trim$(NameOp)

If NameOp <> LastName Then
    sTxt = "Renamed to " & NameOp
    SetMiniInfo sTxt 'used later
    
    If Status = Connected Then
        
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
        
        
        AddText sTxt, , True
    Else
        SetInfo sTxt, False
    End If
    
End If

If bLostFocus Then
    If bHadError Then
        modDisplay.ShowBalloonTip txtName, "Rename", "You might want to check your name - it has been truncated/altered", TTI_WARNING
    End If
End If

LastName = NameOp
txtName.Text = NameOp 'Trim$(txtName.Text)
txtName.Selstart = Len(txtName.Text)
'Call CheckAwayChecked

End Sub

Private Sub txtOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetInfoPanel "Message Text Box"
HideExtras
End Sub

Private Sub txtStatus_GotFocus()
txtStatus.Selstart = 0
txtStatus.Sellength = Len(txtStatus.Text)
End Sub

Private Sub txtStatus_LostFocus()
ReStatus txtStatus.Text
txtStatus.Visible = False
ResetFocus
End Sub

Private Function TruncateStatus(ByVal sNew As String) As String
Do While TextWidth(sNew) > 1850
    sNew = Left$(sNew, Len(sNew) - 1)
Loop
TruncateStatus = sNew
End Function

Public Sub ReStatus(ByVal sNewStatus As String)
Dim i As Integer
Dim sNew As String, sToSend As String, sInfoText As String
Dim bTell As Boolean
Dim StatusDefault As String

If modDisplay.VisualStyle = False Then
    StatusDefault = sqB_Status_sqB
'Else
    'tb banner will take care of it
End If


sNew = RemoveChars(sNewStatus)

If sNew = StatusDefault Then
    sNew = vbNullString
End If


sNewStatus = sNew
sNew = Trim$(TruncateStatus(sNew))
If sNew <> sNewStatus Then bTell = True


If LastStatus <> sNew Then
    If LenB(sNew) Then
        txtStatus.Text = sNew
        
        
        If bTell Then
            sInfoText = "Status was too long - Shortened to '" & sNew & "'"
            SetInfo sInfoText, True
        Else
            sInfoText = "Status set to '" & sNew & "'"
        End If
        
        SetMiniInfo "Set status to: " & sNew
        
        If Status = Connected Then
            sToSend = LastName & " set their status to '" & sNew & "'"
            
'            If Server Then
'                modMessaging.DistributeMsg eCommands.Info & sToSend, -1
'            Else
'                modMessaging.SendData eCommands.Info & sToSend
'            End If
            
            AddText sInfoText, , True
        End If
        
    Else
        
        sInfoText = "Status Removed"
        SetMiniInfo sInfoText
        
        If Status = Connected Then
            
            sToSend = LastName & " removed their status"
'            If Server Then
'                modMessaging.DistributeMsg eCommands.Info & sToSend, -1
'            Else
'                modMessaging.SendData eCommands.Info & sToSend
'            End If
            AddText sInfoText, , True
        End If
    End If
    
    If LenB(sToSend) Then 'should be true
        SendInfoMessage sToSend
    End If
    
    If Status <> Connected Then
        SetInfo sInfoText, False
    End If
    
    LastStatus = sNew
    'Stick(0).sStatus set on list timer
End If

If LenB(sNew) = 0 Then
    txtStatus.Text = StatusDefault 'set above
End If

frmSystray.mnuAFK.Checked = (LCase$(LastStatus) = "afk")
frmSystray.mnuStatusAway.Checked = frmSystray.mnuAFK.Checked

End Sub

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
        ucFileTransfer.Listen modPorts.DPPort
    End If
End If


End Sub

Private Sub ucFileTransfer_ReceivedFile(sFileName As String)
Dim iClient As Integer, iSock As Integer, i As Integer
Dim FNameOnly As String
Dim sSavePath As String, sSaveName As String, sRootSaveName As String

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
    Clients(iClient).bDPIsGIF = (Right$(sFileName, 3) = "gif")
    
    ShowDP iClient
    
    
    AddText "Received New Display Picture" & _
        IIf(LenB(Clients(iClient).sName), " for " & Clients(iClient).sName, vbNullString), , True
    
    
    If mnuOptionsDPSaveAll.Checked Then
        'save pic
        On Error GoTo FCEH
        sSavePath = FT_Path() & "\Display Pictures"
        
        If FileExists(sSavePath, vbDirectory) = False Then
            MkDir sSavePath
        End If
        
        sRootSaveName = sSavePath & "\" & FormatApostrophe(Clients(iClient).sName) & " DP (" & MakeTimeFile() & ")"
        
        sSaveName = sRootSaveName & "." & IIf(Clients(iClient).bDPIsGIF, "gif", "jpg")
        
        i = 0
        While FileExists(sSaveName)
            sSaveName = sRootSaveName & vbSpace & CStr(i) & "." & IIf(Clients(iClient).bDPIsGIF, "gif", "jpg")
        Wend
        
        FileCopy sFileName, sSaveName
    End If
    
End If

EH:
Exit Sub
FCEH:
AddText "Error Saving Display Picture - " & Err.Description, TxtError, True
End Sub

'Private Sub mnuOptionsWindow2Implode_Click()
'Call AnimClick(mnuOptionsWindow2Implode)
'End Sub
'
'Private Sub mnuOptionsWindow2Slide_Click()
'Call AnimClick(mnuOptionsWindow2Slide)
'End Sub
'
'Private Sub mnuOptionsWindow2All_Click()
'Call AnimClick(mnuOptionsWindow2All)
'End Sub
'
'Private Sub mnuOptionsWindow2NoImplode_Click()
'Call AnimClick(mnuOptionsWindow2NoImplode)
'End Sub
'
'Private Sub mnuOptionsWindow2Fade_Click()
'Call AnimClick(mnuOptionsWindow2Fade)
'End Sub

'Public Sub AnimClick(Optional ByRef mnu As Menu, Optional ByVal AnimType As eAnimType = -1)
'
'mnuOptionsWindow2Implode.Checked = False
'mnuOptionsWindow2Slide.Checked = False
'mnuOptionsWindow2Fade.Checked = False
'mnuOptionsWindow2All.Checked = False
'mnuOptionsWindow2NoImplode.Checked = False
'
'If AnimType = -1 Then
'    If Not (mnu Is Nothing) Then
'        mnu.Checked = True
'    End If
'Else
'
'    Select Case AnimType
'        Case eAnimType.aImplode
'            mnuOptionsWindow2Implode.Checked = True
'        'Case eAnimType.aSlide
'            'mnuOptionsWindow2Slide.Checked = True
'        Case eAnimType.aRandom
'            mnuOptionsWindow2All.Checked = True
'        Case eAnimType.aFade
'            'mnuOptionsWindow2Fade.Checked = True
'            mnuOptionsWindow2Slide.Checked = True
'        Case eAnimType.None
'            mnuOptionsWindow2NoImplode.Checked = True
'    End Select
'
'End If
'
'End Sub

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

'##################################################################################

Private Sub cmdXPArray_Click(Index As Integer)

Select Case Index
    Case 0
        cmdListen_Click
    Case 1
        cmdClose_Click
    Case 2
        cmdAdd_Click
    Case 3
        cmdRemove_Click
    Case 4
        cmdScan_Click
    Case 5
        cmdPrivate_Click
End Select

End Sub
Private Sub cmdArray_Click(Index As Integer)
cmdXPArray_Click Index
End Sub

Private Sub cmdXPArray_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    'Case 0
        'cmdListen_Click
    'Case 1
        'cmdclose_MouseDown Button, Shift, X, Y
    Case 2
        cmdAdd_MouseDown Button, Shift, X, Y
    'Case 3
        'cmdremove_MouseDown Button, Shift, X, Y
    'Case 4
        'cmdscan_MouseDown Button, Shift, X, Y
    'Case 5
        'cmdPrivate_Click
End Select
End Sub
Private Sub cmdArray_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdXPArray_MouseDown Index, Button, Shift, X, Y
End Sub


Private Sub cmdXPArray_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        cmdListen_MouseMove Button, Shift, X, Y
    Case 1
        cmdClose_MouseMove Button, Shift, X, Y
    Case 2
        cmdAdd_MouseMove Button, Shift, X, Y
    Case 3
        cmdRemove_MouseMove Button, Shift, X, Y
    Case 4
        cmdScan_MouseMove Button, Shift, X, Y
    'Case 5
        'cmdprivate_MouseMove Button, Shift, X, Y
End Select
End Sub
Private Sub cmdArray_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdXPArray_MouseMove Index, Button, Shift, X, Y
End Sub

Private Sub cmdNormCls_Click()
cmdCls_Click
End Sub

Private Sub cmdxpCls_Click()
cmdNormCls_Click
End Sub

Private Sub SetCmdButtons()
Dim bFancy As Boolean
Dim i As Integer

bFancy = modDisplay.CanShow_XPButtons()

For i = 0 To cmdArray.UBound
    cmdArray(i).Visible = Not bFancy
    cmdXPArray(i).Visible = bFancy
Next i


cmdNormCls.Visible = Not bFancy
cmdXPCls.Visible = bFancy


cmdReply(0).Visible = False: cmdReply(1).Visible = False
cmdXPReply(0).Visible = False: cmdXPReply(1).Visible = False

End Sub

Public Sub EnableCmd(i As Integer, Optional ByVal bEn As Boolean = True)

cmdArray(i).Enabled = bEn
cmdXPArray(i).Enabled = bEn

End Sub
Public Sub EnableCmdCls(ByVal bEn As Boolean)
cmdXPCls.Enabled = bEn
cmdNormCls.Enabled = bEn
End Sub

Private Sub ShowCmdReplys(Optional bShow As Boolean = True)

If modDisplay.CanShow_XPButtons() Then
    cmdXPReply(0).Visible = bShow
    cmdXPReply(1).Visible = bShow
    
    If bShow Then
        cmdXPReply(1).Default = True
        
        SetFocus2 cmdXPReply(1)
    End If
Else
    cmdReply(0).Visible = bShow
    cmdReply(1).Visible = bShow
    
    If bShow Then
        cmdReply(1).Default = True
        
        SetFocus2 cmdReply(1)
    End If
End If

End Sub

'##################################################################################

Public Sub SetIcon(ByVal St As eStatus)

pSetIcon CInt(St) + 1, Me.mnuFileGameMode.Checked

End Sub

Private Sub pSetIcon(ByVal iImg As Integer, _
    Optional pbGameMode As Boolean = False)

Dim lhWndTop As Long, lhWnd As Long, lHandle As Long
Dim bDev As Boolean, bAdvDev As Boolean

lHandle = modDev.getDevLevel()
bDev = lHandle > modDev.Dev_Level_None
bAdvDev = lHandle > modDev.Dev_Level_Normal

If pbGameMode Then
    imgStatus.Picture = frmSystray.imgGame.ListImages(iImg).Picture
ElseIf bAdvDev Then
    imgStatus.Picture = frmSystray.imgUberDev.ListImages(iImg).Picture
ElseIf bDev Then
    imgStatus.Picture = frmSystray.imgDev.ListImages(iImg).Picture
Else
    imgStatus.Picture = frmSystray.img32x32.ListImages(iImg).Picture
End If


If bAdvDev Then
    lHandle = frmSystray.img16x16UberDev.ListImages(iImg).Picture.Handle
ElseIf bDev Then
    lHandle = frmSystray.img16x16Dev.ListImages(iImg).Picture.Handle
Else
    lHandle = frmSystray.img16x16.ListImages(iImg).Picture.Handle
End If

'frmMain.Icon = frmSystray.img48x48.ListImages(iImg).Picture
SendMessageByLong frmMain.hWnd, WM_SETICON, ICON_SMALL, lHandle
frmSystray.IconHandle = lHandle


lhWnd = Me.hWnd
lhWndTop = lhWnd
Do While lhWnd > 0
    lhWnd = GetWindow(lhWnd, GW_OWNER)
    
    If lhWnd > 0 Then
        lhWndTop = lhWnd
    End If
Loop
SendMessageByLong lhWndTop, WM_SETICON, ICON_BIG, imgStatus.Picture.Handle
'SendMessageByLong lhWndTop, WM_SETICON, ICON_SMALL, imgStatus.Picture.Handle

End Sub

'these are for FormLoad - Loading sub forms
Public Property Get ConnectedIcon() As IPictureDisp
'Set ConnectedIcon = frmSystray.img32x32.ListImages(2).Picture
Set ConnectedIcon = frmSystray.img16x16.ListImages(2).Picture
End Property
Public Property Get IdleIcon() As IPictureDisp
'Set IdleIcon = frmSystray.img32x32.ListImages(1).Picture
Set IdleIcon = frmSystray.img16x16.ListImages(1).Picture
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

On Error GoTo EH 'in case text is dragged over

ReDim Files(Data.Files.Count - 1)

For i = 0 To Data.Files.Count - 1
    Files(i) = Data.Files(i + 1)
    'convert from 1-based to 0-based, and convert type (in a way)
Next i

DragDrop Files

Exit Sub
EH:
SetInfo "That's not a file...", True
End Sub

Private Sub txtOut_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, state As Integer)
'SetInfo "Drag over the textbox to open file transfer", False
rtfIn_OLEDragOver Nothing, Effect, Button, Shift, X, Y, state
End Sub
Private Sub rtfIn_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, state As Integer)
SetInfo "Drag over the textbox to open file transfer", False
End Sub

Private Sub txtOut_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Files() As String
Dim i As Integer

On Error GoTo EH
ReDim Files(Data.Files.Count - 1)

For i = 0 To Data.Files.Count - 1
    Files(i) = Data.Files(i + 1)
    'convert from 1-based to 0-based, and convert type (in a way)
Next i

DragDrop Files

Exit Sub
EH:
If Err.Number = 461 Then
    SetInfo "Error: Can't read fancy address data (VB limitation)", True
Else
    SetInfo "Error (" & Err.Number & "): " & Err.Description, True
End If
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
        SetInfo "You can only send one file at once", True
    Else
        
        If modLoadProgram.frmManualFT_Loaded = False Then
            mnuOptionsMessagingWindowsFT_Click
        End If
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

Select Case LCase$(Right$(sPath, 4))
    Case ".jpg", ".bmp", ".jpeg", ".gif"
        IsDP_File = True
    Case Else
        IsDP_File = False
End Select

End Function


Private Sub imgStatus_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Status = Connected Then
    If Data.Files.Count > 1 Then
        SetInfo "You can only have one Display Picture...", True
        
    ElseIf IsDP_File(Data.Files(1)) Then
        SetMyDP Data.Files(1)
        
    Else
        SetInfo "That picture must be a jpg, .bmp, .jpeg or .gif", True
        
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
Private Sub imgStatus_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, state As Integer)
SetInfo "Drag Over to Set Display Picture", False
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, state As Integer)
imgStatus_OLEDragOver Data, Effect, Button, Shift, X, Y, state
End Sub

Private Sub txtName_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, state As Integer)
imgStatus_OLEDragOver Data, Effect, Button, Shift, X, Y, state
End Sub

Private Sub imgDP_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, state As Integer)
imgStatus_OLEDragOver Data, Effect, Button, Shift, X, Y, state
End Sub
