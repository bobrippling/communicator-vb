VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6975
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   9600
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   6840
      TabIndex        =   65
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Frame fraPaths 
      Caption         =   "Paths"
      Height          =   855
      Left            =   240
      TabIndex        =   53
      Top             =   5400
      Width           =   9255
      Begin VB.PictureBox picPaths 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9015
         TabIndex        =   54
         Top             =   360
         Width           =   9015
         Begin VB.PictureBox picLogs 
            Height          =   375
            Left            =   8640
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   61
            Top             =   0
            Width           =   375
            Begin VB.CommandButton cmdLogs 
               Caption         =   "..."
               Height          =   255
               Left            =   0
               TabIndex        =   62
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.TextBox txtLogs 
            Height          =   285
            Left            =   5520
            TabIndex        =   59
            Top             =   0
            Width           =   3015
         End
         Begin VB.PictureBox picBrowse 
            Height          =   375
            Left            =   3960
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   57
            Top             =   0
            Width           =   375
            Begin VB.CommandButton cmdBrowseFTP 
               Caption         =   "..."
               Height          =   255
               Left            =   0
               TabIndex        =   58
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.TextBox txtFTPLocalLoc 
            Height          =   285
            Left            =   1560
            TabIndex        =   56
            Top             =   0
            Width           =   2295
         End
         Begin VB.Label lblLogs 
            Caption         =   "Log Path:"
            Height          =   255
            Left            =   4800
            TabIndex        =   60
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblFTPLoc 
            Caption         =   "Received Files Path:"
            Height          =   255
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   1575
         End
      End
   End
   Begin MSComDlg.CommonDialog CDColour 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Pick a Colour"
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "Reload"
      Height          =   375
      Left            =   5400
      TabIndex        =   64
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Frame fraRest 
      Caption         =   "Misc"
      Height          =   5175
      Left            =   5040
      TabIndex        =   28
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtVoicePort 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   40
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtDPPort 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   38
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtFTPort 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   36
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtSpacePort 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   34
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtStickPort 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         Top             =   840
         Width           =   1215
      End
      Begin VB.PictureBox picPort 
         Height          =   3435
         Left            =   3000
         ScaleHeight     =   3375
         ScaleWidth      =   1275
         TabIndex        =   41
         Top             =   300
         Width           =   1335
         Begin VB.CommandButton cmdDefVoice 
            Caption         =   "Reset"
            Height          =   375
            Left            =   0
            TabIndex        =   47
            Top             =   2400
            Width           =   1215
         End
         Begin VB.CommandButton cmdDefDP 
            Caption         =   "Reset"
            Height          =   375
            Left            =   0
            TabIndex        =   46
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdDefFT 
            Caption         =   "Reset"
            Height          =   375
            Left            =   0
            TabIndex        =   45
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdDefSpace 
            Caption         =   "Reset"
            Height          =   375
            Left            =   0
            TabIndex        =   44
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdDefStick 
            Caption         =   "Reset"
            Height          =   375
            Left            =   0
            TabIndex        =   43
            Top             =   480
            Width           =   1215
         End
         Begin VB.CommandButton cmdDefMain 
            Caption         =   "Reset"
            Height          =   375
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdDisconnect 
            Caption         =   "Disconnect"
            Height          =   375
            Left            =   0
            TabIndex        =   48
            Top             =   2880
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.PictureBox picFont 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   945
         ScaleWidth      =   4185
         TabIndex        =   50
         Top             =   4080
         Width           =   4215
         Begin VB.CommandButton cmdResetFont 
            Caption         =   "Reset Font"
            Height          =   375
            Left            =   2400
            TabIndex        =   52
            Top             =   560
            Width           =   1695
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Select Font"
            Height          =   375
            Left            =   2400
            TabIndex        =   51
            Top             =   10
            Width           =   1695
         End
      End
      Begin VB.TextBox txtMainPort 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblPort 
         Caption         =   "Recording Port"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lblPort 
         Caption         =   "Display Picture Port"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblPort 
         Caption         =   "File Transfer Port"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblPort 
         Caption         =   "Space Game Port"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblPort 
         Caption         =   "Stick Game Port"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line lnMiscSep 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   120
         X2              =   4320
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label lblConnectionInfo 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   49
         Top             =   3120
         Width           =   2805
      End
      Begin VB.Label lblPort 
         Caption         =   "Main Port"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   735
      End
      Begin VB.Line lnMiscSep 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   120
         X2              =   4320
         Y1              =   3840
         Y2              =   3840
      End
   End
   Begin VB.Frame fraColours 
      Caption         =   "Colours"
      Height          =   5175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4815
      Begin VB.PictureBox picCmd 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         ScaleHeight     =   3975
         ScaleWidth      =   4455
         TabIndex        =   8
         Top             =   960
         Width           =   4455
         Begin VB.TextBox txtColour 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   480
            Width           =   2535
         End
         Begin VB.PictureBox picBG 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3000
            ScaleHeight     =   255
            ScaleWidth      =   1185
            TabIndex        =   25
            Top             =   2670
            Width           =   1215
         End
         Begin VB.TextBox txtColour 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   120
            Width           =   2535
         End
         Begin VB.TextBox txtColour 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   870
            Width           =   2535
         End
         Begin VB.TextBox txtColour 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1230
            Width           =   2535
         End
         Begin VB.TextBox txtColour 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1590
            Width           =   2535
         End
         Begin VB.TextBox txtColour 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   2310
            Width           =   2535
         End
         Begin VB.TextBox txtColour 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   2670
            Width           =   1215
         End
         Begin VB.TextBox txtColour 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1950
            Width           =   2535
         End
         Begin VB.CommandButton cmdDefaultColours 
            Caption         =   "Default Colours"
            Height          =   375
            Left            =   480
            TabIndex        =   27
            Top             =   3600
            Width           =   3375
         End
         Begin VB.CommandButton cmdPick 
            Caption         =   "Pick Colour for ..."
            Height          =   375
            Left            =   480
            TabIndex        =   26
            Top             =   3120
            Width           =   3375
         End
         Begin VB.Label lblColour 
            Alignment       =   1  'Right Justify
            Caption         =   "Sent Messages"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   12
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblColour 
            Alignment       =   1  'Right Justify
            Caption         =   "Recieved Messages"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Top             =   150
            Width           =   1575
         End
         Begin VB.Label lblColour 
            Alignment       =   1  'Right Justify
            Caption         =   "Other Sent Messages"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   11
            Top             =   870
            Width           =   1575
         End
         Begin VB.Label lblColour 
            Alignment       =   1  'Right Justify
            Caption         =   "Information Messages"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   14
            Top             =   1230
            Width           =   1575
         End
         Begin VB.Label lblColour 
            Alignment       =   1  'Right Justify
            Caption         =   "Error Messages"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   16
            Top             =   1590
            Width           =   1575
         End
         Begin VB.Label lblColour 
            Alignment       =   1  'Right Justify
            Caption         =   "Unknown/Other"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   20
            Top             =   2310
            Width           =   1575
         End
         Begin VB.Label lblColour 
            Alignment       =   1  'Right Justify
            Caption         =   "BackGround"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   24
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label lblColour 
            Alignment       =   1  'Right Justify
            Caption         =   "Questions"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   18
            Top             =   1950
            Width           =   1575
         End
      End
      Begin VB.Label lblNote 
         Alignment       =   2  'Center
         Caption         =   "notage"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   63
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8280
      TabIndex        =   66
      Top             =   6480
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20115
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20115
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   3
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20115
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   1
         Top             =   300
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private txtGotFocus As Integer
Private Const PickCap As String = "Pick Colour for "
Private Const SlideCap As String = "Drawing Box Height: "

Private TempFontName As String, TempFontSize As Single
Private TempBold As Boolean, TempItalic As Boolean

Private Sub cmdBrowseFTP_Click()

Dim TmpPath As String

TmpPath = Trim$(modDirBrowse.BrowseFolder(Me.hWnd, "Received Files Path"))

If LenB(TmpPath) Then
    txtFTPLocalLoc.Text = TmpPath
    Setting_Changed True
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDefaultColours_Click()
ShowDefaultColours
End Sub

Private Sub ShowDefaultColours()

txtColour(0).Text = FillHex(MOrange)
txtColour(1).Text = FillHex(vbBlue)
txtColour(2).Text = FillHex(MGreen)
txtColour(3).Text = FillHex(vbRed)
txtColour(4).Text = FillHex(MBrown)
txtColour(5).Text = FillHex(vbBlack)
txtColour(6).Text = FillHex(vbWhite)
txtColour(7).Text = FillHex(MPurple)

End Sub

Private Sub cmdDisconnect_Click()

'frmMain.CleanUp True
frmMain.cmdClose_Click

lblConnectionInfo.Visible = False
txtMainPort.Enabled = True
cmdDisconnect.Visible = False
txtMainPort_Change
End Sub

Private Sub cmdFont_Click()

With CDColour
    .FontName = TempFontName
    .FontSize = TempFontSize
    .FontItalic = TempItalic
    .FontBold = TempBold
    
    .flags = cdlCFForceFontExist Or cdlCFLimitSize Or cdlCFBoth
    .Max = MaxFont
    .Min = MinFont
    
    On Error GoTo EH
    .ShowFont
    
    TempFontName = .FontName
    TempFontSize = .FontSize
    TempItalic = .FontItalic
    TempBold = .FontBold
    cmdResetFont.Enabled = True
End With

ShowFontEG
Setting_Changed True

EH:
End Sub

Private Sub cmdLogs_Click()

Dim TmpPath As String

TmpPath = Trim$(modDirBrowse.BrowseFolder(Me.hWnd, "Log Path"))

If LenB(TmpPath) Then
    txtLogs.Text = TmpPath
    Setting_Changed True
End If

End Sub

Private Sub cmdReload_Click()
LoadAll
Setting_Changed False
End Sub

Private Sub cmdResetFont_Click()

TempFontName = frmMain.rtfFontName
TempFontSize = frmMain.rtfFontSize
TempBold = frmMain.rtfBold
TempItalic = frmMain.rtfItalic

ShowFontEG

cmdResetFont.Enabled = False

End Sub

Private Sub ShowFontEG()
Dim sFontText As String

With picFont
    '.AutoRedraw = True
    .Cls
    
    
    .FontName = TempFontName
    .FontSize = TempFontSize
    .FontBold = TempBold
    .FontItalic = TempItalic
    
    sFontText = TempFontName & " Text"
    
    .CurrentX = 10
    .CurrentY = .ScaleHeight / 2 - .TextHeight(sFontText) / 2
    
    picFont.Print sFontText
End With

End Sub

Private Sub cmdOk_Click()
Dim Cancel As Boolean

SetAll Cancel, True

If Not Cancel Then Unload Me

End Sub

Private Sub cmdApply_Click()
Dim bCancel As Boolean

SetAll bCancel, False

If Not bCancel Then
    modSpeech.Say "Settings applied"
    Setting_Changed False
End If

End Sub

Private Sub cmdPick_Click()
Dim Erro As Boolean

'Md1.CDColour erro, txtColour(txtGotFocus).Text, CDColour

CDColour.flags = cdlCCFullOpen + cdlCCRGBInit
If txtGotFocus <> 6 Then
    CDColour.Color = txtColour(txtGotFocus).ForeColor
Else
    CDColour.Color = txtColour(txtGotFocus).BackColor
End If

On Error GoTo Err
CDColour.ShowColor

txtColour(txtGotFocus).Text = FillHex(CDColour.Color)

modDisplay.ShowBalloonTip txtColour(txtGotFocus), "Colour Set", "Colour set as " & txtColour(txtGotFocus).Text

Err:
End Sub

Private Sub Form_Load()
Const BarLow As Integer = 1680
Const BarHigh As Integer = 1440

Me.Left = Screen.width / 2 - Me.width / 2 'frmMain.Left + frmMain.width / 2 - Me.width / 2
Me.Top = Screen.height / 2 - Me.height / 2 'frmMain.Top + frmMain.height / 2 - Me.height / 2

lblNote.Caption = "Note: Changes will only apply to new text, not old text" & vbNewLine & _
                  "'Sent Messages' is for typed messages," & vbNewLine & "'Other Sent Messages' is for shakes, etc"

lblConnectionInfo.BorderStyle = ccNone
picFont.BorderStyle = ccNone
picPort.BorderStyle = ccNone
picBrowse.BorderStyle = ccNone
picLogs.BorderStyle = ccNone

Me.txtFTPLocalLoc.Locked = False
Me.txtLogs.Locked = False

LoadAll
'ShowFontEG <-- done above

txtColour_GotFocus 0

If frmMain.mnuOptionsAdvDisplayVistaControls.Checked Then
    modDisplay.SetButtonIcon cmdOk.hWnd, frmMain.GetCommandIconHandle()
End If

cmdOk.Default = True
cmdOk.Enabled = True 'let them OK out of it first time
cmdReload.Enabled = False
cmdApply.Enabled = False
Call FormLoad(Me, , , False)

End Sub

Private Sub Setting_Changed(bDifferentFromSaved As Boolean)
cmdOk.Enabled = bDifferentFromSaved
cmdApply.Enabled = bDifferentFromSaved
cmdReload.Enabled = bDifferentFromSaved
cmdCancel.Caption = IIf(bDifferentFromSaved, "Cancel", "Close")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call frmMain.SetInactive
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FormLoad Me, True
End Sub

Private Sub lblColour_Click(Index As Integer)
SetFocus2 txtColour(Index)
End Sub

Private Sub lblColour_DblClick(Index As Integer)
Call lblColour_Click(Index)
End Sub

Private Sub lblLogs_DblClick()
OpenFolder vbNormalFocus, txtLogs.Text
End Sub

Private Sub lblFTPLoc_DblClick()
OpenFolder vbNormalFocus, txtFTPLocalLoc.Text
End Sub

Private Sub picBG_Click()
txtColour_GotFocus 6
End Sub

'Private Sub slideDraw_Change()
''Const Max As Integer = 4500
''Const Min As Integer = 3000
'Static Ignore As Boolean
'
'If Not Ignore Then
'    Ignore = True
'    slideDraw.Value = (Left$(CStr(slideDraw.Value), 3)) * 10
'
'    lblSlide.Caption = SlideCap & Trim$(Str$(slideDraw.Value))
'    Ignore = False
'End If
'End Sub

Private Sub txtColour_Change(Index As Integer)
Dim lCol As Long: lCol = CNorm(txtColour(Index).Text)

If Index = 6 Then
    Me.picBG.BackColor = lCol
Else
    txtColour(Index).ForeColor = CNorm(txtColour(Index).Text)
End If

Setting_Changed True
Call frmMain.SetInactive
End Sub

Private Sub txtColour_KeyPress(Index As Integer, KeyAscii As Integer)
Beep
End Sub

Private Sub txtColour_GotFocus(Index As Integer)
txtGotFocus = Index
cmdPick.Caption = PickCap & lblColour(Index).Caption
cmdPick.Default = True
End Sub

Private Sub LoadAll()

txtColour(0).Text = FillHex(TxtReceived)
txtColour(1).Text = FillHex(TxtSent)
txtColour(2).Text = FillHex(TxtInfo)
txtColour(3).Text = FillHex(TxtError)
txtColour(4).Text = FillHex(TxtQuestion)
txtColour(5).Text = FillHex(TxtUnknown)
txtColour(6).Text = FillHex(TxtBackGround)
txtColour(7).Text = FillHex(TxtForeGround)


txtMainPort.Text = CStr(MainPort)
txtStickPort.Text = CStr(StickPort)
txtSpacePort.Text = CStr(SpacePort)
txtFTPort.Text = CStr(FTPort)
txtDPPort.Text = CStr(DPPort)
txtVoicePort.Text = CStr(VoicePort)

'slideDraw.Value = frmMain.DrawHeight
'txtUpdate.Text = modVars.UpdateURL
'lblUpdateInfo.Caption = modVars.UpdateURL & " will be checked for an update."

'For i = 1 To Screen.FontCount
'    If LenB(Screen.Fonts(i)) <> 0 Then
'        cboFont.AddItem Screen.Fonts(i)
'    End If
'Next i

''list alphabetically
'Dim Items() As String
'Dim i As Integer, j As Integer
'Dim Temp As String
''Put the items in a variant array
'ReDim Items(0 To Screen.FontCount - 1)
'
'For i = 0 To Screen.FontCount - 1
'    Items(i) = Screen.Fonts(i)
'Next i
'
''http://www.xtremevbtalk.com/showthread.php?t=279869
'For i = LBound(Items) To UBound(Items)
'    For j = i + 1 To UBound(Items)
'        If Items(i) > Items(j) Then
'            Temp = Items(i)
'            Items(i) = Items(j)
'            Items(j) = Temp
'        End If
'    Next j
'Next i
'
''Clear the listbox
'
''Add the sorted array back to the listbox
'For i = LBound(Items) To UBound(Items)
'    cboFont.AddItem Items(i)
'Next i
''end -----------------------------------------------

cmdResetFont_Click

'-----------------------------------------------

txtFTPLocalLoc.Text = Trim$(modPaths.SavedFilesPath)
txtFTPLocalLoc.Selstart = Len(txtFTPLocalLoc.Text)
txtLogs.Text = modPaths.logPath
txtLogs.Selstart = Len(txtLogs.Text)

If Status <> Idle Then
    txtMainPort.Enabled = False
    cmdDefMain.Enabled = False
    
    cmdDisconnect.Visible = True
    cmdDisconnect.Default = True
    
    lblConnectionInfo.Caption = "You cannot change the port because you are connecting, listening, or connected."
    
Else
    lnMiscSep(0).Y1 = 3120
    lnMiscSep(0).Y2 = lnMiscSep(0).Y1
    lnMiscSep(1).Y1 = lnMiscSep(0).Y1
    lnMiscSep(1).Y2 = lnMiscSep(0).Y1
    picPort.height = 2835
    lblConnectionInfo.Visible = False
End If
End Sub

Private Function AttemptToChangePort(ByVal sPort As String, ByRef iCurrent As Integer, iDefault As Integer, sPortName As String) As Boolean
Dim Ans As VbMsgBoxResult

If LenB(sPort) Then
    If IsNumeric(sPort) Then
        If val(sPort) <> iCurrent Then
            If val(sPort) >= 1 And val(sPort) <= 65535 Then
                
                'only if they haven't already changed RPort
                If val(sPort) <> iDefault And iCurrent = iDefault Then
                    Ans = MsgBoxEx("Changing the " & sPortName & " from " & CStr(iDefault) & vbNewLine & _
                        "can prevent you from connecting to other Communicators." & vbNewLine & "Are you sure you want to change it?", _
                        "Communicator uses Port " & CStr(iDefault) & " to connect to other Communicators. " & _
                        "Changing this port will prevent you connecting or hosting with others", _
                        vbYesNo + vbQuestion, "Change " & sPortName) ', , , frmMain.Icon)
                    
                Else
                    Ans = vbYes
                End If
                
                If Ans = vbYes Then
                    iCurrent = val(sPort)
                    'ByRef
                    
                    AttemptToChangePort = True
                End If
                
            Else
                MsgBoxEx "Error - " & sPortName & " must be between 1 and 65535", "Port Numbers in Windows must be between 1 and 65535", vbExclamation, sPortName, , , frmMain.Icon, , Me.hWnd
            End If
        Else
            AttemptToChangePort = True
        End If
    Else
        MsgBoxEx "Error - " & sPortName & " must be a number", "Hello? Port NUMBER. It must be a number...", vbExclamation, sPortName, , , frmMain.Icon, , Me.hWnd
    End If
Else
    MsgBoxEx "Please enter something for the " & sPortName, "Hello? Enter a number, if you'd be so kind", vbExclamation, sPortName, , , frmMain.Icon, , Me.hWnd
End If

End Function

Private Sub SetAll(ByRef Cancel As Boolean, bClosingWindow As Boolean)
Dim sPath As String
Dim Ans As VbMsgBoxResult
Dim iTmpPort As Integer
Dim bChangedPath As Boolean

If modVars.Status = Idle Then
    iTmpPort = modPorts.MainPort 'it's a property let/get so...
    If AttemptToChangePort(txtMainPort.Text, iTmpPort, modPorts.DefaultMainPort, "Main Port") = False Then
        Cancel = True
        Exit Sub
    Else
        modPorts.MainPort = iTmpPort
    End If
End If
iTmpPort = modPorts.StickPort 'it's a property let/get so...
If AttemptToChangePort(txtStickPort.Text, iTmpPort, modPorts.DefaultStickPort, "Stick Game Port") = False Then
    Cancel = True
    Exit Sub
Else
    modPorts.StickPort = iTmpPort
End If
iTmpPort = modPorts.SpacePort 'it's a property let/get so...
If AttemptToChangePort(txtSpacePort.Text, iTmpPort, modPorts.DefaultSpacePort, "Space Port") = False Then
    Cancel = True
    Exit Sub
Else
    modPorts.SpacePort = iTmpPort
End If
iTmpPort = modPorts.FTPort 'it's a property let/get so...
If AttemptToChangePort(txtFTPort.Text, iTmpPort, modPorts.DefaultFTPort, "File Transfer Port") = False Then
    Cancel = True
    Exit Sub
Else
    modPorts.FTPort = iTmpPort
End If
iTmpPort = modPorts.DPPort 'it's a property let/get so...
If AttemptToChangePort(txtDPPort.Text, iTmpPort, modPorts.DefaultDPPort, "Display Picture Port") = False Then
    Cancel = True
    Exit Sub
Else
    modPorts.DPPort = iTmpPort
    'apply properly
    frmMain.ucFileTransfer.Disconnect 'it'll be listened by modDP
End If
iTmpPort = modPorts.VoicePort 'it's a property let/get so...
If AttemptToChangePort(txtVoicePort.Text, iTmpPort, modPorts.DefaultVoicePort, "Recording Port") = False Then
    Cancel = True
    Exit Sub
Else
    modPorts.VoicePort = iTmpPort
    'apply properly
    frmMain.ucVoiceTransfer.Disconnect 'it'll be listened by modDP
End If


TxtReceived = CNorm(txtColour(0).Text)
TxtSent = CNorm(txtColour(1).Text)
TxtInfo = CNorm(txtColour(2).Text)
TxtError = CNorm(txtColour(3).Text)
TxtQuestion = CNorm(txtColour(4).Text)
TxtUnknown = CNorm(txtColour(5).Text)
TxtBackGround = CNorm(txtColour(6).Text)
TxtForeGround = CNorm(txtColour(7).Text)


frmMain.rtfFontName = TempFontName
frmMain.rtfFontSize = TempFontSize
frmMain.rtfBold = TempBold
frmMain.rtfItalic = TempItalic


sPath = Trim$(txtFTPLocalLoc.Text)
If LenB(sPath) Then
    If FileExists(sPath, vbDirectory) Then
        If Right$(sPath, 1) = "\" Then
            sPath = Left$(sPath, Len(sPath) - 1)
        End If
        modPaths.SavedFilesPath = sPath
    Else
        If MsgBoxEx("FTP File Transfer Location Doesn't Exist" & vbNewLine & _
                    "Reset" & IIf(bClosingWindow, " + Close Window", vbNullString) & "?", _
                    "Hello thar", _
                    vbExclamation + vbYesNo, "Error", , , frmMain.Icon) = vbYes Then
            
            'reset
            modPaths.default_savedFilesPath
            Me.txtFTPLocalLoc.Text = modPaths.SavedFilesPath
            bChangedPath = True
        Else
            Cancel = True
        End If
    End If
End If

sPath = Trim$(txtLogs.Text)
If LenB(sPath) Then
    If FileExists(sPath, vbDirectory) Then
        If Right$(sPath, 1) = "\" Then
            sPath = Left$(sPath, Len(sPath) - 1)
        End If
        modPaths.logPath = sPath
    Else
        If MsgBoxEx("Log Path Doesn't Exist" & vbNewLine & _
                    "Reset path" & IIf(bClosingWindow, " + Close Window", vbNullString) & "?", _
                    "Hi", _
                    vbExclamation + vbYesNo, "Error", , , frmMain.Icon) = vbYes Then
            
            modPaths.default_logPath
            Me.txtLogs.Text = modPaths.logPath
            bChangedPath = True
        Else
            Cancel = True
        End If
    End If
End If

If bChangedPath Then
    frmMain.resetLogPath
End If

End Sub

Private Function CNorm(ByVal sHex As String) As Long
On Error Resume Next
CNorm = Abs(val("&H" & sHex & "&"))
End Function

Private Function FillHex(ByVal Col As Long) As String
Dim sHex As String
sHex = Hex$(Col)

Do While Len(sHex) < 6
    sHex = "0" & sHex
Loop

FillHex = sHex
End Function

'####################################################################################
'####################################################################################

Private Sub cmdDefMain_Click()
txtMainPort.Text = modPorts.DefaultMainPort
cmdDefMain.Enabled = False
End Sub
Private Sub cmdDefStick_Click()
txtStickPort.Text = modPorts.DefaultStickPort
cmdDefStick.Enabled = False
End Sub
Private Sub cmdDefSpace_Click()
txtSpacePort.Text = modPorts.DefaultSpacePort
cmdDefSpace.Enabled = False
End Sub
Private Sub cmdDefFT_Click()
txtFTPort.Text = modPorts.DefaultFTPort
cmdDefFT.Enabled = False
End Sub
Private Sub cmdDefDP_Click()
txtDPPort.Text = modPorts.DefaultDPPort
cmdDefDP.Enabled = False
End Sub
Private Sub cmdDefVoice_Click()
txtVoicePort.Text = modPorts.DefaultVoicePort
cmdDefVoice.Enabled = False
End Sub

Private Sub txtFTPLocalLoc_Change()
Setting_Changed True
End Sub

Private Sub txtLogs_Change()
Setting_Changed True
End Sub

Private Sub txtMainPort_Change()
On Error Resume Next
cmdDefMain.Enabled = (val(txtMainPort.Text) <> modPorts.DefaultMainPort)
Setting_Changed True
End Sub
Private Sub txtStickPort_Change()
On Error Resume Next
cmdDefStick.Enabled = (val(txtStickPort.Text) <> modPorts.DefaultStickPort)
Setting_Changed True
End Sub
Private Sub txtSpacePort_Change()
On Error Resume Next
cmdDefSpace.Enabled = (val(txtSpacePort.Text) <> modPorts.DefaultSpacePort)
Setting_Changed True
End Sub
Private Sub txtFTPort_Change()
On Error Resume Next
cmdDefFT.Enabled = (val(txtFTPort.Text) <> modPorts.DefaultFTPort)
Setting_Changed True
End Sub
Private Sub txtDPPort_Change()
On Error Resume Next
cmdDefDP.Enabled = (val(txtDPPort.Text) <> modPorts.DefaultDPPort)
Setting_Changed True
End Sub
Private Sub txtVoicePort_Change()
On Error Resume Next
cmdDefVoice.Enabled = (val(txtVoicePort.Text) <> modPorts.DefaultVoicePort)
Setting_Changed True
End Sub
