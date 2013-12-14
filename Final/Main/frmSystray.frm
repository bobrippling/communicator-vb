VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSystray 
   Caption         =   "Systray [Caption Set Programatically]"
   ClientHeight    =   4215
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7935
   Icon            =   "frmSystray.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSystray.frx":038A
   ScaleHeight     =   4215
   ScaleWidth      =   7935
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.ImageList img32x32 
      Left            =   840
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":06CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":13A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":2080
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":2D5A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGame 
      Left            =   120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":3A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":470E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":53E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":60C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDev 
      Left            =   840
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":6D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":7A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":8750
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":942A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgButton 
      Left            =   3360
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":A104
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img48x48 
      Left            =   120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":ADDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":CAB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":E792
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":1046C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16x16 
      Left            =   1560
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":12146
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":124E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":1287A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":12C14
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16x16Dev 
      Left            =   840
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":12FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":13348
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":136E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":13A7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgUberDev 
      Left            =   1560
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":13E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":14AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":157CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":164A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16x16UberDev 
      Left            =   1560
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":1717E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":17518
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":178B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSystray.frx":17C4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuPopupShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuPopupHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuPopupFolder 
         Caption         =   "Open Folder"
      End
      Begin VB.Menu mnuPopupNew 
         Caption         =   "Open New Communicator..."
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupScan 
         Caption         =   "Network Scan"
      End
      Begin VB.Menu mnuPopupHost 
         Caption         =   "Host"
      End
      Begin VB.Menu mnuPopupCloseAndHost 
         Caption         =   "Close Socket and Host"
      End
      Begin VB.Menu mnuPopupCloseC 
         Caption         =   "Close Socket"
      End
      Begin VB.Menu mnuPopupSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupGameMode 
         Caption         =   "Game Mode"
      End
      Begin VB.Menu mnuPopupSingleClick 
         Caption         =   "Single Click Tray Icon"
      End
      Begin VB.Menu mnuAFK 
         Caption         =   "Set as AFK"
      End
      Begin VB.Menu mnuPopupSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupUpdates 
         Caption         =   "Check For Updates"
      End
      Begin VB.Menu mnuPopupWhatsnew 
         Caption         =   "What's new in this version?"
      End
      Begin VB.Menu mnuPopupSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupExit 
         Caption         =   "Exit"
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
         Caption         =   "(Re)Obtain External IP"
      End
      Begin VB.Menu mnuSBObtainLocal 
         Caption         =   "(Re)Obtain Internal IP"
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
      Begin VB.Menu mnuFontSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFontCopy 
         Caption         =   "Copy"
      End
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Set Status"
      Begin VB.Menu mnuStatusAway 
         Caption         =   "AFK Status"
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
      Begin VB.Menu mnuCommandsTestSpeech 
         Caption         =   "Test Speech"
      End
      Begin VB.Menu mnuCommandsStopSpeech 
         Caption         =   "Stop Speech"
      End
      Begin VB.Menu mnuCommandsSpeech 
         Caption         =   "Speech Tags"
         Begin VB.Menu mnuCommandsSpeechActors 
            Caption         =   "Actors"
            Begin VB.Menu mnuCommandsSpeechActorsHL 
               Caption         =   "Half Life 2 Announcer"
            End
            Begin VB.Menu mnuCommandsSpeechActorsJW 
               Caption         =   "High Voice Jimmy"
            End
         End
         Begin VB.Menu mnuCommandsSpeechEmph 
            Caption         =   "Emphasis Tags"
         End
         Begin VB.Menu mnuCommandsSpeechPause 
            Caption         =   "Pause Tags"
         End
         Begin VB.Menu mnuCommandsSpeechSpeed 
            Caption         =   "Speed Tags"
         End
         Begin VB.Menu mnuCommandsSpeechPitch 
            Caption         =   "Pitch Tags"
         End
         Begin VB.Menu mnuCommandsSpeechVolume 
            Caption         =   "Volume Tags"
         End
      End
      Begin VB.Menu mnuCommandsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommandsBold 
         Caption         =   "Bold Text"
      End
      Begin VB.Menu mnuCommandsItalic 
         Caption         =   "Italic Text"
      End
      Begin VB.Menu mnuCommandsUnderline 
         Caption         =   "Underlined Text"
      End
      Begin VB.Menu mnuCommandsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommandsBugAlert 
         Caption         =   "Bug Alert"
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
   Begin VB.Menu mnuInfoPopup 
      Caption         =   "Info Popup"
      Begin VB.Menu mnuInfoPopupTop 
         Caption         =   "Always On Top"
      End
      Begin VB.Menu mnuInfoPopupLock 
         Caption         =   "Lock Position"
      End
      Begin VB.Menu mnuInfoPopupDock 
         Caption         =   "Enable Docking"
      End
      Begin VB.Menu mnuInfoPopupSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfoPopupClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmSystray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public BalloonClickShow As Boolean

Private Declare Function apiShell_NotifyIconA Lib "shell32.dll" Alias "Shell_NotifyIconA" _
   (ByVal dwMessage As Long, lpData As NOTIFYICONDATAA) As Long
   
Private Declare Function apiShell_NotifyIconW Lib "shell32.dll" Alias "Shell_NotifyIconW" _
   (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Long


Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NOTIFYICON_VERSION = 3

Private Type NOTIFYICONDATAA
   cbSize As Long             ' 4
   hWnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip As String * 128      ' 152
   dwState As Long            ' 156
   dwStateMask As Long        ' 160
   szInfo As String * 256     ' 416
   uTimeOutOrVersion As Long  ' 420
   szInfoTitle As String * 64 ' 484
   dwInfoFlags As Long        ' 488
   guidItem As Long           ' 492
End Type
Private Type NOTIFYICONDATAW
   cbSize As Long             ' 4
   hWnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip(0 To 255) As Byte    ' 280
   dwState As Long            ' 284
   dwStateMask As Long        ' 288
   szInfo(0 To 511) As Byte   ' 800
   uTimeOutOrVersion As Long  ' 804
   szInfoTitle(0 To 127) As Byte ' 932
   dwInfoFlags As Long        ' 936
   guidItem As Long           ' 940
End Type


Private nfIconDataA As NOTIFYICONDATAA
Private nfIconDataW As NOTIFYICONDATAW

Private Const NOTIFYICONDATAA_V1_SIZE_A = 88
Private Const NOTIFYICONDATAA_V1_SIZE_U = 152
Private Const NOTIFYICONDATAA_V2_SIZE_A = 488
Private Const NOTIFYICONDATAA_V2_SIZE_U = 936

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

'Private Const WM_USER = &H400

Private Const NIN_SELECT = WM_USER
Private Const NINF_KEY = &H1
Private Const NIN_KEYSELECT = (NIN_SELECT Or NINF_KEY)
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

'Public Event SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
'Public Event SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
'Public Event SysTrayMouseMove()
'Public Event SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
'Public Event MenuClick(ByVal lIndex As Long, ByVal sKey As String)
'Public Event BalloonShow()
'Public Event BalloonHide()
'Public Event BalloonTimeOut()
'Public Event BalloonClicked()

Public Enum EBalloonIconTypes
   NIIF_NONE = 0
   NIIF_INFO = 1
   NIIF_WARNING = 2
   NIIF_ERROR = 3
   NIIF_NOSOUND = &H10
End Enum

Private m_bAddedMenuItem As Boolean
Private m_iDefaultIndex As Long

Private m_bUseUnicode As Boolean
Private m_bSupportsNewVersion As Boolean

Private Function Shell_NotifyIconA(dwMessage As Long, lpData As NOTIFYICONDATAA) As Long
Dim lRet As Long

lRet = apiShell_NotifyIconA(dwMessage, lpData)

If lRet = 0 And dwMessage <> NIM_SETVERSION Then
    'fail, re-add to tray
    Call AddToTray
    
    lRet = apiShell_NotifyIconA(dwMessage, lpData)
End If


Shell_NotifyIconA = lRet

End Function
Private Function Shell_NotifyIconW(dwMessage As Long, lpData As NOTIFYICONDATAW) As Long
Dim lRet As Long

lRet = apiShell_NotifyIconW(dwMessage, lpData)

If lRet = 0 And dwMessage <> NIM_SETVERSION Then
    'fail, re-add to tray
    Call AddToTray
    
    lRet = apiShell_NotifyIconW(dwMessage, lpData)
End If

Shell_NotifyIconW = lRet

End Function

Public Property Get ToolTip() As String
Dim sTip As String
Dim iPos As Long

If m_bUseUnicode Then
    sTip = nfIconDataW.szTip
Else
    sTip = nfIconDataA.szTip
End If


iPos = InStr(sTip, vbNullChar)
If (iPos > 0) Then
    sTip = Left$(sTip, iPos - 1)
End If

ToolTip = sTip

End Property

Public Property Let ToolTip(ByVal sTip As String)
   If (m_bUseUnicode) Then
      stringToArray sTip, nfIconDataW.szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
      nfIconDataW.uFlags = NIF_TIP
      Shell_NotifyIconW NIM_MODIFY, nfIconDataW
   Else
      If (sTip & Chr$(0) <> nfIconDataA.szTip) Then
         nfIconDataA.szTip = sTip & Chr$(0)
         nfIconDataA.uFlags = NIF_TIP
         Shell_NotifyIconA NIM_MODIFY, nfIconDataA
      End If
   End If
End Property

Public Property Get IconHandle() As Long
    If m_bUseUnicode Then
        IconHandle = nfIconDataW.hIcon
    Else
        IconHandle = nfIconDataA.hIcon
    End If
End Property

Public Property Let IconHandle(ByVal hIcon As Long)
   If (m_bUseUnicode) Then
      If (hIcon <> nfIconDataW.hIcon) Then
         nfIconDataW.hIcon = hIcon
         nfIconDataW.uFlags = NIF_ICON
         Shell_NotifyIconW NIM_MODIFY, nfIconDataW
      End If
   Else
      If (hIcon <> nfIconDataA.hIcon) Then
         nfIconDataA.hIcon = hIcon
         nfIconDataA.uFlags = NIF_ICON
         Shell_NotifyIconA NIM_MODIFY, nfIconDataA
      End If
   End If
End Property

Private Sub stringToArray( _
      ByVal sString As String, _
      bArray() As Byte, _
      ByVal lMaxSize As Long _
   )

Dim b() As Byte
Dim i As Long
Dim j As Long
   If LenB(sString) Then
      b = sString
      For i = LBound(b) To UBound(b)
         bArray(i) = b(i)
         If (i = (lMaxSize - 2)) Then
            Exit For
         End If
      Next i
      For j = i To lMaxSize - 1
         bArray(j) = 0
      Next j
   End If
End Sub
Private Function unicodeSize(ByVal lSize As Long) As Long
   If (m_bUseUnicode) Then
      unicodeSize = lSize * 2
   Else
      unicodeSize = lSize
   End If
End Function

Private Property Get nfStructureSize() As Long
   If (m_bSupportsNewVersion) Then
      If (m_bUseUnicode) Then
         nfStructureSize = NOTIFYICONDATAA_V2_SIZE_U
      Else
         nfStructureSize = NOTIFYICONDATAA_V2_SIZE_A
      End If
   Else
      If (m_bUseUnicode) Then
         nfStructureSize = NOTIFYICONDATAA_V1_SIZE_U
      Else
         nfStructureSize = NOTIFYICONDATAA_V1_SIZE_A
      End If
   End If
End Property

Public Sub ShowBalloonTip( _
      ByVal sMessage As String, _
      Optional ByVal sTitle As String = "Communicator", _
      Optional ByVal eIcon As EBalloonIconTypes, _
      Optional ByVal lTimeOutMs As Long = 5000, _
      Optional ByVal ForceIt As Boolean = False)

If ForceIt = False Then
    'If frmMain.mnuOptionsBalloonMessages.Checked = False Then Exit Sub
    If frmMain.mnuFileGameMode.Checked Then Exit Sub
End If

'ALWAYS prevent
If modVars.StealthMode Then Exit Sub


If modSpeech.sBalloon Then
    modSpeech.Say sMessage
End If

If modAlert.bBalloonTips Then
    pShowBalloonTip sMessage, sTitle, eIcon, lTimeOutMs
Else
    modAlert.ShowAlert sTitle, sMessage
End If

End Sub

Private Sub pShowBalloonTip(ByVal sMessage As String, _
      Optional ByVal sTitle As String = "Communicator", _
      Optional ByVal eIcon As EBalloonIconTypes, _
      Optional ByVal lTimeOutMs As Long = 5000)

Dim lR As Long

If (m_bSupportsNewVersion) Then
   If (m_bUseUnicode) Then
      stringToArray sMessage, nfIconDataW.szInfo, 512
      stringToArray sTitle, nfIconDataW.szInfoTitle, 128
      
      nfIconDataW.uTimeOutOrVersion = lTimeOutMs
      nfIconDataW.dwInfoFlags = eIcon
      nfIconDataW.uFlags = NIF_INFO
      
      lR = Shell_NotifyIconW(NIM_MODIFY, nfIconDataW)
   Else
      nfIconDataA.szInfo = sMessage
      nfIconDataA.szInfoTitle = sTitle
      nfIconDataA.uTimeOutOrVersion = lTimeOutMs
      nfIconDataA.dwInfoFlags = eIcon
      nfIconDataA.uFlags = NIF_INFO
      
      lR = Shell_NotifyIconA(NIM_MODIFY, nfIconDataA)
   End If
'Else
   'can't do it, fail silently.
End If

End Sub

Public Sub HideBalloon()

If m_bUseUnicode Then
    
    Erase nfIconDataW.szInfo
    Erase nfIconDataW.szInfoTitle
    
    nfIconDataW.uTimeOutOrVersion = 0
    nfIconDataW.dwInfoFlags = 0
    nfIconDataW.uFlags = NIF_INFO
    
    
    Shell_NotifyIconW NIM_MODIFY, nfIconDataW
Else
    With nfIconDataA
      .hWnd = Me.hWnd
      .uID = Me.Icon
      .uFlags = NIF_TIP Or NIF_INFO
      
      .szTip = vbNullChar
      .szInfo = vbNullChar
      .szInfoTitle = vbNullChar
      
      .uTimeOutOrVersion = NOTIFYICON_VERSION
   End With
   
   Shell_NotifyIconA NIM_MODIFY, nfIconDataA

End If

End Sub

'--------------------------------------------------------------------------------------------------------------------------

Public Function ShowMenu()


'SetForegroundWindow frmMain.hWnd

SetFocusToTray

mnuPopupHide.Visible = frmMain.Visible
mnuPopupShow.Visible = Not frmMain.Visible

Me.PopupMenu mnuPopup, , , , IIf(frmMain.Visible, mnuPopupHide, mnuPopupShow)

End Function

Private Sub cmdExit_Click()
modLoadProgram.ExitProgram
End Sub

Private Sub cmdShow_Click()
Const kMsg1 As String = "For another Communicator, go to File > New Communicator"
Const kMsg2 As String = "For another Communicator," & vbNewLine & "go to File > New Communicator"


Call TestInTray


'If frmMain.mnuOptionsWindow2BalloonInstance.Checked Then
    ShowBalloonTip "Communicator is already running down here" & vbNewLine & vbNewLine & kMsg2, , NIIF_INFO, , True
'Else
    'If frmMain.Visible = False Then
        'frmMain.ShowForm
    'Else
        'Call FlashWin
    'End If
'End If

AddText kMsg1, , True

End Sub

Private Sub Form_Load()
Me.Caption = modLoadProgram.Systray_Caption
Call AddToTray
End Sub

Public Sub RefreshTray()

'If (m_bUseUnicode) Then
'    nfIconDataW.cbSize = Len(nfIconDataW)
'    Shell_NotifyIconW NIM_MODIFY, nfIconDataW
'Else
'    Shell_NotifyIconA NIM_MODIFY, nfIconDataA
'End If

Call RemoveFromTray
Pause 100
Call AddToTray
frmMain.RefreshIcon

End Sub

Private Function TestInTray() As Boolean
Dim lR As Long

If m_bUseUnicode Then
    lR = Shell_NotifyIconW(NIM_MODIFY, nfIconDataW)
Else
    lR = Shell_NotifyIconA(NIM_MODIFY, nfIconDataA)
End If

TestInTray = CBool(lR)

End Function

Public Function SetFocusToTray() As Boolean

If m_bUseUnicode Then
    SetFocusToTray = CBool(Shell_NotifyIconW(NIM_SETFOCUS, nfIconDataW))
Else
    SetFocusToTray = CBool(Shell_NotifyIconA(NIM_SETFOCUS, nfIconDataA))
End If

End Function

Private Sub AddToTray()

If modVars.StealthMode = False Then

   ' Get version:
   Dim lMajor As Long
   Dim lMinor As Long
   Dim bIsNt As Boolean
   GetWindowsVersion lMajor, lMinor, , , bIsNt

   If (bIsNt) Then
      m_bUseUnicode = True
      If (lMajor >= 5) Then
         ' 2000 or XP
         m_bSupportsNewVersion = True
      End If
   ElseIf (lMajor = 4) And (lMinor = 90) Then
      ' Windows ME
      m_bSupportsNewVersion = True
   End If
   
   
   'Add the icon to the system tray...
   Dim lR As Long
   
   If (m_bUseUnicode) Then
      With nfIconDataW
         .hWnd = Me.hWnd
         .uID = frmSystray.Icon.Handle
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = frmSystray.Icon.Handle
         
         stringToArray modVars.GetTrayText(False), .szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
         
         If (m_bSupportsNewVersion) Then
            .uTimeOutOrVersion = NOTIFYICON_VERSION
         End If
         
         .cbSize = nfStructureSize
         
      End With
      
      lR = Shell_NotifyIconW(NIM_ADD, nfIconDataW)
      If (m_bSupportsNewVersion) Then
         Shell_NotifyIconW NIM_SETVERSION, nfIconDataW
      End If
   Else
      With nfIconDataA
         .hWnd = Me.hWnd
         .uID = Me.Icon
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = Me.Icon.Handle
         .szTip = modVars.GetTrayText(False) & Chr$(0)
         
         If (m_bSupportsNewVersion) Then
            .uTimeOutOrVersion = NOTIFYICON_VERSION
         End If
         
         .cbSize = nfStructureSize
      End With
      
      lR = Shell_NotifyIconA(NIM_ADD, nfIconDataA)
      
      If (m_bSupportsNewVersion) Then
         lR = Shell_NotifyIconA(NIM_SETVERSION, nfIconDataA)
      End If
   End If
   
   'frmMain.mnuOptionsSystray.Checked = True
   InTray = True
Else
    'frmMain.mnuOptionsSystray.Checked = False
   InTray = False
End If

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lX As Long
Dim bCan As Boolean
' VB manipulates the x value according to scale mode:
' we must remove this before we can interpret the
' message windows was trying to send to us:
    

lX = ScaleX(X, Me.ScaleMode, vbPixels)

If bModalFormShown Or modLoadProgram.bLoading Then
    Select Case lX
        Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK, WM_RBUTTONDOWN, NIN_BALLOONUSERCLICK
            Beep
    End Select
    
    Exit Sub
End If


Select Case lX
    Case WM_LBUTTONDOWN
        Call SysTrayMouseDown(vbLeftButton)
    Case WM_LBUTTONDBLCLK
        Call SysTrayDoubleClick(vbLeftButton)
    Case WM_RBUTTONDOWN
        Call SysTrayMouseDown(vbRightButton)
    Case NIN_BALLOONUSERCLICK
        Call BalloonClicked
        
'        Case WM_MOUSEMOVE
'            Call SysTrayMouseMove
'        Case WM_LBUTTONUP
'        Case WM_RBUTTONUP
'        Case WM_RBUTTONDBLCLK
'        Case NIN_BALLOONSHOW
'        Case NIN_BALLOONHIDE
'        Case NIN_BALLOONTIMEOUT
End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call RemoveFromTray
End Sub

Public Sub RemoveFromTray()
If (m_bUseUnicode) Then
    Shell_NotifyIconW NIM_DELETE, nfIconDataW
Else
    Shell_NotifyIconA NIM_DELETE, nfIconDataA
End If

InTray = False
'On Error Resume Next
'frmMain.mnuOptionsSystray.Checked = False

End Sub

Private Sub SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
On Error GoTo errhandler

If eButton <> vbLeftButton Then Exit Sub

Call SystrayShowForm

Exit Sub
errhandler:
'MsgBox "Error: " & Err.Description, vbOKOnly + vbExclamation, "Error"
End Sub

Private Sub SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
If (eButton = vbRightButton) Then
    ShowMenu
ElseIf eButton = vbLeftButton Then
    If frmMain.mnuOptionsWindow2SingleClick.Checked Then
        Call SystrayShowForm
    ElseIf frmMain.Visible Then
        frmMain.ZOrder vbBringToFront
    End If
End If
End Sub

Private Sub SystrayShowForm()

If frmMain.Visible Then
    Call frmMain.ShowForm(False)
Else
    Call frmMain.ShowForm
    SetForegroundWindow frmMain.hWnd
End If

End Sub

Public Sub BalloonClicked()

On Error Resume Next
If BalloonClickShow Then
    If frmMain.Visible = False And Not bStealth Then
        frmMain.ShowForm
    End If
    BalloonClickShow = False
End If

SetForegroundWindow frmMain.hWnd

End Sub

Private Sub mnuAFK_Click()
frmMain.mnuStatusAway_Click
End Sub

Private Sub mnuFontCopy_Click()
frmMain.mnuFontCopy_Click
End Sub

Private Sub mnuPopupCloseAndHost_Click()
mnuPopupCloseC_Click
mnuPopupHost_Click
End Sub

Private Sub mnuPopupCloseC_Click()
frmMain.cmdClose_Click
End Sub

Private Sub mnuPopupExit_Click()
ExitProgram
End Sub

Private Sub mnuPopupFolder_Click()
frmMain.mnuFileOpenFolder_Click
End Sub

Private Sub mnuPopupGameMode_Click()
'frmMain.mnufileGameMode.Checked = Not frmMain.mnufileGameMode.Checked
frmMain.mnuFileGameMode_Click
End Sub

Private Sub mnuPopupHide_Click()
frmMain.ShowForm False
End Sub

Private Sub mnuPopupHost_Click()
'mnuPopupCloseC_Click
frmMain.cmdListen_Click
End Sub

Private Sub mnuPopupNew_Click()
frmMain.mnuFileNew_Click
End Sub

Private Sub mnuPopupScan_Click()

Call CheckMainVisible

frmMain.cmdScan_Click

frmUDP.cmdScan_Click

End Sub

Private Sub mnuPopupShow_Click()
frmMain.ShowForm
End Sub

Private Sub mnuPopupSingleClick_Click()
mnuPopupSingleClick.Checked = Not mnuPopupSingleClick.Checked
frmMain.mnuOptionsWindow2SingleClick.Checked = mnuPopupSingleClick.Checked
End Sub

'Private Sub mnuPopupAnim_Click()
'mnuPopupAnim.Checked = Not mnuPopupAnim.Checked
''tmrMain.Enabled = (Status = Connected) And mnuPopupAnim.Checked
'End Sub

Private Sub mnuPopupUpdates_Click()

Call CheckMainVisible

frmMain.CheckForUpdates

End Sub

Private Sub mnuPopupWhatsnew_Click()

Call CheckMainVisible

Load frmHelp
frmHelp.cmdChange_Click

'frmHelp.Show vbModeless, frmMain 'show it so we can setfocus

frmHelp.Show vbModal, frmMain

End Sub

Private Sub CheckMainVisible()

If frmMain.Visible = False Then
    On Error Resume Next
    frmMain.ShowForm
    frmMain.Refresh
    
    Pause 5
    
End If

End Sub

'######################################################################
'######################################################################
'menu stuff
'######################################################################
'######################################################################

Private Sub mnuSBCopyrIP_Click()
frmMain.mnuSBCopyrIP_Click
End Sub
Private Sub mnuSBCopylIP_Click()
frmMain.mnuSBCopylIP_Click
End Sub
Private Sub mnuSBObtain_Click()
frmMain.mnuSBObtain_Click
End Sub
Private Sub mnuSBObtainLocal_Click()
frmMain.mnuSBObtainLocal_Click
End Sub
Private Sub mnuFontColour_Click()
frmMain.mnuFontColour_Click
End Sub
Private Sub mnuFontDialog_Click()
frmMain.mnuFontDialog_Click
End Sub

Private Sub mnuStatusAway_Click()
frmMain.mnuStatusAway_Click
End Sub
Private Sub mnuStatusResetName_Click()
frmMain.mnuStatusResetName_Click
End Sub
Private Sub mnuCommandsMe_Click()
frmMain.mnuCommandsMe_Click
End Sub
Private Sub mnuCommandsDescribe_Click()
frmMain.mnuCommandsDescribe_Click
End Sub
Private Sub mnuCommandsSpeechEmph_Click()
frmMain.mnuCommandsSpeechEmph_Click
End Sub
Private Sub mnuCommandsSpeechPause_Click()
frmMain.mnuCommandsSpeechPause_Click
End Sub
Private Sub mnuCommandsSpeechPitch_Click()
frmMain.mnuCommandsSpeechPitch_Click
End Sub
Private Sub mnuCommandsSpeechSpeed_Click()
frmMain.mnuCommandsSpeechSpeed_Click
End Sub
Private Sub mnuCommandsSpeechActorsHL_Click()
frmMain.mnuCommandsSpeechActorsHL_Click
End Sub
Private Sub mnuCommandsSpeechActorsJW_Click()
frmMain.mnuCommandsSpeechActorsJW_Click
End Sub
Private Sub mnuCommandsStopSpeech_Click()
frmMain.mnuCommandsStopSpeech_Click
End Sub
Private Sub mnuCommandsTestSpeech_Click()
frmMain.mnuCommandsTestSpeech_Click
End Sub
Private Sub mnuCommandsSpeechVolume_Click()
frmMain.mnuCommandsSpeechVolume_Click
End Sub
'Private Sub mnuCommandsJeffery_Click()
'frmMain.mnuCommandsJeffery_Click
'End Sub
'Private Sub mnuCommandsGregory_Click()
'frmMain.mnuCommandsGregory_Click
'End Sub
Private Sub mnuDPView_Click()
frmMain.mnuDPView_Click
End Sub
Private Sub mnuDPOpen_Click()
frmMain.mnuDPOpen_Click
End Sub
Private Sub mnuCommandsBugAlert_Click()
frmMain.mnuCommandsBugAlert_Click
End Sub
Private Sub mnuCommandsBold_Click()
frmMain.mnuCommandsBold_Click
End Sub
Private Sub mnuCommandsUnderline_Click()
frmMain.mnuCommandsUnderline_Click
End Sub
Private Sub mnuCommandsItalic_Click()
frmMain.mnuCommandsItalic_Click
End Sub
