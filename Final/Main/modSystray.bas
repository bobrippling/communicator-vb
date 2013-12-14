Attribute VB_Name = "modSystray"
Option Explicit

'declare external procedures
Private Declare Function Shell_NotifyIcon _
    Lib "shell32.dll" _
    Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, _
    lpData As NotifyIconData) As Long
                                    
'Private Declare Function SetForegroundWindow Lib "user32" _
                                    (ByVal hwnd As Long) As Long
                                                                                        
'...constants
Public Const WM_MOUSEMOVE = &H200

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208

Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

'declare enums
Public Enum eSystrayIconModifier
    modifyIcon = &H2
    modifyTip = &H4
End Enum

'declare types
'Private Type NOTIFYICONDATA
    'cbSize                          As Long
    'hWnd                            As Long
    'uID                             As Long
    'uFlags                          As Long
    'uCallbackMessage                As Long
    'hIcon                           As Long
    'szTip                           As String * 64
'End Type

Private Type NotifyIconData
   cbSize As Long             ' 4
   hWnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip As String    ' 280
   dwState As Long            ' 284
   dwStateMask As Long        ' 288
   szInfo As String   ' 800
   uTimeOutOrVersion As Long  ' 804
   szInfoTitle As String ' 932
   dwInfoFlags As Long        ' 936
   guidItem As Long           ' 940
End Type

Public Enum EBalloonIconTypes
   NIIF_NONE = 0
   NIIF_INFO = 1
   NIIF_WARNING = 2
   NIIF_ERROR = 3
   NIIF_NOSOUND = &H10
End Enum

'declare variables
Private SystrayIcon                 As NotifyIconData
Private InTray                      As Boolean

Public Property Get InSystray() As Boolean
InSystray = InTray
End Property


Public Sub AddSystrayIcon(ByVal TrayTipText As String, Icon As Long, _
ByVal hWnd As Long)

With SystrayIcon
    .cbSize = Len(SystrayIcon)
    .hIcon = Icon
    .hWnd = hWnd
    .szTip = TrayTipText & Chr$(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON + NIF_MESSAGE + NIF_TIP
    .uID = vbNull
End With

Call Shell_NotifyIcon(NIM_ADD, SystrayIcon)

frmMain.mnuOptionsSystray.Checked = True

InTray = True

End Sub

Public Sub ModifySystrayIcon(ModifyType As eSystrayIconModifier, _
    ByVal NewValue As Variant)

Select Case ModifyType
    Case modifyIcon
        SystrayIcon.hIcon = CLng(NewValue)
    Case modifyTip
        SystrayIcon.szTip = CStr(NewValue) & Chr$(0)
End Select

Call Shell_NotifyIcon(NIM_MODIFY, SystrayIcon)
frmMain.mnuOptionsSystray.Checked = True

InTray = True

End Sub

Public Sub RemoveSystrayIcon()
Call Shell_NotifyIcon(NIM_DELETE, SystrayIcon)
frmMain.mnuOptionsSystray.Checked = False
End Sub

Public Sub ShowBalloonTip(ByVal sMessage As String, _
      Optional ByVal sTitle As String, _
      Optional ByVal eIcon As EBalloonIconTypes, _
      Optional ByVal lTimeOutMs As Long = 10000)

Dim lR As Long

SystrayIcon.szInfo = StrConv(sMessage, vbUnicode)
SystrayIcon.szInfoTitle = StrConv(sTitle, vbUnicode)
SystrayIcon.uTimeOutOrVersion = lTimeOutMs
SystrayIcon.dwInfoFlags = eIcon
SystrayIcon.uFlags = EBalloonIconTypes.NIIF_INFO

lR = Shell_NotifyIcon(NIM_MODIFY, SystrayIcon)

End Sub

