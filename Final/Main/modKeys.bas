Attribute VB_Name = "modKeys"
Option Explicit

' Declare Type for API call:
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

' API declarations:

Private Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Sub keybd_event Lib "user32" _
    (ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function GetKeyboardState Lib "user32" _
    (pbKeyState As Byte) As Long

Private Declare Function SetKeyboardState Lib "user32" _
    (lppbKeyState As Byte) As Long

' Constant declarations:
'Private Const VK_NUMLOCK = &H90
'Private Const VK_SCROLL = &H91
Private Const VK_CAPITAL = &H14
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub SetCaps(ByVal TurnOn As Boolean)

Dim VerI As OSVERSIONINFO
Dim NumLockState As Boolean
Dim ScrollLockState As Boolean
Dim CapsLockState As Boolean
Dim keys(0 To 255) As Byte

GetKeyboardState keys(0)
'CapsLock handling:
CapsLockState = keys(VK_CAPITAL)

If CapsLockState <> TurnOn Then 'toggle it
    VerI.dwOSVersionInfoSize = Len(VerI)
    GetVersionEx VerI
    
    If VerI.dwPlatformId = VER_PLATFORM_WIN32_NT Then '=== WinNT
        'Simulate Key Press
        keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        
        'Simulate Key Release
        keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        
    ElseIf VerI.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then '=== Win95/98
        
        keys(VK_CAPITAL) = 1
        SetKeyboardState keys(0)
        
    End If
End If

''NumLock handling:
'NumLockState = keys(VK_NUMLOCK)
'If NumLockState <> True Then 'Turn numlock on
'    If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then '=== Win95/98
'
'        keys(VK_NUMLOCK) = 1
'        SetKeyboardState keys(0)
'    ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then '=== WinNT
'        'Simulate Key Press
'        keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
'        'Simulate Key Release
'        keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY _
'            Or KEYEVENTF_KEYUP, 0
'    End If
'End If
'
''ScrollLock handling:
'ScrollLockState = keys(VK_SCROLL)
'If ScrollLockState <> True Then 'Turn Scroll lock on
'    If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then '=== Win95/98
'        keys(VK_SCROLL) = 1
'        SetKeyboardState keys(0)
'    ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then '=== WinNT
'        'Simulate Key Press
'        keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
'        'Simulate Key Release
'        keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY _
'            Or KEYEVENTF_KEYUP, 0
'    End If
'End If

End Sub

Public Function Caps() As Boolean
Dim state As Integer
state = GetKeyState(vbKeyCapital)
Caps = (state = 1 Or state = -127)
End Function
