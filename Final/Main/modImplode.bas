Attribute VB_Name = "modImplode"
Option Explicit

Public fmX As Long, fmY As Long

Private Const AnimTime = 400

Private Const IDANI_CAPTION = &H3

Public Enum eAnimType
    aImplode = 0
    'aSlide = 1
    aFade = 2
    aRandom = 3
    None = 4
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'redraw
Private Declare Function InvalidateRect Lib "user32" ( _
    ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As Any) As Long
'                                                             POINTAPI

'end implode
'begin animate

Private Declare Function pAnimateWindow Lib "user32" Alias "AnimateWindow" (ByVal hWnd As Long, _
    ByVal dwTime As Long, ByVal dwFlags As Long) As Long

Private Const AW_HOR_POSITIVE = &H1 ' Animate the window from
'left to right. This flag can be used with roll or slide
'animation It is ignored when used with the AW_CENTER flag.
Private Const AW_HOR_NEGATIVE = &H2 ' Animate the window from
'right to left. This flag can be used with roll or slide
'animation. It is ignored when used with the AW_CENTER flag.
Private Const AW_VER_POSITIVE = &H4 ' Animate the window from
'top to bottom. This flag can be used with roll or slide
'animation. It is ignored when used with the AW_CENTER flag.
Private Const AW_VER_NEGATIVE = &H8 ' Animate the window from
'bottom to top. This flag can be used with roll or slide
'animation. It is ignored when used with the AW_CENTER flag.
Private Const AW_CENTER = &H10 ' Makes the window appear to
'collapse inward if the AW_HIDE flag is used or expand outward
'if the AW_HIDE flag is not used.
Private Const AW_HIDE = &H10000 ' Hides the window. By default,
'the window is shown.
Private Const AW_ACTIVATE = &H20000 ' Activates the window. Do
'not use this flag with AW_HIDE.
Private Const AW_SLIDE = &H40000 ' Uses slide animation. By
'default, roll animation is used. This flag is ignored when used
'with the AW_CENTER flag.
Private Const AW_BLEND = &H80000 ' Uses a fade effect. This flag
'can be used only if hwnd is a top-level window.

Public Enum eAnimateWindow
    HorizontalP = AW_HOR_POSITIVE
    HorizontalN = AW_HOR_NEGATIVE
    VerticalP = AW_VER_POSITIVE
    VerticalN = AW_VER_NEGATIVE
    Centre = AW_CENTER
    Hide = AW_HIDE
    Activate = AW_ACTIVATE
    Slide = AW_SLIDE
    Blend = AW_BLEND
End Enum

'end animate

''for fading
'Private Const GWL_EXSTYLE = (-20)
'Private Const WS_EX_LAYERED = &H80000
'Private Const LWA_ALPHA = &H2
'
'Private Declare Function GetWindowLong Lib "user32" _
'  Alias "GetWindowLongA" (ByVal hWnd As Long, _
'  ByVal nIndex As Long) As Long
'
'Private Declare Function SetWindowLong Lib "user32" _
'   Alias "SetWindowLongA" (ByVal hWnd As Long, _
'   ByVal nIndex As Long, ByVal dwNewLong As Long) _
'   As Long
'
'Private Declare Function SetLayeredWindowAttributes Lib _
'    "user32" (ByVal hWnd As Long, ByVal crKey As Long, _
'    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'
'
'Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, _
'    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'
'Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
'
'Private TranS As Byte
'Private TimerID As Long
'Private phWnd As Long
''Private Const Freq = 4 '400/4 = every 100 ms
'end fading

'Private Function DrawAnimatedRects(ByVal hWnd As Long, ByVal idAni As Long, _
'    lprcFrom As RECT, lprcTo As RECT) As Long
'
'DrawAnimatedRects = apiDrawAnimatedRects(hWnd, idAni, lprcFrom, lprcTo)
'
'End Function

Private Sub ImplodeForm(hWnd As Long, Optional Reverse As Boolean = False, Optional IsFormCentered As Boolean = False)
On Error Resume Next
Dim f As RECT, i As RECT
GetWindowRect hWnd, f
If IsFormCentered = True Then CenterRect f
i.Left = f.Left + (f.Right - f.Left) / 2
i.Right = i.Left
i.Top = f.Top + (f.Bottom - f.Top) / 2
i.Bottom = i.Top

If Not Reverse Then
    DrawAnimatedRects hWnd, IDANI_CAPTION, f, i
Else
    DrawAnimatedRects hWnd, IDANI_CAPTION, i, f
End If
End Sub

Private Sub ImplodeFormToMouse(hWnd As Long, Optional Reverse As Boolean = False, Optional IsFormCentered As Boolean = False)

On Error Resume Next
Dim f As RECT, i As RECT, P As POINTAPI
GetWindowRect hWnd, f
If IsFormCentered Then CenterRect f
GetCursorPos P
i.Left = P.X
i.Right = P.X
i.Top = P.Y
i.Bottom = P.Y
If Not Reverse Then
    DrawAnimatedRects hWnd, IDANI_CAPTION, f, i
Else
    DrawAnimatedRects hWnd, IDANI_CAPTION, i, f
End If
End Sub

Public Sub ImplodeFormToTray(hWnd As Long, Optional Reverse As Boolean = False, Optional IsFormCentered As Boolean = False)
On Error Resume Next                  'ByRef Rec As RECT,
Dim f As RECT, i As RECT, P As POINTAPI
GetWindowRect hWnd, f
If IsFormCentered = True Then CenterRect f
GetWindowRect GetTrayhWnd, i
i.Left = i.Left + ((i.Right - i.Left) / 2)
i.Right = i.Left
If Not Reverse Then
    DrawAnimatedRects hWnd, IDANI_CAPTION, f, i
Else
    DrawAnimatedRects hWnd, IDANI_CAPTION, i, f
    'Rec = f
End If
End Sub

Public Sub MoveForm(hWnd As Long, rTo As RECT, rFrom As RECT, Optional ByVal Anim As Boolean = True)

'On Error Resume Next

If Anim Then
    DrawAnimatedRects hWnd, IDANI_CAPTION, rFrom, rTo
End If

SetWindowPos hWnd, 0, rTo.Left, rTo.Top, rTo.Right, rTo.Bottom, 0
'DrawAnimatedRects hWnd, IDANI_CAPTION, rTo, rFrom

End Sub

Private Function GetTrayhWnd() As Long
On Error Resume Next
Dim OurParent As Long
Dim OurHandle As Long
OurParent = FindWindow("Shell_TrayWnd", "")
OurHandle = FindWindowEx(OurParent&, 0, "TrayNotifyWnd", vbNullString)
GetTrayhWnd = OurHandle
End Function

Private Sub CenterRect(ByRef r As RECT)
On Error Resume Next
Dim H As Long, w As Long, tbh As Long, sw As Long, sh As Long
H = r.Bottom - r.Top
w = r.Right - r.Left
tbh = GetTaskbarHeight / Screen.TwipsPerPixelY
sw = Screen.width / Screen.TwipsPerPixelX
sh = (Screen.height / Screen.TwipsPerPixelY) - tbh
r.Left = (sw / 2) - (w / 2)
r.Right = (sw / 2) + (w / 2)
r.Top = (sh / 2) - (H / 2)
r.Bottom = (sh / 2) + (H / 2)
End Sub

'#####################################################################################################################
'#####################################################################################################################

''animate window
'Private Sub SlideWindow(ByRef hWnd As Long, ByVal ShowIt As Boolean, ByVal msTime As Long)
'Dim A As Long
'
'A = GetRandomSlideAnim(ShowIt)
'
'If ShowIt Then
'    ShowWindow hWnd, SW_HIDE
'Else
'    ShowWindow hWnd, SW_SHOWNORMAL
'    'A = A Or Hide
'End If
'
'pAnimateWindow hWnd, msTime, A
'
'If ShowIt Then
'    ShowWindow hWnd, SW_HIDE
'    'ShowWindow hWnd, SW_SHOWNORMAL
'End If
'
'End Sub


Public Sub AnimateAWindow(ByVal hWnd As Long, ByVal AnimType As eAnimType, _
    Optional ByVal Reverse As Boolean = False, Optional ByVal ImplodeToMouse As Boolean = True, _
    Optional ByVal ForceAnim As Boolean = False)

Dim r As Single

If frmMain.mnuOptionsWindow2Animation.Checked = False Then Exit Sub

If AnimType = aRandom Then
    r = Rnd()
    
    If r > 0.5 Then
        AnimType = aImplode
    Else
        AnimType = aFade
    End If
    
End If


Select Case AnimType
    Case eAnimType.aImplode
        
        If ImplodeToMouse Then
            ImplodeFormToMouse hWnd, Not Reverse
        Else
           ImplodeForm hWnd, Not Reverse
        End If
        
    Case Else 'eAnimType.aFade
        
        FadeWindowOut hWnd, 200
        
    'Case Else 'eAnimType.aSlide
        
        'If Reverse Then
            'FadeWindow hWnd, Not Reverse, AnimTime
        'Else
        'SlideWindow hWnd, Not Reverse, AnimTime
        'End If
        
End Select

End Sub


'Private Function GetRandomSlideAnim(ByVal ShowIt As Boolean) As eAnimateWindow
'
'Dim TmpH As Byte, TmpV As Byte
'Dim TmpC As Byte, TmpCB As Byte
'Dim RandomAnim As eAnimateWindow
'
''either 0 or 1
'TmpH = Rnd()
'TmpV = Rnd()
'TmpC = Rnd()
'TmpCB = Rnd()
'
'If TmpH Then
'    RandomAnim = HorizontalP
'Else
'    RandomAnim = HorizontalN
'End If
'
'If TmpV Then
'    RandomAnim = RandomAnim Or VerticalP
'Else
'    RandomAnim = RandomAnim Or VerticalN
'End If
'
'If TmpC And ShowIt Then
'    If TmpCB Then
'        RandomAnim = Blend
'    Else
'        RandomAnim = RandomAnim Or Centre
'    End If
'ElseIf ShowIt = False Then
'    RandomAnim = RandomAnim Or Blend
'End If
'
'If ShowIt Then
'    RandomAnim = RandomAnim Or Activate
'Else
'    RandomAnim = RandomAnim Or Hide 'And (Not Activate)
'End If
'
'GetRandomSlideAnim = RandomAnim
'
'End Function

'Private Function GetRandomFadeAnim(ByVal ShowIt As Boolean) As eAnimateWindow
'
'Dim RandomAnim As eAnimateWindow
'
'RandomAnim = Blend
'
'If ShowIt Then
'    RandomAnim = RandomAnim Or Activate
'Else
'    RandomAnim = RandomAnim Or Hide
'End If
'
'GetRandomFadeAnim = RandomAnim
'
'End Function

'---------------------------------------
Public Function FadeWindowOut(ByVal hWnd As Long, msTime As Long) As Long

FadeWindowOut = pAnimateWindow(hWnd, msTime, AW_HIDE Or AW_BLEND)

End Function

'Private Sub FadeWindow(ByRef hWnd As Long, ByVal ShowIt As Boolean, ByVal msTime As Long)
'
'Dim l As Long
''A = GetRandomFadeAnim(ShowIt)
'
'
'If ShowIt Then
'    TranslucentForm hWnd, 0
'End If
'ShowWindow hWnd, SW_SHOWNORMAL
'
''Else
'    'ShowWindow hWnd, SW_SHOWNORMAL
'    'A = A Or Hide
''End If
'
''pAnimateWindow hWnd, msTime, A
'
'phWnd = hWnd
'
'If ShowIt Then
'    TranS = 0
'    TimerID = SetTimer(0, 0, msTime / 40, AddressOf fadeInTimer)
'Else
'    TranS = 255
'    TimerID = SetTimer(0, 0, msTime / 40, AddressOf fadeOutTimer)
'End If
'
'Do
'    'Pause 5
'    'if pause (and therefore doevents) is used, the form may be unloaded, so sleep must be used
'    Sleep 5
'    DoEvents 'still need to 'listen' for timer calls
'Loop Until TimerID = 0
'
'ShowWindow hWnd, SW_HIDE
'phWnd = 0
'TranS = 0
'
''remove the layered bit (slows down resizing etc)
'l = GetWindowLong(hWnd, GWL_EXSTYLE)
'If (l And WS_EX_LAYERED) = WS_EX_LAYERED Then
'    'remove the layerd bit
'    l = l And (Not WS_EX_LAYERED)
'    SetWindowLong hWnd, GWL_EXSTYLE, l 'keep previous stuff
'End If
'
'
'If ShowIt = False Then
'    TranslucentForm hWnd, 255
'    'ShowWindow hWnd, SW_SHOWNORMAL
'End If
'
'End Sub
'
'Private Function TranslucentForm(hWnd As Long, TranslucenceLevel As Byte) As Boolean
'
''0 = completely transparent, 255 = completely opaque
'Dim l As Long
'
'l = GetWindowLong(hWnd, GWL_EXSTYLE)
'
'If (l And WS_EX_LAYERED) = 0 Then
'    'add the layerd bit
'    SetWindowLong hWnd, GWL_EXSTYLE, WS_EX_LAYERED Or l 'keep previous stuff
'End If
'
'SetLayeredWindowAttributes hWnd, 0, TranslucenceLevel, LWA_ALPHA
'
'TranslucentForm = (Err.LastDllError = 0)
'
'End Function
'
'Private Sub fadeInTimer(hWnd As Long, Msg As Long, idTimer As Long, dwTime As Long)
'
'If TranS + 10 > 255 Then
'    TranS = 255
'    'kill timer
'    KillTimer 0, TimerID
'    TimerID = 0
'Else
'    TranS = TranS + 10
'End If
'
'TranslucentForm phWnd, TranS
'
'End Sub
'
'Private Sub fadeOutTimer(hWnd As Long, Msg As Long, idTimer As Long, dwTime As Long)
'
'If TranS - 10 < 0 Then
'    TranS = 0
'    'kill timer
'    KillTimer 0, TimerID
'    TimerID = 0
'Else
'    TranS = TranS - 10
'End If
'
'TranslucentForm phWnd, TranS
'
'End Sub
