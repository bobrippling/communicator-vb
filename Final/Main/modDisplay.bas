Attribute VB_Name = "modDisplay"
Option Explicit

'mirror
Private Const WS_EX_LAYOUTRTL = &H400000
Private Const WS_EX_NOINHERITLAYOUT = &H100000

'################################################
'Private Declare Function Button_SetElevationRequiredState Lib "comctl32" ( _
    ByVal hWnd As Long, ByVal fRequired As Long) As Long

Private Const BCM_SETSHIELD = &H160C&

'################################################
'dwm stuff

Public Const Glass_Border_Indent = 30

'aero glass etc
Private Type ptMargins
    cLeft As Long
    cRight As Long
    cTop As Long
    cBottom As Long
End Type
Private Type ptDWM_BlurBehind
    dwFlags As Long
    fEnable As Long 'bool
    hRgnBlur As Long 'hRGN
    fTransitionOnMaximized As Long 'bool
End Type

Private Declare Sub DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, _
    ByRef pMargins As ptMargins)

Private Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef pfEnabled As Long) As Long

Private Declare Function DwmEnableBlurBehindWindow Lib "dwmapi.dll" (ByVal hWnd As Long, _
    pBlurBehind As ptDWM_BlurBehind) As Long

'Indicates a value for fEnable has been specified.
Private Const DWM_BB_ENABLE = &H1

'show text/banner in an empty textbox
'Private Const EM_SETCUEBANNER = &H1501
Private Const ECM_FIRST = &H1500                'Edit control messages
Private Const EM_SETCUEBANNER = (ECM_FIRST + 1)
'Private Const EM_GETCUEBANNER = (ECM_FIRST + 2) 'Set the cue banner with the lParm = LPCWSTR

'cmd stuff
Private Const BM_SETIMAGE = &HF7
Private Const BS_COMMANDLINK = &HE
Private Const BCM_SETNOTE = &H1609



Private Type EDITBALLOONTIP
    cbStruct As Long
    pszTitle As Long
    pszText As Long
    ttiIcon As Long
End Type

Private Const EM_SHOWBALLOONTIP = (ECM_FIRST + 3)
Private Const EM_HIDEBALLOONTIP = (ECM_FIRST + 4)

Public Enum BalloonTipIconConstants
   TTI_NONE = 0
   TTI_INFO = 1
   TTI_WARNING = 2
   TTI_ERROR = 3
End Enum

'#########################################
'Thumbnailage

Private Const DWM_TNP_RECTDESTINATION = &H1
'Indicates a value for rcDestination has been specified.
Private Const DWM_TNP_RECTSOURCE = &H2
'Indicates a value for rcSource has been specified.
Private Const DWM_TNP_OPACITY = &H4
'Indicates a value for opacity has been specified.
Private Const DWM_TNP_VISIBLE = &H8
'Indicates a value for fVisible has been specified.
Private Const DWM_TNP_SOURCECLIENTAREAONLY = &H10
'Indicates a value for fSourceClientAreaOnly has been specified.

Private Const S_OK = 0

Private Type DWM_THUMBNAIL_PROPERTIES
    dwFlags As Long
    rcDestination As RECT
    rcSource As RECT
    opacity As Byte
    fVisible As Long 'bool
    fSourceClientAreaOnly As Long 'bool
End Type

Private Declare Function DwmRegisterThumbnail Lib "dwmapi.dll" (ByVal hWndDestination As Long, _
    ByVal hWndSource As Long, ByVal phThumbnailId As Long) As Long

Private Declare Function DwmUnregisterThumbnail Lib "dwmapi.dll" (ByVal hThumbnailId As Long) As Long

Private Declare Function DwmUpdateThumbnailProperties Lib "dwmapi.dll" (ByVal hThumbnailId As Long, _
    ByRef ptnProperties As DWM_THUMBNAIL_PROPERTIES) As Long

'Private Declare Function DwmQueryThumbnailSourceSize Lib "dwmapi.dll" (ByVal hThumbnail As Long, ByVal pSize As Long) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long

'#####################################################################
'transparency

Private Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hWnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


'#####################################################################
'WINDOW FLASHING

'window flashing
Private Const FLASHW_STOP = 0 'Stop flashing. The system restores the window to its original state.
Private Const FLASHW_CAPTION = &H1 'Flash the window caption.
Private Const FLASHW_TRAY = &H2 'Flash the taskbar button.
Private Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY) 'Flash both the window caption and taskbar button. This is equivalent to setting the FLASHW_CAPTION Or FLASHW_TRAY flags.
Private Const FLASHW_TIMER = &H4 'Flash continuously, until the FLASHW_STOP flag is set.
Private Const FLASHW_TIMERNOFG = &HC 'Flash continuously until the window comes to the foreground.


Private Type FLASHWINFO
    cbSize As Long
    hWnd As Long
    dwFlags As Long
    uCount As Long
    dwTimeout As Long
End Type

'Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Private Declare Function FlashWindowEx Lib "user32" (pFWI As FLASHWINFO) As Long
Public bResetFlash As Boolean

'#####################################################################
'mouse leave tracking
Private Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
'Private Const TME_CANCEL = &H80000000
'Private Const TME_HOVER = &H1&
Private Const TME_LEAVE = &H2&
'Private Const TME_NONCLIENT = &H10&
'Private Const TME_QUERY = &H40000000

'#####################################################################

'Private Declare Function GetWindowRect Lib "user32" ( _
'    ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function GetClientRect Lib "user32" ( _
'    ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function InvalidateRect Lib "user32" ( _
'    ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'Private Declare Function ScreenToClient Lib "user32" ( _
'    ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'
'Public Sub RepaintWindow( _
'        ByRef objThis As Object, _
'        Optional ByVal bClientAreaOnly As Boolean = True)
'
'Dim tR As RECT
'Dim tP As POINTAPI
'
'If bClientAreaOnly Then
'    GetClientRect objThis.hWnd, tR
'Else
'    GetWindowRect objThis.hWnd, tR
'    tP.X = tR.Left: tP.Y = tR.Top
'    ScreenToClient objThis.hWnd, tP
'    tR.Left = tP.X: tR.Top = tP.Y
'    tP.X = tR.Right: tP.Y = tR.Bottom
'    ScreenToClient objThis.hWnd, tP
'    tR.Right = tP.X: tR.Bottom = tP.Y
'End If
'
'InvalidateRect objThis.hWnd, tR, 1
'
'End Sub

Public Function CanShow_XPButtons() As Boolean

'XP buttons are the angled ones
CanShow_XPButtons = (Not modLoadProgram.bVistaOrW7) And _
                    (Not modLoadProgram.bSafeMode) And _
                    (modLoadProgram.bAllowXPButtons) 'And _
                    (modDisplay.VisualStyle)

End Function

'#####################################################################

Public Function EnableMouseTracking(hWnd As Long) As Long
Dim ET As TRACKMOUSEEVENTTYPE

'initialize structure
With ET
    .cbSize = Len(ET)
    .hwndTrack = hWnd
    .dwFlags = TME_LEAVE
End With

'start the tracking
EnableMouseTracking = TrackMouseEvent(ET)

End Function


'#####################################################################

Public Sub FlashWin(Optional ByVal hWnd As Long = -1)

Dim FlashInfo As FLASHWINFO
Dim lRet As Long

If hWnd = -1 Then
    hWnd = frmMain.hWnd
End If

If modVars.IshWndForegroundWindow(hWnd) = False Then
    
    With FlashInfo
        .cbSize = Len(FlashInfo)
        .dwFlags = FLASHW_ALL
        .dwTimeout = 0 'system default
        .hWnd = hWnd
        .uCount = 1
    End With
    
    'lRet = FlashWindowEx(FlashInfo)
    FlashWindowEx FlashInfo
End If

'Pause 1
'
'If lRet <> 0 Then
'    lRet = FlashWindowEx(FlashInfo)
''Else
'    'yay, tis now flashing
'End If

End Sub

Public Function CompositionEnabled() As Boolean
Dim lComp As Long, bComp As Boolean

On Error GoTo EH
If DwmIsCompositionEnabled(lComp) = 0 Then
    bComp = CBool(lComp)
Else
    bComp = False
End If

CompositionEnabled = bComp

Exit Function
EH:
CompositionEnabled = False
End Function

'###############################################
'Backcolour must be black for it to look correct
'###############################################

Public Sub SetGlassBorders(hWnd As Long, Optional Left As Long = Glass_Border_Indent, _
    Optional Right As Long = 0, Optional Top As Long = 0, _
    Optional Bottom As Long = 0)

Dim Margin As ptMargins

Margin.cBottom = Bottom
Margin.cLeft = Left
Margin.cRight = Right
Margin.cTop = Top

On Error Resume Next
DwmExtendFrameIntoClientArea hWnd, Margin

End Sub

Public Sub RemoveGlass(hWnd As Long)
Dim Margin As ptMargins

On Error Resume Next
DwmExtendFrameIntoClientArea hWnd, Margin
End Sub

Public Sub SetButtonIcon(ByVal hWnd As Long, ByVal hIcon As Long)
'Set the icon of a button - make sure your button's big enough for the icon
'e.g.
'SetButtonIcon cmdTest.hWnd, frmMain.Icon.Handle
'will set cmdTest's button icon to frmMain's icon

SendMessageByLong hWnd, BM_SETIMAGE, 1&, hIcon

End Sub
Public Function ShowButtonShieldIcon(ByVal hWnd As Long, Optional ByVal bShowShield As Boolean = True) As Boolean

'If bSendMessage = False Then
    'On Error GoTo EH
    'ShowButtonShieldIcon = (Button_SetElevationRequiredState(hWnd, Abs(bShowShield)) = 1)
'Else
ShowButtonShieldIcon = (SendMessageByLong(hWnd, BCM_SETSHIELD, 0, Abs(bShowShield)) = 1)
'End If

'Exit Function
'EH:
'MsgBox "Error: " & Err.Description, vbExclamation, "Error"
End Function


Public Sub SetTextBoxBanner(hWnd As Long, ByVal sTxt As String)

'SendMessageByString hWnd, EM_SETCUEBANNER, 0&, StrConv(sTxt, vbUnicode)

SendMessageByLong hWnd, EM_SETCUEBANNER, 0, StrPtr(sTxt)

End Sub

'######################################################################################################
'manual banner

'Private Sub SetupCueControl(ctl As Control, sCue As String)
'
''assign the cue text to the control's edit
''box, as well as to the tag property. Using
''the tag property to store the cue prompt text
''negates the requirement to maintain the text
''in a separate array. If your application design
''uses are using the tag property for another
''purpose, such as to store the dirty text of
''the control, then a string array must be maintained
''along with a mechanism to identify the control
''in order to assign the correct prompt to the
''respective control.
'
'With ctl
'    .ForeColor = vbButtonShadow
'    'tag is set first to ensure
'    'CheckCuePromptChange sets correct value
'    'when the control's Change event fires
'    .Tag = sCue
'    .Text = sCue
'End With
'
'End Sub
'
'Private Sub CheckCuePromptChange(ctl As Control)
'ctl.HelpContextID = (Trim$(ctl.Text) = ctl.Tag)
'End Sub
'
'Private Sub CheckCuePromptOnFocus(ctl As Control)
'With ctl
'    If .HelpContextID = True Then
'        .Text = ""
'        .ForeColor = vbWindowText
'    Else
'        .Selstart = 0
'        .Sellength = Len(.Text)
'        .HelpContextID = False
'    End If
'End With
'End Sub
'
'Private Sub CheckCuePromptBlur(ctl As Control)
'With ctl
'   If Len(Trim$(.Text)) = 0 Then
'      .Text = .Tag
'      .ForeColor = vbButtonShadow
'      .HelpContextID = True
'   End If
'End With
'End Sub

'######################################################################################################
'balloon stuff

Public Sub ShowBalloonTip(ByRef oTB As TextBox, ByVal sBalloonTipTitle As String, _
    ByVal sBalloonTipText As String, Optional eBalloonTipIcon As BalloonTipIconConstants = TTI_INFO)

Dim tEBT As EDITBALLOONTIP

If oTB.Enabled Then
    With tEBT
        .cbStruct = LenB(tEBT)
        
        .pszText = StrPtr(sBalloonTipText)
        .pszTitle = StrPtr(sBalloonTipTitle)
        
        .ttiIcon = eBalloonTipIcon
    End With
    
    SendMessageByAny oTB.hWnd, EM_SHOWBALLOONTIP, 0, tEBT
    'used to be SMW (SendMessageWateva)
End If

End Sub

Public Sub HideBalloonTip(ByVal TB_hWnd As Long)
SendMessageByLong TB_hWnd, EM_HIDEBALLOONTIP, 0, 0
End Sub

'######################################################################################################

Public Function GetManifestString() As String

Dim Tmp As String

Tmp = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
Tmp = Tmp & vbNewLine & "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">"
Tmp = Tmp & vbNewLine & "<assemblyIdentity"
Tmp = Tmp & vbNewLine & "   version=""1.0.0.0"""
Tmp = Tmp & vbNewLine & "   processorArchitecture=""*""" '* was X86
Tmp = Tmp & vbNewLine & "   name=""" & App.EXEName & """"
Tmp = Tmp & vbNewLine & "   type=""win32"""
Tmp = Tmp & vbNewLine & "/>"
Tmp = Tmp & vbNewLine & "<description>" & App.FileDescription & "</description>"
Tmp = Tmp & vbNewLine & "<dependency>"
Tmp = Tmp & vbNewLine & "   <dependentAssembly>"
Tmp = Tmp & vbNewLine & "     <assemblyIdentity"
Tmp = Tmp & vbNewLine & "       type=""win32"""
Tmp = Tmp & vbNewLine & "       name=""Microsoft.Windows.Common-Controls""" 'ensure name = LCase
Tmp = Tmp & vbNewLine & "       version=""6.0.0.0"""
Tmp = Tmp & vbNewLine & "       processorArchitecture=""*""" '* was X86
Tmp = Tmp & vbNewLine & "       publicKeyToken=""6595b64144ccf1df"""
Tmp = Tmp & vbNewLine & "       language=""*"""
Tmp = Tmp & vbNewLine & "     />"
Tmp = Tmp & vbNewLine & "   </dependentAssembly>"
Tmp = Tmp & vbNewLine & "</dependency>"
Tmp = Tmp & vbNewLine & "</assembly>"

GetManifestString = Tmp

End Function

Public Function setVisualStyle(ByVal TurnOn As Boolean) As Boolean
Dim FName As String, Path As String
Dim iFile As Integer
Dim bSuccess As Boolean

FName = App.EXEName & ".exe.manifest"
    
Path = AppPath() & FName

On Error Resume Next

If TurnOn Then
    iFile = FreeFile()
    Open Path For Output As #iFile
        Print #iFile, GetManifestString;
    Close #iFile
    SetAttr Path, vbHidden
    
    bSuccess = FileExists(Path, vbHidden)
Else
    SetAttr Path, vbNormal
    If FileExists(Path) Then
        Kill Path
    End If
    
    bSuccess = Not FileExists(Path, vbHidden)
End If

setVisualStyle = bSuccess

End Function

Public Property Get VisualStyle() As Boolean
Dim FName As String, Path As String
Dim bExists As Boolean

FName = App.EXEName & ".exe.manifest"
Path = AppPath() & FName

bExists = FileExists(Path, vbHidden)
If bExists = False Then
    bExists = FileExists(Path)
End If

VisualStyle = bExists

End Property

'###################################################################################

Public Sub SetTransparentStyle(hWnd As Long, Optional bAdd As Boolean = True)
Dim l As Long

l = GetWindowLong(hWnd, GWL_EXSTYLE)

If bAdd Then
    If (l And WS_EX_LAYERED) = 0 Then
        'add the layerd bit
        SetWindowLong hWnd, GWL_EXSTYLE, WS_EX_LAYERED Or l
    End If
ElseIf (l And WS_EX_LAYERED) = WS_EX_LAYERED Then
    SetWindowLong hWnd, GWL_EXSTYLE, l And (Not WS_EX_LAYERED)
End If

End Sub

Public Sub SetTransparency(hWnd As Long, btLevel As Byte)
SetLayeredWindowAttributes hWnd, 0, btLevel, LWA_ALPHA
End Sub

'###################################################################################

Public Sub Mirror(ByRef Frm As Form, Optional ByVal bDoControls As Boolean = True, _
    Optional ByVal bMirror As Boolean = True)

Dim Ctrl As Control


MirrorhWnd Frm.hWnd, bMirror

If bDoControls Then
    For Each Ctrl In Frm.Controls
        On Error Resume Next
        MirrorhWnd Ctrl.hWnd, bMirror
    Next Ctrl
End If


End Sub

Public Sub MirrorhWnd(hWnd As Long, Optional ByVal bMirror As Boolean = True)

If bMirror Then
    SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYOUTRTL Or WS_EX_NOINHERITLAYOUT
Else
    SetWindowLong hWnd, GWL_EXSTYLE, (GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYOUTRTL) And Not WS_EX_NOINHERITLAYOUT
End If

End Sub

'###################################################################################

Public Function RegisterThumbnail(ByRef hThumb As Long, _
    ByVal hWnd_Source As Long, ByVal hWnd_Dest As Long) As Boolean

Dim l As Long

On Error Resume Next
l = DwmRegisterThumbnail(hWnd_Dest, hWnd_Source, VarPtr(hThumb))

RegisterThumbnail = (Err.Number = 0) And (l = S_OK)

End Function

Public Function UnRegisterThumbnail(hThumb As Long) As Boolean
Dim l As Long

On Error Resume Next
l = DwmUnregisterThumbnail(hThumb)

UnRegisterThumbnail = (Err.Number = 0) And (l = S_OK)
hThumb = 0

End Function

Public Function SetThumbNailProps(hThumb As Long, btOpacity As Byte, bVisible As Boolean, bClientOnly As Boolean, _
    sourceRect As RECT, destRect As RECT, _
    Optional bSetSrcRect As Boolean = True, Optional bSetDestRect As Boolean = True) As Boolean

Dim lR As Long
Dim dskThumbProps As DWM_THUMBNAIL_PROPERTIES


With dskThumbProps
    .dwFlags = DWM_TNP_VISIBLE Or DWM_TNP_SOURCECLIENTAREAONLY Or DWM_TNP_OPACITY
    
    If bSetSrcRect Then
        .dwFlags = .dwFlags Or DWM_TNP_RECTSOURCE
        .rcSource = sourceRect
    End If
    
    If bSetDestRect Then
        .dwFlags = .dwFlags Or DWM_TNP_RECTDESTINATION
        .rcDestination = destRect
    End If
    
    
    .fVisible = Abs(bVisible)
    .fSourceClientAreaOnly = Abs(bClientOnly)
    .opacity = btOpacity
End With

On Error Resume Next
lR = DwmUpdateThumbnailProperties(hThumb, dskThumbProps)

SetThumbNailProps = (Err.Number = 0) And (lR = S_OK)

End Function
Public Function SetThumbNailRect(hThumb As Long, rcRect As RECT, bSetSrcRect As Boolean) As Boolean

Dim lR As Long
Dim dskThumbProps As DWM_THUMBNAIL_PROPERTIES


With dskThumbProps
    If bSetSrcRect Then
        .dwFlags = DWM_TNP_RECTSOURCE
        .rcSource = rcRect
    Else
        .dwFlags = DWM_TNP_RECTDESTINATION
        .rcDestination = rcRect
    End If
End With

On Error Resume Next
lR = DwmUpdateThumbnailProperties(hThumb, dskThumbProps)

SetThumbNailRect = (Err.Number = 0) And (lR = S_OK)

End Function

