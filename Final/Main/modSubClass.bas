Attribute VB_Name = "modSubClass"
Option Explicit


#Const HOOK_TXTOUT = False



Public Const WndProcStr = "WNDPROC"
Private Const ObjPtrStr = "OBJPTR"
Public Declare Function GetProp Lib "user32" Alias "GetPropA" ( _
    ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" ( _
    ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" ( _
    ByVal hWnd As Long, ByVal lpString As String) As Long


'##########################################################################################
'subclassing code
Private OldWndProc As Long

Public Const GWL_WNDPROC = (-4)
'// Windows API Call for catching messages

Private PbSubClassing As Boolean

'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, ByVal nIndex As Long, ByRef dwNewLong As Long) As Long


Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

'for tab intercepting
Private Const WM_ACTIVATE As Long = &H6
Private Const WA_INACTIVE As Long = 0
Private Const HC_ACTION   As Long = 0
Private Const VK_TAB      As Long = &H9
Private Const WH_KEYBOARD As Long = 2
Private Const KF_UP       As Long = &H8000

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
                        ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, _
                        ByVal dwThreadId As Long) As Long

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
                        ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                        ByVal hWnd As Long, ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long


Private hKybdHook As Long 'Handle to our hook procedure
Private lPrevProc  As Long 'Original windows procedure address
'##########################################################################################


'taskbar destroyed/created
Private WM_TASKBARCREATED As Long

Private Declare Function RegisterWindowMessage Lib "user32" _
    Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
'end

'###################################################
'###################################################
'stick game

'mouse wheel
Private Const WM_MOUSEWHEEL = &H20A
Private Const WHEEL_DELTA = 120
Private OldStickWndProc As Long


Private Const WM_NCRBUTTONDOWN = &HA4
'###################################################
'###################################################

'mini window
Private OldMiniWndProc As Long
'focus stuff
    Public Const WM_MOUSEMOVE = &H200&
    Public Const WM_MOUSELEAVE = &H2A3&
    
    Public Const WM_KILLFOCUS = &H6&
    Public Const WM_MOVING = &H216&
'end mini

    
'menu code-----------------------------------------------------------------------------
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" ( _
    ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, _
    ByVal lpNewItem As String) As Long
    
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, _
                                                ByVal bRevert As Long) As Long
                                                
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, _
                    ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" ( _
    ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, _
    ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long 'c bool

'Public Declare Function DeleteMenu Lib "user32" ( _
    ByVal hMenu As Long, ByVal iditem As Long, ByVal wFlags As Long) As Long


Private Const WM_SYSCOMMAND = &H112
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Public Const MF_BYPOSITION = &H400&

'Public Const MF_REMOVE = &H1000&
''Private Const SC_CLOSE = &HF060
'Public Const MF_BYCOMMAND = &H0
''Private Const MF_GRAYED = &H1
''Public Const SC_MOVE = &HF010
'Public Const SC_MAXIMIZE = &HF030
''Public Const SC_MINIMIZE = &HF020

Private Const IDM_ExitProgram As Long = 1010
'end menu code------------------------------------------------------------------------------

'for resizing-----------------------------------
'Private Const SC_SIZE = &HF000&

Private minX As Long
Private minY As Long
Private maxX As Long
Private maxY As Long

Public Const WM_GETMINMAXINFO As Long = &H24

'Private Type POINTAPI
'    X As Long
'    Y As Long
'End Type

Public Type MINMAXINFO
    ptReserved As PointAPI
    ptMaxSize As PointAPI
    ptMaxPosition As PointAPI
    ptMinTrackSize As PointAPI
    ptMaxTrackSize As PointAPI
End Type

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, _
    hpvSource As Any, ByVal cbCopy As Long)
    
'end resizing-----------------------------------

'rtf urling
Public lpfnOldWinProc          As Long
Public m_lRTBhWnd              As Long

Private Type CharRange
   cpMin As Long
   cpMax As Long
End Type
Private Type NMHDR
   hWndFrom As Long
   idfrom As Long
   Code As Long
End Type
Private Type ENLINK
   NMHDR As NMHDR
   Msg As Long
   wParam As Long
   lParam As Long
   chrg As CharRange
End Type

Private Const WM_DESTROY                As Long = &H2
Private Const WM_NOTIFY                 As Long = &H4E& '78
Private Const EN_LINK                   As Long = &H70B&

'Private Const SW_SHOWNORMAL = 1
'Private Const SW_NORMAL = 1

'Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal _
    lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Function GetWindowTextLength Lib "user32" Alias _
    "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)

'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal _
        wParam As Long, ByVal lParam As Long) As Long

'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

'menuing
Private Const WM_MENUSELECT = &H11F
Private Const MF_SYSMENU = &H2000&
Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" ( _
    ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

'#####################################################
'suspend code
Private Const WM_POWERBROADCAST As Long = &H218

Private Const PBT_APMSUSPEND As Long = &H4
Private Const PBT_APMQUERYSUSPEND As Long = &H0

Private Const BROADCAST_QUERY_DENY As Long = &H424D5144

''vista
'Private Declare Function RegisterPowerSettingNotification Lib "user32" (hRecipient As Long, _
'    PowerSettingGuid As LPCGUID, Flags As Long) As Long
'
'Private Const DEVICE_NOTIFY_WINDOW_HANDLE = 0& 'For flags
'Private Const GUID_SYSTEM_AWAYMODE As String = "{98a7f580-01f7-48aa-9c0f-44352c29e5C0}"                'this one

'end suspend code
'#####################################################


'focus code
'Private Const WM_ACTIVATE       As Long = &H6
Private Const WM_ACTIVATEAPP   As Long = &H1C
'Private Const WA_INACTIVE      As Long = 0
Private Const WA_ACTIVE        As Long = 1
'Private Const WA_CLICKACTIVE   As Long = 2

'Private Const WM_ACTIVATE As Long = &H6
'Private Const HC_ACTION   As Long = 0
'Private Const VK_TAB      As Long = &H9
'Private Const WH_KEYBOARD As Long = 2
'Private Const KF_UP       As Long = &H8000
'end focus code

'##################################################################################
'combo box stuff
Private Const CB_SHOWDROPDOWN = &H14F

Private Const WM_COMMAND = &H111

Private Const CBN_CLOSEUP = &H8
Private Const CBN_DROPDOWN = &H7

'Private Const CBN_SETFOCUS = &H3
'Private Const CBN_SELENDCANCEL = &HA
'##################################################################################

''vista composition
'Private Const WM_DWMCOMPOSITIONCHANGED = &H31E

Public Property Get bSubClassing() As Boolean
bSubClassing = PbSubClassing
End Property

Private Function frmMainWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Dim iHi As Integer, iLo As Integer
Dim Txt As String

If hWnd = frmMain.hWnd Then
    
    Select Case uMsg
        Case WM_SYSCOMMAND
            
            If wParam = IDM_ExitProgram Then
                'MsgBox "VB Web Append to System Menu Example", vbInformation, "About"
                ExitProgram
                Exit Function
                
            'ElseIf (wParam And &HFFF0) = SC_SIZE Then
                'If frmMain.Height < 7950 Then
                    'frmMain.Height = 7950
                    'Exit Function
                'End If
            End If
            
            
            
        Case WM_GETMINMAXINFO
            
            Dim MMI As MINMAXINFO
            
            CopyMemory MMI, ByVal lParam, LenB(MMI)
            
            'set the MINMAXINFO data to the
            'minimum and maximum values set
            'by the option choice
            
            With MMI
                .ptMinTrackSize.X = minX
                .ptMinTrackSize.Y = minY
                '.ptMaxTrackSize.X = maxX
                '.ptMaxTrackSize.Y = maxY
            End With
            
            CopyMemory ByVal lParam, MMI, LenB(MMI)
            
            frmMainWindowProc = 0
            Exit Function
            
        Case WM_TASKBARCREATED
            
            frmSystray.RefreshTray 'add + remove
            
            
        Case WM_MENUSELECT
            
            CopyMemory iLo, wParam, 2
            CopyMemory iHi, ByVal VarPtr(wParam) + 2, 2
            
            If (iHi And MF_SYSMENU) = 0 Then 'not a top-level menu/sys menu
                
                Dim m As MENUITEMINFO, aCap As String
                
                m.dwTypeData = Space$(64)
                m.cbSize = Len(m)
                m.cch = 64
                m.fMask = MIIM_DATA Or MIIM_TYPE
                
                Txt = vbNullString
                
                If GetMenuItemInfo(lParam, CLng(iLo), False, m) Then
                    
                    aCap = m.dwTypeData & Chr$(0)
                    aCap = Left$(aCap, InStr(aCap, Chr$(0)) - 1)
                    
                    'Debug.Print aCap
                    
                    Select Case True
                        Case aCap Like "New Communicator*": Txt = vbNullString '<--bugged'Txt = "Open a new Communicator"
                        Case aCap Like "Game Mode*": Txt = "Game Mode - Suppress Alerts like Balloon Tips, etc"
                        Case aCap Like "Stealth Mode*": Txt = "Put Communicator into Stealth Mode - Notepad Masquerade"
                        'Case aCap Like "Window Animation": Txt = "Decide how to animate windows when they are shown/hidden"
                        Case aCap Like "Communicator Website*": Txt = "Open the web browser, and go to Communicator's site"
                        Case aCap Like "Login*": Txt = "Login to view/save game statistics"
                        Case aCap Like "HTTP Download": Txt = "Download Communicator's latest version"
                        Case aCap Like "Port Forwarding": Txt = "Forward ports, so people outside your router can connect"
                        Case aCap Like "File Transfer": Txt = "Upload a file to send to someone"
                        Case aCap Like "Clear Screen": Txt = "Clear the text box"
                        Case aCap Like "Save As...": Txt = "Save the conversation"
                        
                        Case aCap = "Inactive Timer"
                            If frmMain.ucInactiveTimer.Enabled Then
                                Txt = "Inactive Timer Interval: " & CStr(frmMain.ucInactiveTimer.InactiveInterval / 60000) & " minutes"
                            End If
                        
                        'Case aCap Like "*": Txt = ""
                        
                        Case Else: Txt = vbNullString
                        
                    End Select
                    
                End If
                
                frmMain.SetInfoPanel Txt
                
            End If
            
            
        Case WM_ACTIVATEAPP
            
            Select Case LoWord(wParam)
                Case WA_INACTIVE
                    frmMain.Form_apiLostFocus
                Case WA_ACTIVE
                    frmMain.Form_apiGotFocus
            End Select
            
        Case WM_POWERBROADCAST
            
            If wParam = PBT_APMQUERYSUSPEND Then
                
                Txt = Time$()
                
                AddConsoleText "Suspend request received from Windows at " & Txt, , , , True
                
                If frmMain.mnuOptionsAdvNoStandby.Checked Then
                    
                    frmMainWindowProc = BROADCAST_QUERY_DENY
                    AddConsoleText "Suspend request denied at " & Txt
                    Exit Function
                    
                ElseIf frmMain.mnuOptionsAdvNoStandbyConnected.Checked Then
                    If Status = Connected Then
                        frmMainWindowProc = BROADCAST_QUERY_DENY
                        AddConsoleText "Suspend request denied at " & Txt
                        Exit Function
                    End If
                End If
                
            End If
            
        'Case WM_DWMCOMPOSITIONCHANGED
            
            'I suspect the underlying vb class is placing a higher level
            'hook in the chain when it becomes the active window. So we
            'will trap the wm_active message and reset our hook back to the top.
'        Case WM_ACTIVATE
'            If LoWord(wParam) = WA_INACTIVE Then
'                'Remove the hook
'                UnhookWindowsHookEx hKybdHook
'                hKybdHook = 0
'            Else
'                'Set the hook in the hook chain
'                If hKybdHook = 0 Then
'                    hKybdHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, App.hInstance, 0&)
'                End If
'            End If
            
            
        Case WM_DESTROY
            'In case it wasn't already done, un-subclass the window
            Call CallWindowProc(OldWndProc, hWnd, uMsg, wParam, lParam)
            
            SubClass hWnd, False
            
    End Select
    
    
'ElseIf hWnd = frmMain.wbDP.hWnd Then
    
    'Debug.Print Hex$(uMsg)
    
End If

frmMainWindowProc = CallWindowProc(OldWndProc, hWnd, uMsg, wParam, lParam)

End Function

'#############################################################################

#If HOOK_TXTOUT Then
    Private Function KeyboardProc(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    If Code >= 0 Then
        If Code = HC_ACTION Then
            If frmMain.ActiveControl.hWnd = frmMain.txtOut.hWnd Then 'This will keep the tab order working for other controls
                If wParam = VK_TAB Then 'Only filter the tab key
                    If (lParam >= 0 And KF_UP) Then 'It gets ugly without this
                        frmMain.txtOut_KeyPress vbKeyTab
                        KeyboardProc = 1 'Do not call next hook in the hook chain
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    KeyboardProc = CallNextHookEx(hKybdHook, Code, wParam, ByVal lParam) 'We don't need this so lets see if any other hook in the chain do
    
    End Function
    
    Public Property Let txtOut_Hooked(bHook As Boolean)
    
    
    If hKybdHook Then
        UnhookWindowsHookEx hKybdHook
        hKybdHook = 0
    ElseIf bEnableKeyboardHook Then
        hKybdHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, App.hInstance, 0&)
    End If
    
    End Property
    Public Property Get txtOut_Hooked() As Boolean
    txtOut_Hooked = (hKybdHook <> 0)
    End Property
#End If

'#############################################################################

Public Sub SubClass(ByVal hWnd As Long, Optional ByVal DoIt As Boolean = True)

If DoIt Then
    OldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf frmMainWindowProc)
    
    AddToSysMenu hWnd
    
    #If HOOK_TXTOUT Then
        If frmMain.Visible Then 'get the keyboard hook set
            If hKybdHook = 0 Then
                If bEnableKeyboardHook Then
                    hKybdHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, App.hInstance, 0&)
                End If
            End If
        'Else
            'keyboard hook is set when shown
        End If
    #End If
Else
    Call SetWindowLong(hWnd, GWL_WNDPROC, OldWndProc)
    
    RemoveFromSysMenu hWnd
End If

PbSubClassing = DoIt

frmMain.mnuDevSubClass.Checked = PbSubClassing
'frmMain.sbMain.Panels(3).Visible = PbSubClassing

AddConsoleText "Subclassing: " & CStr(DoIt)

End Sub

Public Sub SubClass_Init()
WM_TASKBARCREATED = RegisterWindowMessage("TaskbarCreated")
End Sub

Private Sub AddToSysMenu(ByVal hWnd As Long)

Dim lhSysMenu As Long, lRet As Long, lCount As Long

lhSysMenu = GetSystemMenu(hWnd, 0&)

lCount = GetMenuItemCount(lhSysMenu)

''c bool
'lRet = ModifyMenu(lhSysMenu, lCount - 1, MF_STRING Or MF_BYPOSITION, 0&, "Send To Tray")

lRet = AppendMenu(lhSysMenu, MF_SEPARATOR, 0&, vbNullString)
lRet = AppendMenu(lhSysMenu, MF_STRING, IDM_ExitProgram, "Exit Program")


End Sub

Private Sub RemoveFromSysMenu(ByVal hWnd As Long)

Dim hSysMenu As Long, menuCount As Long


hSysMenu = GetSystemMenu(hWnd, 0&)
menuCount = GetMenuItemCount(hSysMenu)

RemoveMenu hSysMenu, menuCount - 1, MF_BYPOSITION

'Remove the seperator bar
RemoveMenu hSysMenu, menuCount - 2, MF_BYPOSITION

'ModifyMenu hSysMenu, menuCount - 3, MF_STRING Or MF_BYPOSITION, 0&, "Close"

End Sub


'-------------------------------
'resizing
'-------------------------------
Public Sub SetMinMaxInfo(ByVal lMinX As Long, ByVal lMinY As Long, _
                         ByVal lMaxX As Long, ByVal lMaxY As Long)

minX = lMinX
minY = lMinY
maxX = lMaxX
maxY = lMaxY

End Sub

'url rtf stuff
Public Function RtfWndProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

'http://www.vbmonster.com/Uwe/Forum.aspx/vb/25788/How-to-add-a-hyperlink-text-in-a-text-box
'RTFIN SUBCLASSING PROC

Dim udtENLINK           As ENLINK
Dim sUrl                As String
Dim lPos1               As Long
Dim lPos2               As Long
Dim sRTBText            As String
Dim lRTBTextLength      As Long

'Debug.Print uMsg

Select Case uMsg
    Case WM_DESTROY
        'In case it wasn't already done, un-subclass the window
        Call CallWindowProc(lpfnOldWinProc, hWnd, uMsg, wParam, lParam)
        
        'Call frmMain.rtfIn.DisableURLHook(hwnd)
        SetWindowLong hWnd, GWL_WNDPROC, lpfnOldWinProc

        lpfnOldWinProc = 0


    Case WM_NOTIFY

        'Now it gets a bit tricky.  lParam is a pointer to an ENLINK structure.
        'The pointer is the memory address of where the structure resides.
        'We need to fill a local variable for that structure from the pointer.
        CopyMemory udtENLINK, ByVal lParam, Len(udtENLINK)

        'Make sure the notification is from the RTB
        If udtENLINK.NMHDR.hWndFrom <> m_lRTBhWnd Then
            RtfWndProc = CallWindowProc(lpfnOldWinProc, hWnd, uMsg, wParam, lParam)
            Exit Function
        End If

        'Make sure this is the EN_LINK notification
        If udtENLINK.NMHDR.Code <> EN_LINK Then
            RtfWndProc = CallWindowProc(lpfnOldWinProc, hWnd, uMsg, wParam, lParam)
            Exit Function
        End If

        'Now see if this is a left mouse up message
        If udtENLINK.Msg = WM_LBUTTONUP Then
            'We get the first and last character position of the link From
            'the CHARRANGE structure.
            lPos1 = udtENLINK.chrg.cpMin
            lPos2 = udtENLINK.chrg.cpMax

            'Because we don't have a direct reference to the RichTextBox
            'control, we need to get its text via API functions

            'This function gives us the length of the text (number of characters)
            lRTBTextLength = GetWindowTextLength(m_lRTBhWnd)

            'Set up a buffer variable which will receive the text.
            'Buffer variables used in API functions almost always must
            'be pre-allocated to the size of the text that the buffer
            'will receive; otherwise, the function usually causes the
            'application to crash or only part of the text will be
            'retrieved.  Add 1 to accomodate a terminating null
            'character.
            sRTBText = String$(lRTBTextLength + 1, vbNullChar)
            
            
            'Get the text from the RTB
            Call GetWindowText(m_lRTBhWnd, sRTBText, lRTBTextLength + 1)

            'Extract the URL
            sUrl = Trim$(Mid$(StripNulls(sRTBText), lPos1 + 1, lPos2 - lPos1))
            
            'AddConsoleText "URL Clicked: " & sUrl
            
            'Launch the URL using whatever the default application is
            'Call ShellExecute(0, "open", sUrl, vbNullString, vbNullString, SW_SHOW)
            OpenURL sUrl

            'Return a non-zero since we processed the message and don't want
            'the control to process it.
            RtfWndProc = 1
        Else
            RtfWndProc = CallWindowProc(lpfnOldWinProc, hWnd, uMsg, wParam, lParam)
        End If

    Case Else
        RtfWndProc = CallWindowProc(lpfnOldWinProc, hWnd, uMsg, wParam, lParam)
        
End Select

End Function

Public Function StripNulls(OriginalStr As String) As String
'This removes the extra Nulls so String comparisons will work
If InStr(OriginalStr, Chr(0)) Then
    OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If

StripNulls = OriginalStr
End Function


Private Function StickWindowProc(ByVal hWnd As Long, ByVal iMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Dim Delta As Long
'Static Travel As Long

If modStickGame.StickFormLoaded Then
    If hWnd = frmStickGame.hWnd Then
        
        If iMsg = WM_MOUSEWHEEL Then
            'Delta = HiWord(wParam)
            
            'x = LoWord(lParam)
            'y = HiWord(lParam)
            
            'Travel = Travel Mod WHEEL_DELTA
            
            frmStickGame.Form_WheelScroll (wParam > 0)
            
        ElseIf iMsg = WM_NCRBUTTONDOWN Then
            StickWindowProc = 0 'prevent right click on title bar
            Exit Function
        End If
        
    End If
End If

StickWindowProc = CallWindowProc(OldStickWndProc, hWnd, iMsg, wParam, lParam)

End Function

Public Sub SubClassStick(ByVal hWnd As Long, Optional ByVal DoIt As Boolean = True)

If DoIt Then
    OldStickWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf StickWindowProc)
Else
    SetWindowLong hWnd, GWL_WNDPROC, OldStickWndProc
    OldStickWndProc = 0
End If

End Sub

Public Property Get bStickSubClassing() As Boolean
bStickSubClassing = CBool(OldStickWndProc)
End Property

Private Function HiWord(DWord As Long) As Integer
'top 2 bytes
CopyMemory HiWord, ByVal VarPtr(DWord) + 2&, 2&
End Function

Private Function LoWord(DWord As Long) As Integer
'bottom 2 bytes
CopyMemory LoWord, DWord, 2&
End Function

'Public Function HiWord(lDWord As Long) As Integer
'HiWord = (lDWord And &HFFFF0000) \ &H10000
'End Function

'Private Function LoWord(dw As Long) As Integer
'
'If dw And &H8000& Then
'    LoWord = &H8000& Or (dw And &H7FFF&)
'Else
'    LoWord = dw And &HFFFF&
'End If
'
'End Function

'##########################
Public Sub SubclassAuto(Frm As Object, Optional ByVal bDo As Boolean = True)
Dim hWnd As Long

hWnd = Frm.hWnd

If bDo Then
    SetProp hWnd, WndProcStr, SetWindowLong(hWnd, GWL_WNDPROC, AddressOf AutoWindowProc)
    SetProp hWnd, ObjPtrStr, ObjPtr(Frm)
Else
    SetWindowLong hWnd, GWL_WNDPROC, RemoveProp(hWnd, WndProcStr)
    RemoveProp hWnd, ObjPtrStr
End If
End Sub

Private Function AutoWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Dim Obj As Object


Set Obj = modVars.ObjFromPtr(GetProp(hWnd, ObjPtrStr))
AutoWindowProc = Obj.WindowProc(hWnd, uMsg, wParam, lParam)
Set Obj = Nothing

End Function


'Public Sub Subclass_frmManual(ByVal hwnd As Long, Optional ByVal bDo As Boolean = True)
'
'If bDo Then
'    SetProp hwnd, WndProcStr, SetWindowLong(hwnd, GWL_WNDPROC, AddressOf frmManual_WindowProc)
'Else
'    SetWindowLong hwnd, GWL_WNDPROC, RemoveProp(hwnd, WndProcStr)
'End If
'
'End Sub
'
'Private Function frmManual_WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, _
'    ByVal wParam As Long, ByVal lParam As Long) As Long
'
'Dim otherMsg As Long
'
'If uMsg = WM_COMMAND Then
'    otherMsg = HiWord(wParam)
'
'    If otherMsg = CBN_CLOSEUP Then
'        'Debug.Print "ClosedUp (Eaten)"
'
'        frmManual.cboIP_Click
'        frmManual_WindowProc = 1
'        Exit Function
'
''    ElseIf otherMsg = CBN_DROPDOWN Then
''        Debug.Print "DropDown (Eaten)"
''
''        frmManual.cboIP_Click
''        frmManual_WindowProc = 1
''        Exit Function
'
'    End If
'End If
'
'frmManual_WindowProc = CallWindowProc(GetProp(hwnd, WndProcStr), hwnd, uMsg, wParam, lParam)
'
'End Function


'##########################
'Public Sub SubclasstxtIP(ByVal hWnd As Long, Optional ByVal bDo As Boolean = True)
'Dim lRet As Long
'
'If bDo Then
'    lRet = SetProp(hWnd, WndProcStr, SetWindowLong(hWnd, GWL_WNDPROC, AddressOf txtIPWindowProc))
'Else
'    SetWindowLong hWnd, GWL_WNDPROC, RemoveProp(hWnd, WndProcStr)
'End If
'
'End Sub
'
'Private Function txtIPWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
'    ByVal wParam As Long, ByVal lParam As Long) As Long
'
''Select Case uMsg
'    'case x
''End Select
'Debug.Print "&H" & Hex$(uMsg)
'
'txtIPWindowProc = CallWindowProc(GetProp(hWnd, WndProcStr), hWnd, uMsg, wParam, lParam)
'End Function


