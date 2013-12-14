Attribute VB_Name = "modMsgboxEx"
Option Explicit

'MsgBoxEx for VB
'Variable position custom MsgBox by Ray Mercer
'Copyright (C) 1999 by Ray Mercer - All rights reserved
'Based on a sample I posted to news://msnews.microsoft.com/microsoft.public.vb.general.discussion
'Based on an earlier post by Didier Lefebvre <didier.lefebvre@free.fr> in the same newsgroup
'Latest version available at www.shrinkwrapvb.com
'
'You are free to use this code in your own projects and modify it in any way you see fit
'however you may not redistribute this archive sample without the express written consent
'from the author - Ray Mercer <raymer@shrinkwrapvb.com>
'
'*******************
'HOW TO USE
'*******************
'Just pop this module in your VB5 or 6 project.  Then you can call MsgBoxEx instead of MsgBox
'MsgBoxEx will return the same vbMsgBoxResults as MsgBox, but adds the frm, Left, and Top parameters.
'
' Useage sample:
'
'Dim ret As VbMsgBoxResult
'ret = MsgBoxEx(Me, "This is a test", vbOKCancel, "Cool!", 10, 10)
'If ret = vbOK Then
'    MsgBox "User pressed OK!"
'End If
'
' *Note if you leave out the Left and Top parameters the MsgBox will center itself over the Form
'
'e.g.;
'Call MsgBoxEx(Me, "This is a test")
'
'This will center the msgBox and use the default (vbOKonly) button style and default (app.title) title text
'
'Enjoy!


'Win32 API decs

'Hook functions
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Constants
Private Const WH_CBT            As Long = 5
Private Const HCBT_ACTIVATE     As Long = 5
Private Const HWND_TOP          As Long = 0
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const STM_SETICON       As Long = &H170

'APP-SPECIFIC
Private Const SWVB_DEFAULT      As Long = &HFFFFFFFF '-1 is reserved for centering
Private Const SWVB_CAPTION_DEFAULT As String = "SWVB_DEFAULT_TO_APP_TITLE"

''Types
'Private Type RECT
'   Left As Long
'   Top As Long
'   Right As Long
'   Bottom As Long
'End Type

'module-level member variables
Private m_Hook As Long
Private m_Left As Long
Private m_Top As Long
Private m_hIcon As Long


Public Function MsgBoxEx(ByVal Prompt As String, _
                ByVal sContent As String, _
                Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                Optional ByVal Title As String = SWVB_CAPTION_DEFAULT, _
                Optional ByVal Left As Long = SWVB_DEFAULT, _
                Optional ByVal Top As Long = SWVB_DEFAULT, _
                Optional ByVal Icon As Long = 0&, _
                Optional ByVal bNoTaskDialog As Boolean = False, _
                Optional ByVal lOwnerhWnd As Long = 0&) As VbMsgBoxResult

Dim hInst As Long
Dim threadID As Long
Dim wndRect As RECT

Dim bDoMsgBox As Boolean
Dim lRet As VbMsgBoxResult

hInst = App.hInstance
threadID = GetCurrentThreadId()

'Save the new arguments as member variables to be used from the MsgBoxHook proc
m_Left = Left
m_Top = Top
m_hIcon = Icon

'default the msgBox caption to app.title
If Title = SWVB_CAPTION_DEFAULT Then
    Title = App.Title
End If

'if user wants custom icon make sure dialog has an icon to replace
'If m_hIcon <> 0& Then
    'Buttons = Buttons Or vbInformation
'End If

'decide whether to show taskdialog
If modLoadProgram.bVistaOrW7 Then
    If modDisplay.VisualStyle Then
        bDoMsgBox = False
    Else
        bDoMsgBox = True
    End If
Else
    bDoMsgBox = True
End If


If bNoTaskDialog Then
    bDoMsgBox = True
End If


'AddConsoleText "MsgBoxEx Called", , True
'AddConsoleText "bDoMsgbox = " & CStr(bDoMsgBox)

If bDoMsgBox = False Then
    
    If lOwnerhWnd = 0 Then
        lOwnerhWnd = frmMain.hWnd
    End If
    
    'lOwnerhWnd can be 0
    lRet = modVista.TaskDialog(Title, Prompt, sContent, Buttons, lOwnerhWnd, VB_To_TDIcon(Buttons))
    
    'AddConsoleText "lRet = " & CStr(lRet)
    
    If lRet = -1 Then
        'fail, do msgboxEx
        bDoMsgBox = True
    Else
        MsgBoxEx = lRet
    End If
End If



If bDoMsgBox Then
    'First "subclass" the MsgBox function
    m_Hook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHook, hInst, threadID)
    
    'show the MsgBox and let hook proc take care of the rest...
    MsgBoxEx = MsgBox(Prompt, Buttons, Title)
End If



'just in case...
If m_Hook Then
    UnHook
End If


'AddConsoleText "Exiting MsgBoxEx...", , , True

End Function

Private Function MsgBoxHook(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim height As Long
Dim width As Long
Dim nSize As Long
Dim wndRect As RECT
Dim sBuffer As String
Dim fWidth As Long
Dim fHeight As Long
Dim X As Long
Dim Y As Long
Dim hIconWnd As Long

'Debug.Print "hook proc called"
'Call next hook in the chain and return the value
'(this is the polite way to allow other hooks to function too)
MsgBoxHook = CallNextHookEx(m_Hook, nCode, wParam, lParam)


' hook only the activate msg
If nCode = HCBT_ACTIVATE Then
    'handle only standard MsgBox class windows
    sBuffer = Space$(32) 'this is the most efficient method to allocate strings in VB
                         'according to Brad Martinez's results with tools from NuMega
    
    nSize = GetClassName(wParam, sBuffer, 32&) 'GetClassName will truncate the class name if it doesn't fit in the buffer
                                              'we only care about the first 6 chars anyway
    If Left$(sBuffer, nSize) <> "#32770" Then
        Exit Function 'not a standard msgBox
                      'we can just quit because we already called CallNextHookEx
    End If
     
    'store MsgBox window size in case we need it
    Call GetWindowRect(wParam, wndRect)
    
    'handle divide by zero errors (should never happen)
    On Error GoTo errorTrap
    height = (wndRect.Bottom - wndRect.Top) / 2
    width = (wndRect.Right - wndRect.Left) / 2
    
    'store parent window size
    Call GetWindowRect(GetParent(wParam), wndRect)
    
    'handle divide by zero errors (should never happen)
    On Error GoTo errorTrap
    fHeight = wndRect.Top + (wndRect.Bottom - wndRect.Top) / 2
    fWidth = wndRect.Left + (wndRect.Right - wndRect.Left) / 2
    
    'By default centre MsgBox on the form
    'if user passed in specific values then use those instead
    If m_Left = SWVB_DEFAULT Then 'default
        X = fWidth - width
    Else
        X = m_Left
    End If
    
    If m_Top = SWVB_DEFAULT Then 'default
        Y = fHeight - height
    Else
        Y = m_Top
    End If

    'Manually set the MsgBox window position before Windows shows it
    SetWindowPos wParam, HWND_TOP, X, Y, 0, 0, SWP_NOSIZE + SWP_NOZORDER + SWP_NOACTIVATE
    
    'If user passed in custom icon use that instead of the standard Windows icon
    If m_hIcon Then
        hIconWnd = FindWindowEx(wParam, 0&, "Static", vbNullString)
        SendMessageByLong hIconWnd, STM_SETICON, m_hIcon, 0&
    End If

errorTrap:
    'unhook the dialog and we are out clean!
    UnHook
    'Debug.Print "unhook"
End If

End Function

Private Sub UnHook()
UnhookWindowsHookEx m_Hook
m_Hook = 0
End Sub
