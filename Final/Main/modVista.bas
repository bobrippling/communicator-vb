Attribute VB_Name = "modVista"
Option Explicit

'=======================================
'Downloaded from Visual Basic Thunder
'www.vbthunder.com
'Created on: 3/07/2006
'=======================================
'Last modified 3/07/2006
'=======================================

'Private Const WM_USER = &H400&

'// ===================== Task Dialog =========================
'#ifndef NOTASKDIALOG
'
'#ifdef _WIN32
'#include <pshpack1.h>
'#End If
'
'typedef HRESULT (CALLBACK *PFTASKDIALOGCALLBACK)(__in HWND hwnd, __in UINT msg, __in WPARAM wParam, __in LPARAM lParam, __in LONG_PTR lpRefData);

'Private Enum TASKDIALOG_FLAGS
'    TDF_ENABLE_HYPERLINKS = &H1&
'    TDF_USE_HICON_MAIN = &H2&
'    TDF_USE_HICON_FOOTER = &H4&
'    TDF_ALLOW_DIALOG_CANCELLATION = &H8&
'    TDF_USE_COMMAND_LINKS = &H10&
'    TDF_USE_COMMAND_LINKS_NO_ICON = &H20&
'    TDF_EXPAND_FOOTER_AREA = &H40&
'    TDF_EXPANDED_BY_DEFAULT = &H80&
'    TDF_VERIFICATION_FLAG_CHECKED = &H100&
'    TDF_SHOW_PROGRESS_BAR = &H200&
'    TDF_SHOW_MARQUEE_PROGRESS_BAR = &H400&
'    TDF_CALLBACK_TIMER = &H800&
'    TDF_POSITION_RELATIVE_TO_WINDOW = &H1000&
'    TDF_RTL_LAYOUT = &H2000&
'    TDF_NO_DEFAULT_RADIO_BUTTON = &H4000&
'End Enum
'typedef DWORD TASKDIALOG_FLAGS;                         // Note: _TASKDIALOG_FLAGS is an int

'Private Enum TASKDIALOG_MESSAGES
'    TDM_NAVIGATE_PAGE = WM_USER + 101&
'    TDM_CLICK_BUTTON = WM_USER + 102&                  '// wParam = Button ID
'    TDM_SET_MARQUEE_PROGRESS_BAR = WM_USER + 103&      '// wParam = 0 (nonMarque) wParam != 0 (Marquee)
'    TDM_SET_PROGRESS_BAR_STATE = WM_USER + 104&        '// wParam = new progress state
'    TDM_SET_PROGRESS_BAR_RANGE = WM_USER + 105&        '// lParam = MAKELPARAM(nMinRange& nMaxRange)
'    TDM_SET_PROGRESS_BAR_POS = WM_USER + 106&          '// wParam = new position
'    TDM_SET_PROGRESS_BAR_MARQUEE = WM_USER + 107&      '// wParam = 0 (stop marquee), wParam != 0 (start marquee), lparam = speed (milliseconds between repaints)
'    TDM_SET_ELEMENT_TEXT = WM_USER + 108&              '// wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
'    TDM_CLICK_RADIO_BUTTON = WM_USER + 110&            '// wParam = Radio Button ID
'    TDM_ENABLE_BUTTON = WM_USER + 111&                 '// lParam = 0 (disable), lParam != 0 (enable), wParam = Button ID
'    TDM_ENABLE_RADIO_BUTTON = WM_USER + 112&           '// lParam = 0 (disable), lParam != 0 (enable), wParam = Radio Button ID
'    TDM_CLICK_VERIFICATION = WM_USER + 113&            '// wParam = 0 (unchecked), 1 (checked), lParam = 1 (set key focus)
'    TDM_UPDATE_ELEMENT_TEXT = WM_USER + 114&           '// wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
'    TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE = WM_USER + 115& '// wParam = Button ID, lParam = 0 (elevation not required)& lParam != 0 (elevation required)
'    TDM_UPDATE_ICON = WM_USER + 116&                   '// wParam = icon element (TASKDIALOG_ICON_ELEMENTS), lParam = new icon (hIcon if TDF_USE_HICON_* was set, PCWSTR otherwise)
'End Enum 'TASKDIALOG_MESSAGES;

'Private Enum TASKDIALOG_NOTIFICATIONS
'    TDN_CREATED = 0
'    TDN_NAVIGATED = 1
'    TDN_BUTTON_CLICKED = 2           '// wParam = Button ID
'    TDN_HYPERLINK_CLICKED = 3        '// lParam = (LPCWSTR)pszHREF
'    TDN_TIMER = 4                    '// wParam = Milliseconds since dialog created or timer reset
'    TDN_DESTROYED = 5
'    TDN_RADIO_BUTTON_CLICKED = 6     '// wParam = Radio Button ID
'    TDN_DIALOG_CONSTRUCTED = 7
'    TDN_VERIFICATION_CLICKED = 8     '// wParam = 1 if checkbox checked, 0 if not, lParam is unused and always 0
'    TDN_HELP = 9
'    TDN_EXPANDO_BUTTON_CLICKED = 10  '// wParam = 0 (dialog is now collapsed), wParam != 0 (dialog is now expanded)
'End Enum 'TASKDIALOG_NOTIFICATIONS;

'Private Type TASKDIALOG_BUTTON
'    nButtonID As Long 'int
'    pszButtonText As Long 'PCWSTR
'End Type 'TASKDIALOG_BUTTON;

'Private Enum TASKDIALOG_ELEMENTS
'    TDE_CONTENT = 0
'    TDE_EXPANDED_INFORMATION = 1
'    TDE_FOOTER = 2
'    TDE_MAIN_INSTRUCTION = 3
'End Enum 'TASKDIALOG_ELEMENTS;
'
'Private Enum TASKDIALOG_ICON_ELEMENTS
'    TDIE_ICON_MAIN = 0
'    TDIE_ICON_FOOTER = 1
'End Enum 'TASKDIALOG_ICON_ELEMENTS;

Private Const TD_WARNING_ICON As Integer = -1       'MAKEINTRESOURCEW(-1)
Private Const TD_ERROR_ICON As Integer = -2         'MAKEINTRESOURCEW(-2)
Private Const TD_INFORMATION_ICON As Integer = -3   'MAKEINTRESOURCEW(-3)
Private Const TD_SHIELD_ICON As Integer = -4        'MAKEINTRESOURCEW(-4)

Private Enum TASKDIALOG_COMMON_BUTTON_FLAGS
    TDCBF_OK_BUTTON = &H1&               '// selected control return value IDOK
    TDCBF_YES_BUTTON = &H2&              '// selected control return value IDYES
    TDCBF_NO_BUTTON = &H4&               '// selected control return value IDNO
    TDCBF_CANCEL_BUTTON = &H8&           '// selected control return value IDCANCEL
    TDCBF_RETRY_BUTTON = &H10&           '// selected control return value IDRETRY
    TDCBF_CLOSE_BUTTON = &H20&           '// selected control return value IDCLOSE
End Enum
'typedef DWORD TASKDIALOG_COMMON_BUTTON_FLAGS;           // Note: _TASKDIALOG_COMMON_BUTTON_FLAGS is an int

'Private Type TASKDIALOGCONFIG
'    cbSize As Long 'UINT
'    hwndParent As Long 'HWND
'    hInstance As Long 'HINSTANCE                        // used for MAKEINTRESOURCE() strings
'    dwFlags As TASKDIALOG_FLAGS 'TASKDIALOG_FLAGS       // TASKDIALOG_FLAGS (TDF_XXX) flags
'    dwCommonButtons As TASKDIALOG_COMMON_BUTTON_FLAGS ' // TASKDIALOG_COMMON_BUTTON (TDCBF_XXX) flags
'    pszWindowTitle As Long 'PCWSTR                      // string or MAKEINTRESOURCE()
''    Union
''    {
'        hMainIcon As Long
''        HICON   hMainIcon;
''        PCWSTR  pszMainIcon;
''    };
'    pszMainInstruction As Long 'PCWSTR
'    pszContent As Long 'PCWSTR
'    cButtons As Long 'UINT
'    pButtons As Long 'const TASKDIALOG_BUTTON  *pButtons;
'    nDefaultButton As Long 'int
'    cRadioButtons As Long 'UINT
'    pRadioButtons As Long 'const TASKDIALOG_BUTTON  *pRadioButtons;
'    nDefaultRadioButton As Long 'int
'    pszVerificationText As Long 'PCWSTR
'    pszExpandedInformation As Long 'PCWSTR
'    pszExpandedControlText As Long 'PCWSTR
'    pszCollapsedControlText As Long 'PCWSTR
'    'Union
'    '{
'        hFooterIcon As Long
'    '    HICON   hFooterIcon;
'    '    PCWSTR  pszFooterIcon;
'    '};
'    pszFooter As Long 'PCWSTR
'    pfCallback As Long 'PFTASKDIALOGCALLBACK
'    lpCallbackData As Long 'LONG_PTR
'    cxWidth As Long 'UINT             // width of the Task Dialog's client area in DLU's. If 0, Task Dialog will calculate the ideal width.
'End Type

'WINCOMMCTRLAPI HRESULT WINAPI TaskDialogIndirect(const TASKDIALOGCONFIG *pTaskConfig, __out_opt int *pnButton, __out_opt int *pnRadioButton, __out_opt BOOL *pfVerificationFlagChecked);
'Private Declare Function apiTaskDialogIndirect Lib "comctl32.dll" Alias "TaskDialogIndirect" ( _
    pTaskConfig As TASKDIALOGCONFIG, pnButton As Long, _
    pnRadioButton As Long, pfVerificationFlagChecked As Long) As Long
'WINCOMMCTRLAPI HRESULT WINAPI TaskDialog(__in_opt HWND hwndParent, __in_opt HINSTANCE hInstance, __in_opt PCWSTR pszWindowTitle, __in_opt PCWSTR pszMainInstruction, __in_opt PCWSTR pszContent, TASKDIALOG_COMMON_BUTTON_FLAGS dwCommonButtons, __in_opt PCWSTR pszIcon, __out_opt int *pnButton);
Private Declare Function apiTaskDialog Lib "comctl32.dll" Alias "TaskDialog" ( _
    ByVal hwndParent As Long, ByVal hInstance As Long, ByVal pszWindowTitle As Long, _
    ByVal pszMainInstruction As Long, ByVal pszContent As Long, _
    ByVal dwCommonButtons As TASKDIALOG_COMMON_BUTTON_FLAGS, _
    ByVal pszIcon As Long, pnButton As Long) As Long

'#ifdef _WIN32
'#include <poppack.h>
'#End If
'
'#endif // NOTASKDIALOG

'// ==================== End TaskDialog =======================



Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const IDOK = 1
Private Const IDCANCEL = 2
Private Const IDABORT = 3
Private Const IDRETRY = 4
Private Const IDIGNORE = 5
Private Const IDYES = 6
Private Const IDNO = 7

Public Enum eTD_ICONS
    eTD_WARNING_ICON = TD_WARNING_ICON
    eTD_ERROR_ICON = TD_ERROR_ICON
    eTD_INFORMATION_ICON = TD_INFORMATION_ICON
    eTD_SHIELD_ICON = TD_SHIELD_ICON
End Enum

'##########################################################################################
''Remove Titlebar + Icon(Vista)
'
'
'Private Enum WINDOWTHEMEATTRIBUTETYPE
'    WTA_NONCLIENT = 1
'End Enum
'Private Type WTA_OPTIONS
'    dwFlags As Long
'    dwMask As Long
'End Type
'
'Private Declare Function SetWindowThemeAttribute Lib "uxTheme.dll" ( _
'    ByVal hWnd As Long, ByVal eAttribute As WINDOWTHEMEATTRIBUTETYPE, _
'    ByRef pvAttribute As WTA_OPTIONS, ByVal cbAttribute As Long) As Long
'
'Private Const WTNCA_NODRAWCAPTION = &H1
''Prevents the window caption from being drawn.
'Private Const WTNCA_NODRAWICON = &H2
''Prevents the system icon from being drawn.
'Private Const WTNCA_NOSYSMENU = &H4
''Prevents the system icon menu from appearing.

'##########################################################################################


Public Function TaskDialog(ByVal sWindowTitle As String, ByVal sMainInstruction As String, _
    ByVal sContent As String, lStyle As VbMsgBoxStyle, hWnd As Long, _
    Optional pIcon As eTD_ICONS = eTD_ICONS.eTD_INFORMATION_ICON) As VbMsgBoxResult


Dim ClickedButton As Long, lResult As Long

On Error GoTo EH
lResult = apiTaskDialog(hWnd, 0, _
    ByVal StrPtr(sWindowTitle), _
    ByVal StrPtr(sMainInstruction), _
    ByVal StrPtr(sContent), _
    VBStyle_To_TaskDialogStyle(lStyle), _
    MAKEINTRESOURCE(pIcon), _
    ClickedButton)

If lResult = 0 Then
    If ClickedButton = IDOK Then
        TaskDialog = vbOK
    ElseIf ClickedButton = IDCANCEL Then
        TaskDialog = vbCancel
    ElseIf ClickedButton = IDABORT Then
        TaskDialog = vbAbort
    ElseIf ClickedButton = IDRETRY Then
        TaskDialog = vbRetry
    ElseIf ClickedButton = IDIGNORE Then
        TaskDialog = vbIgnore
    ElseIf ClickedButton = IDYES Then
        TaskDialog = vbYes
    ElseIf ClickedButton = IDNO Then
        TaskDialog = vbNo
    Else
        TaskDialog = vbOK
    End If
Else
    TaskDialog = -1
End If

Exit Function
EH:
'can't do taskdialog
TaskDialog = -1
End Function

Private Function VBStyle_To_TaskDialogStyle(lStyle As VbMsgBoxStyle) As TASKDIALOG_COMMON_BUTTON_FLAGS

If (lStyle And vbOKCancel) = vbOKCancel Then
    VBStyle_To_TaskDialogStyle = TDCBF_OK_BUTTON + TDCBF_CANCEL_BUTTON
    
ElseIf (lStyle And vbYesNo) = vbYesNo Then
    VBStyle_To_TaskDialogStyle = TDCBF_YES_BUTTON + TDCBF_NO_BUTTON
    
ElseIf (lStyle And vbYesNoCancel) = vbYesNoCancel Then
    VBStyle_To_TaskDialogStyle = TDCBF_YES_BUTTON + TDCBF_NO_BUTTON + TDCBF_CANCEL_BUTTON
    
ElseIf (lStyle And vbAbortRetryIgnore) = vbAbortRetryIgnore Then
    VBStyle_To_TaskDialogStyle = TDCBF_CANCEL_BUTTON + TDCBF_RETRY_BUTTON + TDCBF_CLOSE_BUTTON
    
ElseIf (lStyle And vbRetryCancel) = vbRetryCancel Then
    VBStyle_To_TaskDialogStyle = TDCBF_RETRY_BUTTON + TDCBF_CANCEL_BUTTON
    
Else
    VBStyle_To_TaskDialogStyle = TDCBF_OK_BUTTON
    
End If

End Function

Public Function VB_To_TDIcon(lStyle As VbMsgBoxStyle) As eTD_ICONS

If (lStyle And vbInformation) = vbInformation Then
    VB_To_TDIcon = eTD_INFORMATION_ICON
    
ElseIf (lStyle And vbExclamation) = vbExclamation Then
    VB_To_TDIcon = eTD_ERROR_ICON
    
ElseIf (lStyle And vbCritical) = vbCritical Then
    VB_To_TDIcon = eTD_WARNING_ICON
    
Else
    VB_To_TDIcon = eTD_INFORMATION_ICON
End If

End Function

Private Function MAKEINTRESOURCE(ByVal iVal As Integer) As Long

Dim l As Long

l = 0
CopyMemory l, iVal, 2
MAKEINTRESOURCE = l

End Function

''##############################################################################
'
'Public Function TaskDialogIndirect( _
'    ByVal sWindowTitle As String, ByVal sMainInstruction As String, ByVal sContent As String, _
'    ByVal sCollapsedControlText As String, ByVal sExpandedControlText As String, ByVal sExpandedText As String, _
'    lStyle As VbMsgBoxStyle, hWnd As Long, Optional pIcon As eTD_ICONS = eTD_ICONS.eTD_INFORMATION_ICON) As VbMsgBoxResult
'
'Dim TDlg As TASKDIALOGCONFIG
'Dim ClickedButton As Long
'Dim lResult As Long
'
'
'Dim tbButtons(0 To 2) As TASKDIALOG_BUTTON
'Dim sBtn1 As String
'Dim sBtn2 As String
'Dim sBtn3 As String
'
'sBtn1 = "Yeah"
'sBtn2 = "Nah"
'sBtn3 = "Repeat the question?"
'tbButtons(0).nButtonID = IDYES
'tbButtons(0).pszButtonText = StrPtr(sBtn1)
'tbButtons(1).nButtonID = IDNO
'tbButtons(1).pszButtonText = StrPtr(sBtn2)
'tbButtons(2).nButtonID = IDCANCEL
'tbButtons(2).pszButtonText = StrPtr(sBtn3)
'
'Dim sVerify As String
'Dim sExpanded As String
'Dim sExpandedControlText As String
'Dim sCollapsedControlText As String
'
'TDlg.cbSize = Len(TDlg)
'TDlg.hwndParent = Me.hWnd
'TDlg.hInstance = 0
'TDlg.dwFlags = 0
'TDlg.dwCommonButtons = 0
'
'TDlg.pszWindowTitle = StrPtr(sWindowTitle)
'TDlg.hMainIcon = 0 'You can specify a custom icon here
'TDlg.pszMainInstruction = StrPtr(sMainInstruction)
'TDlg.pszContent = StrPtr(sContent)
'
'TDlg.cButtons = 3 'Three custom buttons in this example
'TDlg.pButtons = VarPtr(tbButtons(0).nButtonID)
'
'TDlg.cRadioButtons = 3 'For the sake of simplicity, use the same array for the radio buttons as well
'TDlg.pRadioButtons = VarPtr(tbButtons(0).nButtonID)
'
'TDlg.nDefaultButton = IDNO
'TDlg.pszVerificationText = StrPtr(sVerify)
'TDlg.pszExpandedInformation = StrPtr(sExpanded)
'TDlg.pszExpandedControlText = StrPtr(sExpandedControlText)
'TDlg.pszCollapsedControlText = StrPtr(sCollapsedControlText)
'TDlg.hFooterIcon = MAKEINTRESOURCE(TD_INFORMATION_ICON)
'TDlg.pszFooter = 0
'TDlg.pfCallback = 0
'
'lResult = apiTaskDialogIndirect(TDlg, ClickedButton, SelRadio, fVerify)
'
'
'End Function


'##################################################################################################

'Public Sub HideCaption(hWnd As Long, Optional ByVal bHide As Boolean = True)
'Dim Ops As WTA_OPTIONS
'
''"If we set the Mask to the same value as the Flags, the Flags are Added. If not they are Removed"
'
''If Mask = Flags Then
''    AddFlags
''Else
''    RemoveFlags
''End If
'
'Ops.dwFlags = WTNCA_NODRAWCAPTION Or WTNCA_NODRAWICON Or WTNCA_NOSYSMENU
'
'If bHide Then
'    Ops.dwMask = WTNCA_NODRAWCAPTION Or WTNCA_NODRAWICON Or WTNCA_NOSYSMENU
'Else
'    Ops.dwMask = -1#
'End If
'
'On Error Resume Next
'SetWindowThemeAttribute hWnd, WINDOWTHEMEATTRIBUTETYPE.WTA_NONCLIENT, Ops, Len(Ops)
'
'End Sub
