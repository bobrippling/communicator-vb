VERSION 5.00
Begin VB.Form frmInfo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Communicator Info Window"
   ClientHeight    =   645
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrInfo 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pHasFocus As Boolean
Public bIgnoreLostFocus As Boolean

Private pTransparency As Byte
Private Const Trans_Focus = 235
Private Const Trans_NoFocus = 160

'Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

'######################################################

Private Property Let Transparency(nVal As Byte)

modDisplay.SetTransparency Me.hWnd, nVal
pTransparency = nVal

End Property

'######################################################################################################

Public Sub Form_apiGotFocus()

If Not pHasFocus Then
    pHasFocus = True
    Transparency = Trans_Focus
    EnableTracking
End If

End Sub

Public Sub Form_apiLostFocus()

If pHasFocus Then
    If bIgnoreLostFocus = False Then
        pHasFocus = False
        Transparency = Trans_NoFocus
    End If
End If

End Sub

Private Sub EnableTracking()
modDisplay.EnableMouseTracking Me.hWnd
End Sub

'######################################################################################################
'load events

Private Sub Form_Load()

bIgnoreLostFocus = False
pHasFocus = False


mnuInfoPopupTop_Click

DrawBorder Me
Me.Move Screen.width / 2 - Me.width / 2, Screen.height / 2 - Me.height / 2

modDisplay.SetTransparentStyle Me.hWnd
Transparency = Trans_NoFocus
modSubClass.SubclassAuto Me


FormLoad Me, , False, False

EnableTracking

frmMini_Loaded = True
frmMain.mnuFileInfo.Checked = True

Me.Show

tmrMain_Timer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

modSubClass.SubclassAuto Me, False
modDisplay.SetTransparentStyle Me.hWnd, False

FormLoad Me, True, False

frmInfo_Loaded = False
frmMain.mnuFileInfo.Checked = False

End Sub


Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Dim RC As RECT

Select Case uMsg
    Case WM_MOUSELEAVE
        Form_apiLostFocus
        
    Case WM_MOUSEMOVE
        Form_apiGotFocus
        
    Case WM_KILLFOCUS
        bIgnoreLostFocus = False
        Form_apiLostFocus
        
    Case WM_MOVING
        
        If frmSystray.mnuInfoPopupDock.Checked Then
            CopyMemory RC, ByVal lParam, Len(RC)
            Form_Moving RC
            CopyMemory ByVal lParam, RC, Len(RC)
        End If
        
End Select

WindowProc = CallWindowProc(GetProp(hWnd, WndProcStr), hWnd, uMsg, wParam, lParam)

End Function

'######################################################################################################
'mouse events

Private Sub Form_DblClick()
frmMain.ShowForm (Not frmMain.Visible)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bIgnoreLostFocus = False
EnableTracking
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Not frmSystray.mnuInfoPopupLock.Checked Then
        ReleaseCapture
        SendMessageByLong Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    bIgnoreLostFocus = True
    Me.PopupMenu frmSystray.mnuInfoPopup
End If

End Sub

Private Sub Form_Moving(RC As RECT)
Const Dock_Lim = 100 'px
Dim ptCursor As POINTAPI
Dim ScreenSize As Long

GetCursorPos ptCursor


With RC
    ScreenSize = Screen.width / Screen.TwipsPerPixelX
    If ptCursor.X < Dock_Lim Then
        '.Right = .Right - .Left + 1
        .Left = 0
        .Right = ScaleX(Me.width, vbTwips, vbPixels)
        
    ElseIf ptCursor.X > (ScreenSize - Dock_Lim) Then
        
        .Left = ScreenSize - ScaleX(Me.width, vbTwips, vbPixels)
        .Right = ScreenSize
    End If
    
    
    
    ScreenSize = (Screen.height - modVars.GetTaskbarHeight()) / Screen.TwipsPerPixelY
    If ptCursor.Y < Dock_Lim Then
        '.Right = .Right - .Left + 1
        .Top = 0
        .Bottom = ScaleY(Me.height, vbTwips, vbPixels)
        
    ElseIf ptCursor.Y > (ScreenSize - Dock_Lim) Then
        
        .Top = ScreenSize - ScaleY(Me.height, vbTwips, vbPixels)
        .Bottom = ScreenSize
        
    End If
End With



End Sub

'######################################################################################################
'menu clicks

Public Sub mnuInfoPopupTop_Click()
Dim B As Boolean

B = Not frmSystray.mnuInfoPopupTop.Checked
frmSystray.mnuInfoPopupTop.Checked = B

modVars.SetOnTop Me.hWnd, B, False

End Sub

Public Sub mnuInfoPopupLock_Click()

frmSystray.mnuInfoPopupLock.Checked = Not frmSystray.mnuInfoPopupLock.Checked

End Sub

Public Sub mnuInfoPopupClose_Click()

Unload Me

End Sub

Public Sub mnuInfoPopupDock_click()

frmSystray.mnuInfoPopupDock.Checked = Not frmSystray.mnuInfoPopupDock.Checked

End Sub
'######################################################################################################

Private Sub tmrInfo_Timer()
Dim sTxt As String
Dim dDate As Date
Const X_Indent = 30

Me.Cls
DrawBorder Me

sTxt = modMessaging.LastSender
dDate = modMessaging.LastMessageTime

PrintText "Communicator", X_Indent, 10

If LenB(sTxt) = 0 Or dDate = 0 Then
    If Status = Connected Then
        PrintText "No one has said anything yet...", X_Indent
    Else
        PrintText "Not connected...", X_Indent
    End If
Else
    PrintText "Last message received at " & Format$(dDate, "hh:mm AM/PM"), X_Indent
    PrintText "from " & sTxt, X_Indent
End If

End Sub

Private Sub PrintText(sTxt As String, Optional X As Single = -1, Optional Y As Single = -1)
If X >= 0 Then CurrentX = X
If Y >= 0 Then CurrentY = Y
Me.Print sTxt
End Sub
