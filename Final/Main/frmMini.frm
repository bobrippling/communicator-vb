VERSION 5.00
Begin VB.Form frmMini 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Communicator Mini Window"
   ClientHeight    =   1320
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkComm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Communicator"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.Timer tmrReset 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   3960
      Top             =   360
   End
   Begin VB.PictureBox picClose 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Info Label"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   4095
   End
   Begin VB.Image imgIcon 
      Height          =   375
      Left            =   60
      Top             =   60
      Width           =   375
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Communicator Mini Window"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   60
      Width           =   2895
   End
   Begin VB.Line lnSep 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   1920
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmMini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bIgnorechkComm As Boolean

'########################################################################

'show in taskbar
Private Const WS_EX_APPWINDOW = &H40000

'link select
Private Const IDC_HAND = 32649
'Private Const IDC_ARROW = 32512
'Private Const GCW_HCURSOR = (-12)
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
'Private Declare Function SetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
'Private hCursor As Long, hOldCursor As Long

Private Const GreyForecolour = &H707070

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

Private pHasFocus As Boolean, bIgnoreLostFocus As Boolean

Private pTransparency As Byte
Private Const Trans_Focus = 235
Private Const Trans_NoFocus = 160

'######################################################
'Private Const NormHeight = 1155, ExHeight = 2295

'Private Sub chkBandwidth_Click()
'If chkBandwidth.Value = 1 Then
'    Me.height = ExHeight
'    tmrBandwidth.Enabled = True
'    Me.Top = Me.Top - ExHeight + NormHeight
'    tmrBandwidth_Timer
'Else
'    Me.height = NormHeight
'    tmrBandwidth.Enabled = False
'    Me.Top = Me.Top + ExHeight - NormHeight
'End If
'
'Me.Cls
'DrawBorder Me
'End Sub

'Private Sub chkBandwidth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'bIgnoreLostFocus = True
''EnableTracking
'
'Form_apiGotFocus
'End Sub
'
'Private Sub Grph_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form_MouseDown Button, Shift, X, Y
'End Sub
'
'Private Sub picBandwidth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form_MouseDown Button, Shift, X, Y
'End Sub
'
'Private Sub picBandwidth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'chkBandwidth_MouseMove Button, Shift, X, Y
'End Sub
'
'Private Sub tmrBandwidth_Timer()
'Dim Down As Single, Up As Single
'Const H = 195
'
'modNetwork.GetSpeeds Down, Up
'
'picBandwidth.Cls
'picBandWidthDraw "Download Speed: " & CStr(Round(Down, 2)) & " KB/S", 10, H, vbGreen
'picBandWidthDraw "Upload Speed: " & CStr(Round(Up, 2)) & " KB/S", 10, H * 2, vbRed
'
'On Error GoTo GraphEH
'
'With Grph
'    .Redraw = False
'
'    If .MaxValue < Down Then
'        .MaxValue = Down
'    ElseIf .MaxValue < Up Then
'        .MaxValue = Up
'    End If
'
'    .Datasets.Item(1).Points.Add Down
'    .Datasets.Item(2).Points.Add Up
'    .Redraw = True
'End With
'
'Exit Sub
'GraphEH:
'
'chkBandwidth.Value = 0
'
'AddText "Error With Bandwidth Monitor: " & Err.Description, TxtError, True
'
'End Sub
'
'Private Sub picBandWidthDraw(sTxt As String, X As Single, Y As Single, Col As Long)
'
'picBandwidth.CurrentX = X
'picBandwidth.CurrentY = Y
'picBandwidth.ForeColor = Col
'picBandwidth.Print sTxt
'
'End Sub
'
'Private Sub SetupDatasets()
'Dim objDataset As Dataset
'
'Set objDataset = Grph.Datasets.Add()
'With objDataset
'    .Visible = True
'    .ShowPoints = False
'    .ShowBars = False
'    .ShowLines = True
'    .ShowCaps = False
'    .LineColor = vbGreen
'End With
'
'Set objDataset = Nothing
'
'Set objDataset = Grph.Datasets.Add()
'With objDataset
'    .Visible = True
'    .ShowPoints = False
'    .ShowBars = False
'    .ShowLines = True
'    .ShowCaps = False
'    .LineColor = vbRed
'End With
'
'Set objDataset = Nothing
'
'End Sub

'######################################################

Private Property Let Transparency(nVal As Byte)

SetTrans nVal
pTransparency = nVal

End Property

Private Sub SetTrans(btLevel As Byte)
'0 = completely transparent, 255 = completely opaque
modDisplay.SetTransparency Me.hWnd, btLevel
End Sub

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
Dim ET As TRACKMOUSEEVENTTYPE

'initialize structure
ET.cbSize = Len(ET)
ET.hwndTrack = Me.hWnd
ET.dwFlags = TME_LEAVE

'start the tracking
TrackMouseEvent ET

End Sub

'######################################################################################################

Private Sub chkComm_Click()
If Not bIgnorechkComm Then frmMain.ShowForm CBool(chkComm.Value)
End Sub

Public Sub setchkComm_Value(iVal As Integer)
bIgnorechkComm = True
chkComm.Value = iVal
bIgnorechkComm = False
End Sub

Private Sub chkComm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bIgnoreLostFocus = True
'EnableTracking

Form_apiGotFocus
End Sub

'######################################################################################################

Private Sub lblCaption_DblClick()
Form_DblClick
End Sub

Private Sub imgIcon_DblClick()
Form_DblClick
End Sub

Private Sub lblInfo_DblClick()
Form_DblClick
End Sub

Private Sub Form_DblClick()
frmMain.ShowForm Not frmMain.Visible
End Sub

'######################################################################################################

Private Sub Form_Load()

'SetupDatasets

'Me.height = NormHeight

bIgnorechkComm = True
bIgnoreLostFocus = False
pHasFocus = False

Set imgIcon.Picture = frmSystray.img16x16.ListImages(1).Picture
lblCaption.ForeColor = GreyForecolour
Me.Caption = lblCaption.Caption
lnSep.X2 = Me.width + 10

DrawBorder Me
InitPictureBoxes
tmrReset_Timer

Me.chkComm.Value = Abs(frmMain.Visible)

modDisplay.SetTransparentStyle Me.hWnd
Transparency = Trans_NoFocus
modSubClass.SubclassAuto Me

'picMin.Visible = Not modLoadProgram.bVistaOrW7

mnuPopupReset_Click

FormLoad Me, , False, False
'Me.Show vbModeless
Me.Visible = False
SetOnTop frmMini.hWnd 'instead of show

EnableTracking

frmMini_Loaded = True
frmMain.mnuFileMini.Checked = True
bIgnorechkComm = False

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bIgnoreLostFocus = False
EnableTracking
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
    Cancel = True
Else
    modSubClass.SubclassAuto Me, False
    modDisplay.SetTransparentStyle Me.hWnd, False
    frmMini_Loaded = False
    frmMain.mnuFileMini.Checked = False
End If

End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub imgIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    
    If Me.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessageByLong Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
        
        SetInfo "Right Click to Reset Position"
    End If
    
ElseIf Button = vbRightButton Then
    mnuPopupReset_Click
    
    SetInfo "I'm back down here!"
    'PopupMenu frmSystray.mnuMiniPopup, , , , frmSystray.mnuMiniPopupComm
End If

End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long


Select Case uMsg
    Case WM_MOUSELEAVE
        Form_apiLostFocus
        
    Case WM_MOUSEMOVE
        Form_apiGotFocus
        
End Select

WindowProc = CallWindowProc(GetProp(hWnd, WndProcStr), hWnd, uMsg, wParam, lParam)

End Function

'##########################################################################################

Private Sub InitPictureBoxes()
Const HeightLim = 100, MinLim = 10, Max = 255, dWidth = 2

picClose.BorderStyle = 0
'picMin.BorderStyle = 0

picClose.ForeColor = GreyForecolour
'picMin.ForeColor = GreyForecolour

picClose.DrawWidth = dWidth
'picMin.DrawWidth = dWidth

picClose.MousePointer = vbIconPointer
'picMin.MousePointer = vbIconPointer


'min lines
'picMin.Line (MinLim + 100, Max - HeightLim)-(Max - MinLim, Max - HeightLim)

'close lines
picClose.Line (0, 0)-(Max - HeightLim, Max - HeightLim)
picClose.Line (0, Max - HeightLim)-(Max - HeightLim, 0)

End Sub

'###################################################################################################

Public Sub SetInfo(ByVal sInfo As String)

lblInfo.Caption = sInfo
tmrReset.Enabled = True

End Sub

Private Sub tmrReset_Timer()
tmrReset.Enabled = False

SetInfo GetStatus()
End Sub

'###################################################################################################

'Public Sub mnuPopupComm_Click()
'Form_DblClick
'End Sub
'
'Public Sub mnuPopupHide_Click()
'Unload Me
'End Sub

Public Sub mnuPopupReset_Click()
Me.Left = Screen.width - Me.width
Me.Top = Screen.height - Me.height - GetTaskbarHeight()
End Sub

'###################################################################################################

Private Sub ShowLinkSelect()
SetCursor LoadCursor(ByVal 0&, IDC_HAND)
End Sub

Private Sub picClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bIgnoreLostFocus = True
'EnableTracking

Form_apiGotFocus
ShowLinkSelect
End Sub

Private Sub picClose_Click()
Unload Me
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
    apiShowInTaskbar = False
End If
End Sub

Private Property Let apiShowInTaskbar(ByVal bShow As Boolean)
Dim l As Long, hWnd As Long

hWnd = Me.hWnd

l = GetWindowLong(hWnd, GWL_EXSTYLE)

If bShow Then
    l = l Or WS_EX_APPWINDOW
Else
    l = l And Not WS_EX_APPWINDOW
End If


Me.Visible = False
SetWindowLong hWnd, GWL_EXSTYLE, l
Me.Visible = True
End Property
