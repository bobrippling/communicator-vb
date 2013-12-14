VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl smRtfFBox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   6450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   ScaleHeight     =   6450
   ScaleWidth      =   6495
   Begin VB.CommandButton CMDShowFaces 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   6240
      Picture         =   "smRtfControl.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click here to see all emoticons"
      Top             =   4080
      Visible         =   0   'False
      Width           =   165
   End
   Begin RichTextLib.RichTextBox RTFBuff 
      Height          =   1995
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   3519
      _Version        =   393217
      FileName        =   "F:\SOURCE\My Programming Source Code Files\my files\Project smiles\smilecodes.rtf"
      TextRTF         =   $"smRtfControl.ctx":0371
   End
   Begin RichTextLib.RichTextBox RTFSmile_Faces 
      Height          =   2355
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4154
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      MousePointer    =   1
      Appearance      =   0
      FileName        =   "F:\SOURCE\My Programming Source Code Files\my files\Project smiles\simleFaces.rtf"
      TextRTF         =   $"smRtfControl.ctx":0E9C
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   11245
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"smRtfControl.ctx":41B93
   End
End
Attribute VB_Name = "smRtfFBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Default Property Values:
Private Const m_def_Selstart = 0
Private Const m_def_Sellength = 0
Private Const m_def_SelColor = 0
Private Const m_def_SelText = vbNullString
Private Const m_def_SelRtf = "0"
Private Const m_def_ForeColor = 0
Private Const m_def_BackStyle = 0
Private Const m_def_FillColor = 0
Private Const m_def_hDC = 0
Private Const m_def_hWnd = 0
'Property Variables:
Private m_ForeColor As Long
Private m_BackStyle As Integer
Private m_FillColor As Long
Private m_hDC As Long
Private m_hWnd As Long
Private Enable_Smiles As Boolean    'Custom
Private Text_Filter As Boolean      'Custom
Private Filter_Path As String       'Custom
'Event Declarations:
Public Event Click() 'MappingInfo=RTFTEXT,RTFTEXT,-1,Click
Public Event DblClick() 'MappingInfo=RTFTEXT,RTFTEXT,-1,DblClick
Public Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=RTFTEXT,RTFTEXT,-1,KeyDown
Public Event KeyPress(KeyAscii As Integer) 'MappingInfo=RTFTEXT,RTFTEXT,-1,KeyPress
Public Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=RTFTEXT,RTFTEXT,-1,KeyUp
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=RTFTEXT,RTFTEXT,-1,MouseDown
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=RTFTEXT,RTFTEXT,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=RTFTEXT,RTFTEXT,-1,MouseUp
Public Event SelChange() 'MappingInfo=RTFTEXT,RTFTEXT,-1,SelChange
Public Event SmileSelected(ByVal Smile_code As String)
'Variables started
Private SM_Codes(1 To 95) As String
Private Filters() As String

'--------------------------------
'end smiley bit

'This module provides RichTextBox features that are not
'natively exposed by the RichTextBox control but a RichText
'supports.  Currently, this is only for auto-detecting
'URLs (which formats a URL as a hyerplink) and clicking
'the URL to launch a default web browser, email program, etc.

Private Const WM_USER                   As Long = &H400
Private Const EM_GETAUTOURLDETECT       As Long = (WM_USER + 92)
Private Const EM_AUTOURLDETECT          As Long = (WM_USER + 91)
Private Const EM_SETEVENTMASK           As Long = (WM_USER + 69)
Private Const EM_GETEVENTMASK           As Long = (WM_USER + 59)
Private Const ENM_LINK                  As Long = &H4000000

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Private Const GWL_WNDPROC           As Long = (-4)

Public Function DisableURLDetection(ByVal RTBhwnd As Long) As Boolean

Dim lEventMask As Long

'Need to get current event mask
lEventMask = SendMessage(RTBhwnd, EM_GETEVENTMASK, 0&, ByVal 0&)

'Remove the ENM_LINK mask
lEventMask = lEventMask And Not ENM_LINK

'Now set the new event mask
Call SendMessage(RTBhwnd, EM_SETEVENTMASK, 0&, ByVal lEventMask)
Call SendMessage(RTBhwnd, EM_AUTOURLDETECT, 0&, ByVal 0&)

m_lRTBhWnd = 0

DisableURLDetection = True

DisableURLHook UserControl.hWnd

End Function

Public Function EnableURLDetection(ByVal RTBhwnd As Long) As Boolean

Dim lEventMask As Long

Call SendMessage(RTBhwnd, EM_AUTOURLDETECT, 1&, ByVal 0&)

'Need to get current event mask
lEventMask = SendMessage(RTBhwnd, EM_GETEVENTMASK, 0&, ByVal 0&)

'Add the ENM_LINK mask
lEventMask = lEventMask Or ENM_LINK

'Now set the new event mask
Call SendMessage(RTBhwnd, EM_SETEVENTMASK, 0&, ByVal lEventMask)

m_lRTBhWnd = RTBhwnd

EnableURLDetection = True

EnableURLHook UserControl.hWnd

End Function

Private Function EnableURLHook(ByVal hWnd As Long) As Boolean

'This function enables subclassing

'We must already have the RichTextBox's window handle.
'This is set by calling EnableURLDetection.

If m_lRTBhWnd = 0 Then
    EnableURLHook = False
Else
    'Get the address for the previous window procedure
    lpfnOldWinProc = GetWindowLong(hWnd, GWL_WNDPROC)
    If lpfnOldWinProc = 0 Then
        'If the return value is 0, the function failed
        EnableURLHook = False
    Else
        'The return value of SetWindowLong is the address of the previous procedure,
        'so if it's not what we just got above, something went wrong.
        If SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc) <> lpfnOldWinProc Then
            EnableURLHook = False
        Else
            EnableURLHook = True
        End If
    End If
End If

End Function

Public Function DisableURLHook(Optional ByVal hWnd As Long = -1) As Boolean

If hWnd = -1 Then hWnd = UserControl.hWnd

'Restore default window procedure
If SetWindowLong(hWnd, GWL_WNDPROC, lpfnOldWinProc) = 0 Then
    DisableURLHook = False
Else
    DisableURLHook = True
    lpfnOldWinProc = 0
End If

End Function


'------------------------------
'start smiley bit

Public Property Get SelUnderline() As Boolean
rtfText.Sellength = 1
SelUnderline = rtfText.SelUnderline
End Property

Public Property Let FontName(ByVal F As String)
'Dim RtfTxt As String

'RtfTxt = rtfText.TextRTF 'preserve the colour etc
rtfText.Font.Name = F
'rtfText.TextRTF = RtfTxt
End Property

Public Property Get FontName() As String
FontName = rtfText.Font.Name
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
BackColor = rtfText.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
rtfText.BackColor() = New_BackColor
PropertyChanged "BackColor"
End Property

Public Property Get SelText() As String
SelText = rtfText.SelText
End Property

Public Property Let SelText(ByVal Txt As String)
rtfText.SelText = Txt
'RTFBuff.SelText = Txt
If Enable_Smiles Then RefreshAll 'Process_Smiles
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
m_ForeColor = New_ForeColor
PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,Enabled
Public Property Get Enabled() As Boolean
Enabled = rtfText.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
rtfText.Enabled() = New_Enabled
PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,Font
Public Property Get Font() As Font
Set Font = rtfText.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Set rtfText.Font = New_Font
PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
m_BackStyle = New_BackStyle
PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
BorderStyle = rtfText.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
rtfText.BorderStyle() = New_BorderStyle
PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,Refresh
Public Sub Refresh()
rtfText.Refresh
End Sub

Public Sub CMDShowFaces_Click()
If RTFSmile_Faces.Visible = False Then
    RTFSmile_Faces.Visible = True
    On Error Resume Next
    RTFSmile_Faces.SetFocus
Else
    RTFSmile_Faces.Visible = False
End If
End Sub

Private Sub RTFSmile_Faces_Click()
On Error Resume Next
RaiseEvent SmileSelected(SM_Codes(RTFSmile_Faces.Selstart + 1))
RTFSmile_Faces.Visible = False
End Sub

Private Sub RTFSmile_Faces_LostFocus()
RTFSmile_Faces.Visible = False
End Sub

Private Sub RTFTEXT_Click()
RaiseEvent Click
End Sub

Private Sub RTFTEXT_DblClick()
RaiseEvent DblClick
End Sub

Private Sub RTFText_GotFocus()
RTFSmile_Faces.Visible = False
End Sub

Private Sub RTFTEXT_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub RTFTEXT_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub RTFTEXT_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub RTFTEXT_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub RTFTEXT_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub RTFTEXT_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Appearance = rtfText.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
rtfText.Appearance() = New_Appearance
PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,BulletIndent
Public Property Get BulletIndent() As Single
BulletIndent = rtfText.BulletIndent
End Property

Public Property Let BulletIndent(ByVal New_BulletIndent As Single)
rtfText.BulletIndent() = New_BulletIndent
PropertyChanged "BulletIndent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Cls()

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,FileName
Public Property Get FileName() As String
FileName = rtfText.FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
rtfText.FileName() = New_FileName
PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get FillColor() As Long
FillColor = m_FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As Long)
m_FillColor = New_FillColor
PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,Find
Public Function Find(ByVal bstrString As String, Optional ByVal vStart As Variant, Optional ByVal vEnd As Variant, Optional ByVal vOptions As Variant) As Long
Find = rtfText.Find(bstrString, vStart, vEnd, vOptions)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get hDC() As Long
hDC = m_hDC
End Property

Public Property Let hDC(ByVal New_hDC As Long)
m_hDC = New_hDC
PropertyChanged "hDC"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,HideSelection
Public Property Get HideSelection() As Boolean
HideSelection = rtfText.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
rtfText.HideSelection() = New_HideSelection
PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get hWnd() As Long
hWnd = rtfText.hWnd ' m_hWnd
End Property

Public Property Let hWnd(ByVal New_hWnd As Long)
m_hWnd = New_hWnd
PropertyChanged "hWnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,LoadFile
Public Sub LoadFile(ByVal bstrFilename As String, Optional ByVal vFileType As Variant)
rtfText.LoadFile bstrFilename, vFileType
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,Locked
Public Property Get Locked() As Boolean
Locked = rtfText.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
rtfText.Locked() = New_Locked
PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,MaxLength
Public Property Get MaxLength() As Long
MaxLength = rtfText.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
rtfText.MaxLength() = New_MaxLength
PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Set MouseIcon = rtfText.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
Set rtfText.MouseIcon = New_MouseIcon
PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
MousePointer = rtfText.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
rtfText.MousePointer() = New_MousePointer
PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,MultiLine
Public Property Get MultiLine() As Boolean
MultiLine = rtfText.MultiLine
End Property

Public Property Let MultiLine(ByVal New_MultiLine As Boolean)
rtfText.MultiLine() = New_MultiLine
PropertyChanged "MultiLine"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,SaveFile
Public Sub SaveFile(ByVal bstrFilename As String, Optional ByVal vFlags As Variant)
rtfText.SaveFile bstrFilename, vFlags
End Sub

Private Sub RTFTEXT_SelChange()
RaiseEvent SelChange
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=RTFTEXT,RTFTEXT,-1,Text
Public Property Get Text() As String
Text = rtfText.Text
End Property

Public Property Let Text(ByVal New_Text As String)
rtfText.Text() = New_Text
PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13
Public Function AddLINEToEnd(RTF_Text As String, Optional RTF_Color As Long) As String
On Error Resume Next
If RTF_Text = "" Then Exit Function 'If text is empty then exit function
RTFBuff.Text = ""
RTFBuff.SelColor = RTF_Color
RTFBuff.SelText = RTF_Text
rtfText.Selstart = Len(rtfText.Text)
'If Smiles are Enabled Then Show them
If Enable_Smiles = True Then Process_Smiles
If EnableTextFilter = True Then Apply_Filter
rtfText.SelRtf = RTFBuff.TextRTF
End Function

Private Sub UserControl_Initialize()
m_hWnd = rtfText.hWnd
Process_SmileCodes
Process_Filter
RTFSmile_Faces.Visible = False
Me.EnableSmiles = False
End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
m_ForeColor = m_def_ForeColor
m_BackStyle = m_def_BackStyle
m_FillColor = m_def_FillColor
m_hDC = m_def_hDC
m_hWnd = m_def_hWnd
Control_Size
Enable_Smiles = True
Text_Filter = False
Filter_Path = ""
rtfText.Selstart = m_def_Selstart
rtfText.Sellength = m_def_Sellength
rtfText.SelColor = m_def_SelColor
rtfText.SelRtf = m_def_SelRtf
rtfText.SelText = m_def_SelText
End Sub

Private Sub UserControl_LostFocus()
RTFSmile_Faces.Visible = False
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

rtfText.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
rtfText.Enabled = PropBag.ReadProperty("Enabled", True)
Set rtfText.Font = PropBag.ReadProperty("Font", Ambient.Font)
m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
rtfText.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
'RTFTEXT.Appearance = PropBag.ReadProperty("Appearance", 1)
rtfText.BulletIndent = PropBag.ReadProperty("BulletIndent", 0)
rtfText.FileName = PropBag.ReadProperty("FileName", "")
m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
m_hDC = PropBag.ReadProperty("hDC", m_def_hDC)
rtfText.HideSelection = PropBag.ReadProperty("HideSelection", True)
m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
rtfText.Locked = PropBag.ReadProperty("Locked", False)
rtfText.MaxLength = PropBag.ReadProperty("MaxLength", 0)
Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
rtfText.MousePointer = PropBag.ReadProperty("MousePointer", 0)
'RTFText.MultiLine = PropBag.ReadProperty("MultiLine", True)
rtfText.Text = PropBag.ReadProperty("Text", "")
EnableSmiles = PropBag.ReadProperty("EnableSmiles", True)
EnableTextFilter = PropBag.ReadProperty("EnableTextFilter", True)
FilterFile = PropBag.ReadProperty("FilterFile", "")
rtfText.Selstart = PropBag.ReadProperty("Selstart", m_def_Selstart)
rtfText.Sellength = PropBag.ReadProperty("Sellength", m_def_Sellength)
rtfText.SelColor = PropBag.ReadProperty("SelColor", m_def_SelColor)
rtfText.SelRtf = PropBag.ReadProperty("SelRtf", m_def_SelRtf)
rtfText.SelText = PropBag.ReadProperty("SelText", m_def_SelText)

End Sub

Private Sub UserControl_Resize()
Control_Size
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Call PropBag.WriteProperty("BackColor", rtfText.BackColor, &H80000005)
Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
Call PropBag.WriteProperty("Enabled", rtfText.Enabled, True)
Call PropBag.WriteProperty("Font", rtfText.Font, Ambient.Font)
Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
Call PropBag.WriteProperty("BorderStyle", rtfText.BorderStyle, 1)
Call PropBag.WriteProperty("Appearance", rtfText.Appearance, 1)
Call PropBag.WriteProperty("BulletIndent", rtfText.BulletIndent, 0)
Call PropBag.WriteProperty("FileName", rtfText.FileName, "")
Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
Call PropBag.WriteProperty("hDC", m_hDC, m_def_hDC)
Call PropBag.WriteProperty("HideSelection", rtfText.HideSelection, True)
Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
Call PropBag.WriteProperty("Locked", rtfText.Locked, False)
Call PropBag.WriteProperty("MaxLength", rtfText.MaxLength, 0)
Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
Call PropBag.WriteProperty("MousePointer", rtfText.MousePointer, 0)
Call PropBag.WriteProperty("MultiLine", rtfText.MultiLine, True)
Call PropBag.WriteProperty("Text", rtfText.Text, "")
Call PropBag.WriteProperty("EnableSmiles", Enable_Smiles, True)
Call PropBag.WriteProperty("EnableTextFilter", Text_Filter, False)
Call PropBag.WriteProperty("FilterFile", Filter_Path, "")
Call PropBag.WriteProperty("Selstart", rtfText.Selstart, m_def_Selstart)
Call PropBag.WriteProperty("Sellength", rtfText.Sellength, m_def_Sellength)
Call PropBag.WriteProperty("SelColor", rtfText.SelColor, m_def_SelColor)
Call PropBag.WriteProperty("SelRtf", rtfText.SelRtf, m_def_SelRtf)
Call PropBag.WriteProperty("SelText", rtfText.SelText, m_def_SelText)

End Sub
Private Function Process_SmileCodes()
Dim y As Integer
Dim Result As Long
Dim Sm_Data As String

'Fill the Smile codes Varibales
Sm_Data = RTFBuff.Text

For y = 1 To 95
    Result = InStr(1, Sm_Data, vbNewLine)
    If Result Then
        SM_Codes(y) = Left$(Sm_Data, (Result - 1))
        Sm_Data = Right$(Sm_Data, (Len(Sm_Data) - (Result + 1)))
    End If
Next y

RTFBuff.Text = vbNullString
End Function

Private Function Process_Smiles()
Dim y As Integer

For y = 1 To 95
    Do
        If (rtfText.Find(SM_Codes(y))) = "-1" Then
            Exit Do
        Else
            RTFSmile_Faces.Selstart = y - 1
            RTFSmile_Faces.Sellength = 1
            rtfText.SelRtf = RTFSmile_Faces.SelRtf
        End If
        RTFSmile_Faces.Sellength = 0
    Loop
Next y
End Function

Public Function RefreshAll()
Dim y As Integer

If Enable_Smiles Then Process_Smiles


If EnableTextFilter Then 'case when enable smiles is true
    If Filter_Path = "" Then Exit Function 'Err.Raise "11", , "Filter File not specified."
    For y = 0 To UBound(Filters)
        Do
            If Filters(y) = "" Then Exit Do
            If (rtfText.Find(Filters(y))) = "-1" Then
                Exit Do
            Else
                rtfText.SelRtf = "***"
            End If
        Loop
    Next y
End If


rtfText.Selstart = Len(rtfText.Text)

End Function

Public Property Get EnableSmiles() As Boolean
EnableSmiles = Enable_Smiles
End Property

Public Property Let EnableSmiles(ByVal New_Enabled As Boolean)
Enable_Smiles = New_Enabled
PropertyChanged "EnableSmiles"
End Property

Public Property Get EnableTextFilter() As Boolean
EnableTextFilter = Text_Filter
End Property

Public Property Let EnableTextFilter(ByVal New_Enabled As Boolean)
Text_Filter = New_Enabled
PropertyChanged "EnableTextFilter"
End Property
Public Property Get FilterFile() As String
FilterFile = Filter_Path
End Property
Public Property Let FilterFile(ByVal New_Enabled As String)
Filter_Path = New_Enabled
PropertyChanged "FilterFile"
Process_Filter
End Property
Public Function Process_Filter()
On Error GoTo errhandler:
Dim y As Integer
Dim Result As Long
Dim FL_Data As String
'Fill the Words to be filtered
If Filter_Path = vbNullString Then Exit Function

RTFBuff.LoadFile Filter_Path
FL_Data = RTFBuff.Text
RTFBuff.Text = ""
ReDim Filters(0)
Do
    Result = InStr(1, FL_Data, vbCrLf)
    If Result = "0" Then
        Exit Do
    Else
        ReDim Preserve Filters(UBound(Filters) + 1)
        Filters(y) = Left$(FL_Data, (Result - 1))
        FL_Data = Right(FL_Data, (Len(FL_Data) - (Result + 1)))
        y = y + 1
    End If
Loop
errhandler:
If Not Err.Number = 0 Then
    Err.Raise Err.Number, , "Filter Filed Not Found."
End If
End Function
Private Function Apply_Filter()
Dim y As Integer
If Filter_Path = "" Then Err.Raise "11", , "Filter File not specified."
For y = 0 To UBound(Filters)
    Do
        If Filters(y) = "" Then Exit Do
        If (RTFBuff.Find(Filters(y))) = "-1" Then
            Exit Do
        Else
            RTFBuff.SelRtf = "***"
        End If
    Loop
Next y
End Function

Private Function Control_Size()
'Dim b As Boolean

rtfText.width = UserControl.width
rtfText.height = UserControl.height


'If b Then
    'CMDShowFaces.Left = 30 'UserControl.Width - CMDShowFaces.Width
    'CMDShowFaces.Top = UserControl.Height - CMDShowFaces.Height - 10
'End If

RTFSmile_Faces.Top = UserControl.height - 10 - RTFSmile_Faces.height
RTFSmile_Faces.Left = rtfText.width - RTFSmile_Faces.width
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,0
Public Property Get Selstart() As Long
Selstart = rtfText.Selstart
End Property

Public Property Let Selstart(ByVal New_Selstart As Long)
If Ambient.UserMode = False Then Err.Raise 387
rtfText.Selstart = New_Selstart
PropertyChanged "Selstart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,0
Public Property Get Sellength() As Long
Sellength = rtfText.Sellength
End Property

Public Property Let Sellength(ByVal New_Sellength As Long)
If Ambient.UserMode = False Then Err.Raise 387
rtfText.Sellength = New_Sellength
PropertyChanged "Sellength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,2,0
Public Property Get SelColor() As Long
SelColor = rtfText.SelColor
End Property

Public Property Let SelColor(ByVal New_SelColor As Long)
If Ambient.UserMode = False Then Err.Raise 387
rtfText.SelColor = New_SelColor
PropertyChanged "SelColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,0
Public Property Get SelRtf() As String
SelRtf = rtfText.SelRtf
End Property

Public Property Let SelRtf(ByVal New_SelRtf As String)
If Ambient.UserMode = False Then Err.Raise 387
rtfText.SelRtf = New_SelRtf
PropertyChanged "SelRtf"
End Property
'-----------------------
Public Property Get SelFontName() As String
SelFontName = rtfText.SelFontName
End Property

Public Property Let SelFontName(F As String)
rtfText.SelFontName = F
End Property
'-----------------------
Public Property Get SelFontSize() As Single
SelFontSize = rtfText.SelFontSize
End Property

Public Property Let SelFontSize(fS As Single)
rtfText.SelFontSize = fS
End Property
'-----------------------
Public Property Get SelBold() As Boolean
SelBold = rtfText.SelBold
End Property

Public Property Let SelBold(fB As Boolean)
rtfText.SelBold = fB
End Property
'-----------------------
Public Property Get SelItalic() As Boolean
SelItalic = rtfText.SelItalic
End Property

Public Property Let SelItalic(fI As Boolean)
rtfText.SelItalic = fI
End Property
