VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ucListEdit 
   ClientHeight    =   3825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   ScaleHeight     =   3825
   ScaleWidth      =   4740
   Begin VB.TextBox txtEdit 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListViewMain 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "ucListEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Private bDoingSetup As Boolean
Private dirty As Boolean, bFirstColumn As Boolean
Private itmClicked As ListItem
Private dwLastSubitemEdited As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
Private Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
Private Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
Private Const LVHT_ONITEMICON = &H2
Private Const LVHT_ONITEMLABEL = &H4
Private Const LVHT_ONITEMSTATEICON = &H8
Private Const LVHT_ONITEM = (LVHT_ONITEMICON Or _
    LVHT_ONITEMLABEL Or _
    LVHT_ONITEMSTATEICON)
Private Const LVIR_LABEL = 2

Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LVHITTESTINFO
    pt As PointAPI
    flags As Long
    iItem As Long
    iSubItem As Long
End Type

Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long

Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
    
Public listViewObject As ListView

Private pColumnOfIntrest() As Integer
Private pColumnOfIntrest_Count As Integer
Private CurrentColumn As Integer

Public Event AfterLabelEdit(ByRef Text As String, ByVal RowNo As Long, ByVal ColumnNo As Integer)
Public Event ColumnClick(ByVal colIndex As Integer)

Public Sub addColumnOfIntrest(ColumnNo As Integer)
Dim i As Integer

' CAN'T EDIT 0th COLUMN...?
If 0 <= ColumnNo And ColumnNo < ListViewMain.ColumnHeaders.Count Then
    
    For i = 0 To pColumnOfIntrest_Count - 1
        If pColumnOfIntrest(i) = ColumnNo Then Exit Sub
    Next i
    
    ReDim Preserve pColumnOfIntrest(pColumnOfIntrest_Count)
    pColumnOfIntrest(pColumnOfIntrest_Count) = ColumnNo
    pColumnOfIntrest_Count = pColumnOfIntrest_Count + 1
    
Else
    Err.Raise 1, "EditableListView", "Cannot Set Column of Intrest to Undimensioned Column"
End If

End Sub

Private Sub listviewmain_ColumnClick(ByVal ColumnHeader As ColumnHeader)

'hide the text box
txtEdit.Visible = False

'sort the items
ListViewMain.SortKey = ColumnHeader.Index - 1
ListViewMain.SortOrder = Abs(Not ListViewMain.SortOrder = 1)
ListViewMain.Sorted = True

RaiseEvent ColumnClick(ColumnHeader.Index)

End Sub

Private Sub listviewmain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'this routine:
'1. sets the last change if the dirty flag is set
'2. sets a flag to prevent setting the dirty flag
'3. determines the item or subitem clicked
'4. calc's the position for the text box
'5. moves and shows the text box
'6. clears the dirty flag
'7. clears the DoingSetup flag

If pColumnOfIntrest_Count = 0 Then Exit Sub


Dim HTI As LVHITTESTINFO
Dim fpx As Single
Dim fpy As Single
Dim fpw As Single
Dim fph As Single
Dim rc As RECT
Dim topindex As Long
Dim i As Integer

Dim ColumnOfIntrestClicked As Boolean

'prevent the textbox change event from
'registering as dirty when the text is
'assigned to the textbox
bDoingSetup = True

'if a pending dirty flag is set, update the
'last edited item before moving on
If dirty Then
    If bFirstColumn Then
        itmClicked.Text = txtEdit.Text
    Else
        itmClicked.SubItems(dwLastSubitemEdited) = txtEdit.Text
    End If
    RaiseTheEvent
End If

'hide the textbox
txtEdit.Visible = False

'get the position of the click
With HTI
    .pt.X = (X / Screen.TwipsPerPixelX)
    .pt.Y = (Y / Screen.TwipsPerPixelY)
    .flags = LVHT_ONITEM
End With

'find out which subitem was clicked
Call SendMessage(ListViewMain.hWnd, _
                LVM_SUBITEMHITTEST, _
                0, HTI)

'if on an item (HTI.iItem <> -1) and
'the click occurred on the subitem
'column of interest (HTI.iSubItem = 2 -
'which is column 3 (0-based)) move and
'show the textbox

For i = 0 To pColumnOfIntrest_Count - 1
    If HTI.iSubItem = pColumnOfIntrest(i) Then
        ColumnOfIntrestClicked = True
        Exit For
    End If
Next i

If HTI.iItem <> -1 And ColumnOfIntrestClicked Then 'column of intrest
    
    'prevent the listview label editing
    'from occurring if the control has
    'full row select set
    ListViewMain.LabelEdit = lvwManual
    
    CurrentColumn = HTI.iItem
    
    'determine the bounding rectangle
    'of the subitem column
    rc.Left = LVIR_LABEL
    rc.Top = HTI.iSubItem
    
    Call SendMessage(ListViewMain.hWnd, _
                        LVM_GETSUBITEMRECT, _
                        HTI.iItem, _
                        rc)
    
    
    'If UserControl.ListViewMain.Checkboxes Then Exit Sub
    
    
    'we need to keep track of which
    'item was clicked so the item can
    'be updated later
    'position the text box
    Set itmClicked = ListViewMain.ListItems(HTI.iItem + 1)
    
    itmClicked.Selected = True
    
    'get the current top index
    topindex = SendMessage(ListViewMain.hWnd, _
                    LVM_GETTOPINDEX, _
                    0&, _
                    ByVal 0&)
    
    'establish the bounding rect for
    'the subitem in VB terms (the x
    'and y coordinates, and the height
    'and width of the item
    fpx = ListViewMain.Left + (rc.Left * Screen.TwipsPerPixelX) + 80
    
    
    fpy = ListViewMain.Top + (HTI.iItem + 1 - topindex) + (rc.Top * Screen.TwipsPerPixelY)
    
    'a hard-coded height for the text box
    fph = 120
    
    'get the column width for the subitem
    fpw = SendMessage(ListViewMain.hWnd, _
                LVM_GETCOLUMNWIDTH, _
                HTI.iSubItem, _
                ByVal 0&)
    
    'calc the required width of
    'the textbox to fit in the column
    fpw = (fpw * Screen.TwipsPerPixelX) - 40
    
    
    'assign the current subitem
    'value to the textbox
    With txtEdit
        If HTI.iSubItem = 0 Then
            'it's the first column
            .Text = itmClicked.Text
            bFirstColumn = True
        Else
            .Text = itmClicked.SubItems(HTI.iSubItem)
            bFirstColumn = False
        End If
        
        dwLastSubitemEdited = HTI.iSubItem
        
        'position it over the subitem, make
        'visible and assure the text box
        'appears overtop the listview
        .Move fpx, fpy, fpw, fph
        .Visible = True
        .ZOrder 0
        .SetFocus
    End With
    
    'clear the setup flag to allow the
    'textbox change event to set the
    '"dirty" flag, and clear that flag
    'in preparation for editing
    bDoingSetup = False
    dirty = False

End If

End Sub


Private Sub listviewmain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'if showing the text box, set
'focus to it and select any
'text in the control
If txtEdit.Visible Then
    
    With txtEdit
        .SetFocus
        .Selstart = 0
        .Sellength = Len(.Text)
    End With
    
End If

End Sub

Private Sub txtEdit_Change()

If Not bDoingSetup Then
    dirty = True
End If

End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    'txtEdit_LostFocus
    UserControl.ListViewMain.SetFocus 'force defocusing of txtEdit
    txtEdit.Visible = False
    KeyAscii = 0
    
'ElseIf (Not IsNumeric(Chr$(KeyAscii))) And (KeyAscii <> 8) Then
'    If CurrentColumn = 0 Then KeyAscii = 0
End If
End Sub

Private Sub txtEdit_LostFocus()

If dirty And dwLastSubitemEdited > 0 Then
    itmClicked.SubItems(dwLastSubitemEdited) = txtEdit.Text
    dirty = False
    RaiseTheEvent
End If

End Sub

Private Sub UserControl_Initialize()
With ListViewMain
    
    .SortKey = 0
    .Sorted = False
    .View = lvwReport
    .GridLines = True
    .FullRowSelect = True
    .LabelEdit = lvwManual

End With

Set listViewObject = ListViewMain

End Sub

Public Sub UserControl_Resize()

Dim w As Integer, i As Integer

ListViewMain.width = ScaleWidth
ListViewMain.height = ScaleHeight

If ListViewMain.ColumnHeaders.Count Then
    w = ListViewMain.width / ListViewMain.ColumnHeaders.Count - 75
    
    For i = 1 To ListViewMain.ColumnHeaders.Count
        ListViewMain.ColumnHeaders(i).width = w
    Next i
End If

End Sub

Public Sub HideEditBox()
UserControl.txtEdit.Visible = False
End Sub

Private Sub RaiseTheEvent()
Dim Txt As String

Txt = txtEdit.Text

RaiseEvent AfterLabelEdit(Txt, CurrentColumn + 1, dwLastSubitemEdited)

txtEdit.Text = Txt

End Sub
