VERSION 5.00
Begin VB.UserControl ScrollListBox 
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   ScaleHeight     =   2340
   ScaleWidth      =   3450
   Begin VB.ListBox lst 
      Height          =   1230
      ItemData        =   "ScrollListBox.ctx":0000
      Left            =   0
      List            =   "ScrollListBox.ctx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "ScrollListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const Extra As Integer = 5

Public Event Click()
Public Event DblClick()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private pRaiseClickOnMouseDown As Boolean

Public Property Get hWnd() As Long
hWnd = lst.hWnd
End Property

Public Property Let RaiseClickOnMouseDown(ByVal b As Boolean)
pRaiseClickOnMouseDown = b
End Property

Public Property Get RaiseClickOnMouseDown() As Boolean
RaiseClickOnMouseDown = pRaiseClickOnMouseDown
End Property

Private Sub lst_Click()
RaiseEvent Click
End Sub

Private Sub lst_DblClick()
RaiseEvent DblClick
End Sub

Private Sub lst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

RaiseEvent MouseDown(Button, Shift, X, Y)

If pRaiseClickOnMouseDown Then
    lst_Click
End If

End Sub

Private Sub lst_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
pRaiseClickOnMouseDown = True
End Sub

Private Sub UserControl_Resize()
lst.Top = 0
lst.Left = 0
lst.width = UserControl.ScaleWidth
lst.height = UserControl.ScaleHeight
End Sub

Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Integer)

Dim i As Integer, Highest As Integer, TW As Integer

lst.AddItem Item, Index

For i = 0 To lst.ListCount - 1
    TW = TextWidth(lst.List(i))
    
    If TW > Highest Then
        Highest = TW
    End If
Next i

If ScaleMode = vbTwips Then
    Highest = Highest / Screen.TwipsPerPixelX ' if twips change to pixels
End If

If Highest Then Highest = Highest + Extra

SendMessageByLong lst.hWnd, LB_SETHORIZONTALEXTENT, Highest, 0

End Sub

Public Property Get ListCount() As Integer
ListCount = lst.ListCount
End Property

Public Sub Clear()
lst.Clear
SendMessageByLong lst.hWnd, LB_SETHORIZONTALEXTENT, 0, 0
End Sub

Public Property Get ListIndex() As Integer
ListIndex = lst.ListIndex
End Property

Public Property Let ListIndex(ByVal i As Integer)
lst.ListIndex = i
End Property

Public Sub LetItemData(iData As Integer, ListIndex As Integer)
lst.ItemData(ListIndex) = iData
End Sub

Public Function GetItemData(ListIndex As Integer) As Integer
GetItemData = lst.ItemData(ListIndex)
End Function

Public Property Get List(ByVal i As Integer) As String
List = lst.List(i)
End Property

Public Property Get Text() As String
Text = lst.Text
End Property

Public Property Let Enabled(ByVal b As Boolean)
lst.Enabled = b
End Property

Public Sub RemoveItem(ByVal i As Integer)
lst.RemoveItem i
End Sub
