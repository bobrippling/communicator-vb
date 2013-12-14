Attribute VB_Name = "modMenu"
Option Explicit

Public Const Menu_Comm_Colour = &H8000000F
Public Const Menu_Default_Colour = &H8000000F

Private Const MIM_BACKGROUND As Long = &H2
Private Const MIM_APPLYTOSUBMENUS As Long = &H80000000

Private Type MENUINFO
   cbSize As Long
   fMask As Long
   dwStyle As Long
   cyMax As Long
   hbrBack As Long
   dwContextHelpID As Long
   dwMenuData As Long
End Type

Private Declare Function DrawMenuBar Lib "user32" _
  (ByVal hWnd As Long) As Long

Private Declare Function GetMenu Lib "user32" _
  (ByVal hWnd As Long) As Long

Private Declare Function GetSystemMenu Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal bRevert As Long) As Long

Private Declare Function SetMenuInfo Lib "user32" _
  (ByVal hmenu As Long, _
   mi As MENUINFO) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" _
  (ByVal crColor As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32.dll" _
  (ByVal OLE_COLOR As Long, _
   ByVal HPALETTE As Long, _
   pccolorref As Long) As Long

Public Function SetMenuColour(ByVal hWnd As Long, _
                             ByVal dwColour As Long, _
                             ByVal bIncludeSubmenus As Boolean) As Boolean

'set application menu colour
Dim mi As MENUINFO
Dim flags As Long
Dim clrref As Long

'convert a Windows colour (OLE colour)
'to a valid RGB colour if required
clrref = TranslateOLEtoRBG(dwColour)

flags = MIM_BACKGROUND
If bIncludeSubmenus Then
    flags = flags Or MIM_APPLYTOSUBMENUS
End If

With mi
    .cbSize = Len(mi)
    .fMask = flags
    .hbrBack = CreateSolidBrush(clrref)
End With

SetMenuInfo GetMenu(hWnd), mi
DrawMenuBar hWnd

End Function


'Private Function SetSysMenuColour(ByVal hwndfrm As Long, _
'                               ByVal dwColour As Long) As Boolean
'
''set system menu colour
'Dim mi As MENUINFO
'Dim hSysMenu As Long
'Dim clrref As Long
'
''convert a Windows colour (OLE colour)
''to a valid RGB colour if required
'clrref = TranslateOLEtoRBG(dwColour)
'
''get handle to the system menu,
''fill in struct, assign to menu,
''and force a redraw with the
''new attributes
'hSysMenu = GetSystemMenu(Me.hWnd, False)
'
'With mi
'    .cbSize = Len(mi)
'    .fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
'    .hbrBack = CreateSolidBrush(clrref)
'End With
'
'SetMenuInfo hSysMenu, mi
'DrawMenuBar hSysMenu
'
'End Function


Private Function TranslateOLEtoRBG(ByVal dwOleColour As Long) As Long
  
'check to see if the passed colour
'value is and OLE or RGB colour, and
'if an OLE colour, translate it to
'a valid RGB color and return. If the
'colour is already a valid RGB colour,
'the function returns the colour without
'change
OleTranslateColor dwOleColour, 0, TranslateOLEtoRBG

End Function

