Attribute VB_Name = "modGDI"
Option Explicit

'Public Type POINTAPI
'    x As Long
'    y As Long
'End Type

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, _
    ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, _
    lpPoint As Any, ByVal nCount As Long) As Long

Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, _
    ByVal hRgn As Long, ByVal hBrush As Long) As Long

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const ALTERNATE = 1 ' ALTERNATE and WINDING are
Private Const WINDING = 2 ' constants for FillMode

'Private Const BLACKBRUSH = 4 ' Constant for brush type.

'pen selection
Private Const PS_SOLID = 0
Private Const PS_DOT = 2
'Private Const PS_NULL = 5

Private Const NULL_PEN = 8

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
    ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" ( _
    ByVal hdc As Long, ByVal hObject As Long) As Long


'hatch
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
'Private Declare Function PaintRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
'Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
'Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Const HS_DIAGCROSS = 5, HS_CROSS = 4

Private Declare Function CreateRoundRectRgn Lib "gdi32" ( _
    ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, _
    ByVal nBottomRect As Long, ByVal nWidthEllipse As Long, ByVal nHeightEllipse As Long) As Long


Public Sub DrawPoly(Poly() As POINTAPI, hdc As Long, FillCol As Long) ', Optional iMode As Long = ALTERNATE)
Dim hBrush As Long, hOldPen As Long
Dim hRgn As Long
Dim LB As Integer, UB As Integer

LB = LBound(Poly)
UB = UBound(Poly)


'draw it
Polygon hdc, Poly(LB), UB - LB + 1

If FillCol > -1 Then
    'make a new brush
    hBrush = CreatePen(PS_SOLID, 1, FillCol)
    hOldPen = SelectObject(hdc, hBrush)
    
    
    'Creates region to fill with colour
    hRgn = CreatePolygonRgn(Poly(LB), UB, ALTERNATE)
    
    'If the creation of the region was successful then colour it
    If hRgn Then FillRgn hdc, hRgn, hBrush
    
    DeleteObject hRgn
    
    'select in old brush, and delete current brush
    DeleteObject SelectObject(hdc, hOldPen)
End If

End Sub

Public Sub DrawPoly_NoOutline(Poly() As POINTAPI, hdc As Long, FillCol As Long)
Dim hBrush As Long, hOldPen As Long
Dim hRgn As Long
Dim LB As Integer, UB As Integer

LB = LBound(Poly)
UB = UBound(Poly)

'make a new brush
hBrush = CreatePen(PS_SOLID, 1, FillCol)
hOldPen = SelectObject(hdc, hBrush)

'Creates region to fill with colour
hRgn = CreatePolygonRgn(Poly(LB), UB, ALTERNATE)

'If the creation of the region was successful then colour it
If hRgn Then FillRgn hdc, hRgn, hBrush

DeleteObject hRgn
'select in old brush, and delete current brush
DeleteObject SelectObject(hdc, hOldPen)

End Sub

Public Sub HatchCircle(hdc As Long, lFillCol As Long, iSize As Integer, X As Single, Y As Single)
Dim hBrush As Long, hBrushDiag As Long, hOldPen As Long
Dim hRgn As Long, iSizeX2 As Long

hBrush = CreateHatchBrush(HS_CROSS, lFillCol)
hBrushDiag = CreateHatchBrush(HS_DIAGCROSS, lFillCol)

iSizeX2 = iSize * 2
hRgn = CreateRoundRectRgn(X - iSize, Y - iSize, X + iSize, Y + iSize, iSizeX2, iSizeX2)

'If the creation of the region was successful then colour it
If hRgn Then
    hOldPen = SelectObject(hdc, hBrush)
    FillRgn hdc, hRgn, hBrush
    
    DeleteObject SelectObject(hdc, hBrushDiag)
    FillRgn hdc, hRgn, hBrushDiag
End If


DeleteObject hRgn
DeleteObject hBrushDiag

End Sub

'Public Sub HatchPoly(Poly() As POINTAPI, hdc As Long, FillCol As Long)
'Dim hBrush As Long, hOldPen As Long
'Dim hRgn As Long
'Dim LB As Integer, UB As Integer
'
'LB = LBound(Poly)
'UB = UBound(Poly)
'
'hBrush = CreateHatchBrush(HS_DIAGCROSS, FillCol)
'hOldPen = SelectObject(hdc, hBrush)
'
''Creates region to fill with colour
'hRgn = CreatePolygonRgn(Poly(LB), UB, ALTERNATE)
'
''If the creation of the region was successful then colour it
'If hRgn Then FillRgn hdc, hRgn, hBrush
'
'DeleteObject hRgn
'DeleteObject SelectObject(hdc, hBrush)
'
'End Sub
