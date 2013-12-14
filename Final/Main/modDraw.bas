Attribute VB_Name = "modDraw"
Option Explicit

Private Const PS_SOLID = 0

Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function RectangleGDI Lib "gdi32" Alias "Rectangle" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

'---- circle
Private Declare Function Ellipse Lib "GDI32.dll" ( _
    ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SetDCBrushColor Lib "GDI32.dll" ( _
    ByVal hDC As Long, ByVal crColor As Long) As Long

Private Declare Function GetStockObject Lib "GDI32.dll" (ByVal nIndex As Long) As Long

Private Const DC_BRUSH As Long = 18



'http://www.bitwisemag.com/copy/vb/vb_graphics1.html

'http://www.vbexplorer.com/VBExplorer/gdi1.asp
'good explanations

'Line (x,y)-(x,y), BoxColour
'Circle (PowerUp.x, PowerUp.y), Powerup_Radius, vbRed

'note: api uses pixels (scale)

Public Sub DLine(ByRef X1 As Single, ByRef Y1 As Single, _
    ByRef X2 As Single, ByRef Y2 As Single, _
    ByRef Colour As Long, ByRef Width As Integer, _
    ByRef hDC As Long)

Dim Pen As Long

Pen = CreatePen(PS_SOLID, Width, Colour)
DeleteObject SelectObject(hDC, Pen)

MoveToEx hDC, X1, Y1, 0
LineTo hDC, X2, Y2

End Sub

Public Sub DRectangle(ByRef X1 As Single, ByRef Y1 As Single, _
                    ByRef X2 As Single, ByRef Y2 As Single, _
                    ByRef Width As Single, ByRef BorderColour As Long, _
                    ByRef hDC As Long) ', Optional ByRef FillColour As Long = -1)

Dim Pen As Long

Pen = CreatePen(PS_SOLID, Width, BorderColour)
DeleteObject SelectObject(hDC, Pen)

RectangleGDI hDC, X1, Y1, X2, Y2

End Sub

Public Sub DCircle(ByRef hDC As Long, ByRef X As Single, ByRef Y As Single, _
                ByRef Radius As Single, ByRef Colour As Long, _
                ByRef Width As Single)


Dim i As Long
Dim hOldBrush As Long

hOldBrush = SelectObject(hDC, GetStockObject(DC_BRUSH))
'DeleteObject hOldBrush

SetDCBrushColor hDC, Colour

Ellipse hDC, X, Y, X + Radius * 2, Y + Radius * 2

End Sub
